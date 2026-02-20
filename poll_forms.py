#!/usr/bin/env python3
"""Poll Microsoft Forms and notify when one goes live.

Auth: Device code flow via Microsoft Office public client.
Tokens auto-refresh (~1hr access, 90-day rolling refresh).
One-time setup: run with --auth flag.
"""

import json
import os
import subprocess
import sys
import time

import httpx

# --- Config ---
TENANT = "cb72c54e-4a31-4d9e-b14a-1ea36dfac94c"
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Microsoft Office (public)
FORMS_APP = "c9a559d2-7aab-4f13-a6ed-e7e9c52aec87"
FORMS_GROUP = "5385ae13-9f9d-4598-a665-dc861def3047"
TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "forms_tokens.json")
API_BASE = f"https://forms.office.com/formapi/api/{TENANT}/groups/{FORMS_GROUP}/light/runtimeFormsWithResponses"
POLL_INTERVAL = 5

FORMS = {
    "Transform": "k8MLfL0MtK",
    "Pivot": "WSaa1rckHR",
    "Describe": "qRvV95Yfgb",
    "Merging/Join": "tra20NMpXr",
    "Groupby": "QiazHAjwuW",
}

FORM_IDS = {
    "k8MLfL0MtK": "TsVyyzFKnk2xSh6jbfrJTBOuhVOdn5hFpmXchh3vMEdUMlI4WlZQTVQ5SlNCTEMxRlZLRFVaWTMzVSQlQCN0PWcu",
    "WSaa1rckHR": "TsVyyzFKnk2xSh6jbfrJTBOuhVOdn5hFpmXchh3vMEdUN01DQTNRWDlaS1pDUkE3WFROTFVLR0haOCQlQCN0PWcu",
    "qRvV95Yfgb": "TsVyyzFKnk2xSh6jbfrJTBOuhVOdn5hFpmXchh3vMEdUNVVVVjVVUlhYUUpaR09QMTJZUkRWS0RRMCQlQCN0PWcu",
    "tra20NMpXr": "TsVyyzFKnk2xSh6jbfrJTBOuhVOdn5hFpmXchh3vMEdUQloxRkREWEoyT05aTlZPNlRaMjI0QkM3OCQlQCN0PWcu",
    "QiazHAjwuW": "TsVyyzFKnk2xSh6jbfrJTBOuhVOdn5hFpmXchh3vMEdURFJUUVdPVkkxWllFVDhMWk1XTDE0S1U4MiQlQCN0PWcu",
}


# --- Auth ---
def device_code_auth():
    """One-time interactive auth via microsoft.com/device."""
    resp = httpx.post(
        f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/devicecode",
        data={"client_id": CLIENT_ID, "scope": f"{FORMS_APP}/.default offline_access"},
    )
    result = resp.json()

    print(f"\n  Go to: {result['verification_uri']}")
    print(f"  Enter code: {result['user_code']}\n")
    print("  Waiting for you to sign in...", end="", flush=True)

    interval = result.get("interval", 5)
    token_url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"

    while True:
        time.sleep(interval)
        print(".", end="", flush=True)
        resp = httpx.post(token_url, data={
            "client_id": CLIENT_ID,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": result["device_code"],
        })
        body = resp.json()

        if resp.status_code == 200:
            body["_obtained_at"] = int(time.time())
            save_tokens(body)
            print("\n  Authenticated! Tokens saved.")
            return body

        if body.get("error") == "authorization_pending":
            continue
        if body.get("error") == "expired_token":
            print("\n  Code expired. Run --auth again.")
            sys.exit(1)
        print(f"\n  Error: {body.get('error_description')}")
        sys.exit(1)


def load_tokens() -> dict:
    try:
        with open(TOKEN_FILE) as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"No tokens found. Run: python3 {sys.argv[0]} --auth")
        sys.exit(1)


def save_tokens(tokens: dict):
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f)


def refresh_access_token(tokens: dict) -> dict:
    """Use refresh token to get a new access token."""
    resp = httpx.post(
        f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID,
            "grant_type": "refresh_token",
            "refresh_token": tokens["refresh_token"],
            "scope": f"{FORMS_APP}/.default offline_access",
        },
    )
    new_tokens = resp.json()
    new_tokens["_obtained_at"] = int(time.time())
    save_tokens(new_tokens)
    return new_tokens


def ensure_fresh_token(tokens: dict) -> dict:
    """Refresh if access token is within 5 min of expiry."""
    obtained = tokens.get("_obtained_at", 0)
    expires_in = tokens.get("expires_in", 0)
    if time.time() > obtained + expires_in - 300:
        print("  [refreshing token...]", flush=True)
        return refresh_access_token(tokens)
    return tokens


# --- Polling ---
def notify(title: str):
    subprocess.run(["say", title])


def check_form(client: httpx.Client, form_id: str) -> tuple[bool, str]:
    url = f"{API_BASE}('{form_id}')?$expand=questions($expand=choices)&$top=1"
    try:
        resp = client.get(url)
        if resp.status_code == 200:
            return True, "OPEN"
        body = resp.json()
        code = body.get("error", {}).get("code", "?")
        if code == "5000":
            return False, "closed"
        if code == "5001":
            return False, "already submitted"
        return False, f"error {code}: {body.get('error', {}).get('message', '?')}"
    except httpx.TimeoutException:
        return False, "timeout"
    except Exception as e:
        return False, f"error: {e}"


def main():
    if "--auth" in sys.argv:
        device_code_auth()
        return

    tokens = load_tokens()
    notified = set()
    print(f"Polling {len(FORMS)} forms every {POLL_INTERVAL}s")
    print(f"Forms: {', '.join(FORMS.keys())}\n", flush=True)

    while True:
        tokens = ensure_fresh_token(tokens)
        ts = time.strftime("%H:%M:%S")

        with httpx.Client(
            headers={"Authorization": f"Bearer {tokens['access_token']}"},
            timeout=10,
        ) as client:
            for name, short in FORMS.items():
                if name in notified:
                    continue
                is_open, detail = check_form(client, FORM_IDS[short])
                print(f"  [{ts}] {name}: {detail}", flush=True)
                if is_open:
                    notify(f"OPEN: {name}")
                    notified.add(name)
                    print(f"  >>> {name} is OPEN! <<<", flush=True)

        remaining = len(FORMS) - len(notified)
        if remaining == 0:
            print("\nAll forms open. Done.")
            break

        print(f"  --- {remaining} closed, next check in {POLL_INTERVAL}s ---\n", flush=True)
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
