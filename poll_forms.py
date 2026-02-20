#!/usr/bin/env python3
"""Poll Microsoft Forms and get notified when they go live.

Uses the OAuth2 device code flow for auth (no browser automation needed).
Access tokens refresh automatically; refresh tokens last 90 days rolling.
"""

import argparse
import json
import os
import platform
import subprocess
import sys
import time
from pathlib import Path

import httpx

# Microsoft Office public client (works across all M365 tenants, no secret needed)
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
FORMS_APP_ID = "c9a559d2-7aab-4f13-a6ed-e7e9c52aec87"

PROJECT_DIR = Path(__file__).parent
TOKEN_FILE = PROJECT_DIR / "forms_tokens.json"
FORMS_FILE = PROJECT_DIR / "forms.json"

MICROSOFT_LOGIN = "https://login.microsoftonline.com"
FORMS_API = "https://forms.office.com/formapi/api"


# ── Auth ─────────────────────────────────────────────────────────────────────

def device_code_auth(tenant: str):
    """Interactive device code login. User enters a code at microsoft.com/device."""
    resp = httpx.post(
        f"{MICROSOFT_LOGIN}/{tenant}/oauth2/v2.0/devicecode",
        data={"client_id": CLIENT_ID, "scope": f"{FORMS_APP_ID}/.default offline_access"},
    )
    resp.raise_for_status()
    result = resp.json()

    print(f"\n  Go to: {result['verification_uri']}")
    print(f"  Enter code: {result['user_code']}\n")
    print("  Waiting for sign-in...", end="", flush=True)

    interval = result.get("interval", 5)
    token_url = f"{MICROSOFT_LOGIN}/{tenant}/oauth2/v2.0/token"

    while True:
        time.sleep(interval)
        print(".", end="", flush=True)
        r = httpx.post(token_url, data={
            "client_id": CLIENT_ID,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": result["device_code"],
        })
        body = r.json()

        if r.status_code == 200:
            body["_obtained_at"] = int(time.time())
            body["_tenant"] = tenant
            _save_tokens(body)
            print("\n  Authenticated!")
            return

        error = body.get("error")
        if error == "authorization_pending":
            continue
        if error == "expired_token":
            print("\n  Code expired. Try again.")
            sys.exit(1)
        print(f"\n  Error: {body.get('error_description')}")
        sys.exit(1)


def _load_tokens() -> dict:
    try:
        with open(TOKEN_FILE) as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Not authenticated. Run: forms-watcher auth")
        sys.exit(1)


def _save_tokens(tokens: dict):
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f)


def _refresh_tokens(tokens: dict) -> dict:
    tenant = tokens.get("_tenant", "common")
    r = httpx.post(
        f"{MICROSOFT_LOGIN}/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": CLIENT_ID,
            "grant_type": "refresh_token",
            "refresh_token": tokens["refresh_token"],
            "scope": f"{FORMS_APP_ID}/.default offline_access",
        },
    )
    if r.status_code != 200:
        print(f"  Token refresh failed: {r.json().get('error_description', r.text)}")
        print(f"  Re-run: forms-watcher auth")
        sys.exit(1)
    new = r.json()
    new["_obtained_at"] = int(time.time())
    new["_tenant"] = tenant
    _save_tokens(new)
    return new


def _ensure_fresh(tokens: dict) -> dict:
    obtained = tokens.get("_obtained_at", 0)
    expires_in = tokens.get("expires_in", 0)
    if time.time() > obtained + expires_in - 300:
        print("  [refreshing token...]", flush=True)
        return _refresh_tokens(tokens)
    return tokens


# ── Forms config ─────────────────────────────────────────────────────────────

def _load_forms() -> list[dict]:
    """Load watched forms from forms.json."""
    try:
        with open(FORMS_FILE) as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"No forms configured. Run: forms-watcher add <url> [<url> ...]")
        sys.exit(1)


def _save_forms(forms: list[dict]):
    with open(FORMS_FILE, "w") as f:
        json.dump(forms, f, indent=2)


def _resolve_form_url(url: str) -> dict:
    """Follow a forms.office.com/r/xxx short URL to extract the full form ID."""
    r = httpx.get(url, follow_redirects=True)
    final = str(r.url)
    # Extract the id= parameter from the resolved URL
    from urllib.parse import parse_qs, urlparse
    parsed = urlparse(final)
    params = parse_qs(parsed.query)
    form_id = params.get("id", [None])[0]
    if not form_id:
        print(f"  Could not resolve form URL: {url}")
        sys.exit(1)
    # Extract short code from original URL
    short = url.rstrip("/").split("/")[-1]
    return {"short": short, "form_id": form_id, "url": url}


# ── Notification ─────────────────────────────────────────────────────────────

def _notify(message: str):
    system = platform.system()
    if system == "Darwin":
        subprocess.run(["say", message])
    elif system == "Linux":
        subprocess.run(["notify-send", "Forms Watcher", message], check=False)
    else:
        # Windows or unknown - just print loud
        print(f"\a  *** {message} ***")


# ── Polling ──────────────────────────────────────────────────────────────────

def _check_form(client: httpx.Client, tenant: str, group: str, form_id: str) -> tuple[bool, str]:
    url = f"{FORMS_API}/{tenant}/groups/{group}/light/runtimeFormsWithResponses('{form_id}')?$expand=questions($expand=choices)&$top=1"
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


def poll(interval: int):
    forms = _load_forms()
    tokens = _load_tokens()
    tenant = tokens.get("_tenant", "common")
    notified: set[str] = set()

    # Discover group ID from first form (all forms in same tenant share it)
    # TODO: this is hardcoded for now, will be resolved in tenant refactor
    group = "5385ae13-9f9d-4598-a665-dc861def3047"

    print(f"Polling {len(forms)} forms every {interval}s")
    for f in forms:
        print(f"  - {f['url']}")
    print(flush=True)

    with httpx.Client(timeout=10) as client:
        while True:
            tokens = _ensure_fresh(tokens)
            client.headers["Authorization"] = f"Bearer {tokens['access_token']}"
            ts = time.strftime("%H:%M:%S")

            for form in forms:
                fid = form["form_id"]
                if fid in notified:
                    continue
                is_open, detail = _check_form(client, tenant, group, fid)
                label = form.get("name", form["short"])
                print(f"  [{ts}] {label}: {detail}", flush=True)
                if is_open:
                    _notify(f"OPEN: {label}")
                    notified.add(fid)
                    print(f"  >>> {label} is OPEN! <<<", flush=True)

            remaining = len(forms) - len(notified)
            if remaining == 0:
                print("\nAll forms open. Done.")
                break

            print(f"  --- {remaining} closed, next in {interval}s ---\n", flush=True)
            time.sleep(interval)


# ── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        prog="forms-watcher",
        description="Get notified when Microsoft Forms go live.",
    )
    sub = parser.add_subparsers(dest="command")

    auth_p = sub.add_parser("auth", help="Authenticate with your Microsoft account")
    auth_p.add_argument("--tenant", default="common", help="Azure AD tenant ID (default: common)")

    add_p = sub.add_parser("add", help="Add form URLs to watch")
    add_p.add_argument("urls", nargs="+", help="Form URLs (forms.office.com/r/xxx)")
    add_p.add_argument("--name", action="append", help="Label for each form")

    sub.add_parser("list", help="Show watched forms")
    sub.add_parser("clear", help="Remove all watched forms")

    poll_p = sub.add_parser("poll", help="Start polling (default)")
    poll_p.add_argument("--interval", type=int, default=5, help="Seconds between checks (default: 5)")

    args = parser.parse_args()

    if args.command == "auth":
        device_code_auth(args.tenant)

    elif args.command == "add":
        existing = []
        if FORMS_FILE.exists():
            with open(FORMS_FILE) as f:
                existing = json.load(f)
        existing_urls = {f["url"] for f in existing}
        for i, url in enumerate(args.urls):
            if url in existing_urls:
                print(f"  Already watching: {url}")
                continue
            print(f"  Resolving: {url}...", end=" ", flush=True)
            form = _resolve_form_url(url)
            if args.name and i < len(args.name):
                form["name"] = args.name[i]
            existing.append(form)
            print(f"OK ({form['short']})")
        _save_forms(existing)
        print(f"\n  Watching {len(existing)} forms.")

    elif args.command == "list":
        if not FORMS_FILE.exists():
            print("No forms configured.")
            return
        forms = _load_forms()
        for f in forms:
            label = f.get("name", f["short"])
            print(f"  {label}: {f['url']}")

    elif args.command == "clear":
        if FORMS_FILE.exists():
            FORMS_FILE.unlink()
        print("  Cleared all watched forms.")

    elif args.command == "poll" or args.command is None:
        interval = getattr(args, "interval", 5)
        poll(interval)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
