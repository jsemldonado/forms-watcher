#!/usr/bin/env python3
"""Poll Microsoft Forms and get notified when they go live.

Uses the OAuth2 device code flow for auth (no browser automation needed).
Access tokens refresh automatically; refresh tokens last 90 days rolling.
"""

import argparse
import json
import platform
import re
import subprocess
import sys
import time
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import httpx

# Microsoft Office public client (works across all M365 tenants, no secret needed)
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
FORMS_APP_ID = "c9a559d2-7aab-4f13-a6ed-e7e9c52aec87"

PROJECT_DIR = Path(__file__).parent
TOKEN_FILE = PROJECT_DIR / "forms_tokens.json"
FORMS_FILE = PROJECT_DIR / "forms.json"

MICROSOFT_LOGIN = "https://login.microsoftonline.com"
FORMS_BASE = "https://forms.office.com"


# ── Auth ─────────────────────────────────────────────────────────────────────

def _device_code_auth(tenant: str):
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
        print("Not authenticated. Run: forms-watcher auth")
        sys.exit(1)


def _save_tokens(tokens: dict):
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f)
    TOKEN_FILE.chmod(0o600)


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
        print("  Re-run: forms-watcher auth")
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


# ── Form discovery ───────────────────────────────────────────────────────────

def _resolve_form(url: str, access_token: str) -> dict:
    """Resolve a form URL to its full metadata (form_id, tenant, group).

    Uses ResponsePageStartup.ashx which returns prefetchFormUrl containing
    the tenant and group IDs in the API path.
    """
    # Follow short URL to get full form ID
    r = httpx.get(url, follow_redirects=True)
    final = str(r.url)
    params = parse_qs(urlparse(final).query)
    form_id = params.get("id", [None])[0]
    if not form_id:
        print(f"  Could not resolve: {url}")
        sys.exit(1)

    short = url.rstrip("/").split("/")[-1]

    # Hit startup handler to discover tenant + group from prefetchFormUrl
    startup_url = f"{FORMS_BASE}/handlers/ResponsePageStartup.ashx?id={form_id}&route=shorturl&mobile=false"
    r = httpx.get(startup_url, headers={"Authorization": f"Bearer {access_token}"}, timeout=10)
    data = r.json()

    prefetch = data.get("serverInfo", {}).get("prefetchFormUrl", "")
    # Pattern: /formapi/api/{tenant}/groups/{group}/...
    m = re.search(r"/formapi/api/([^/]+)/groups/([^/]+)/", prefetch)
    if not m:
        print(f"  Could not discover tenant/group for: {url}")
        print(f"  prefetchFormUrl: {prefetch}")
        sys.exit(1)

    return {
        "url": url,
        "short": short,
        "form_id": form_id,
        "tenant": m.group(1),
        "group": m.group(2),
    }


def _load_forms() -> list[dict]:
    try:
        with open(FORMS_FILE) as f:
            return json.load(f)
    except FileNotFoundError:
        print("No forms configured. Run: forms-watcher add <url> [<url> ...]")
        sys.exit(1)


def _save_forms(forms: list[dict]):
    with open(FORMS_FILE, "w") as f:
        json.dump(forms, f, indent=2)


# ── Notification ─────────────────────────────────────────────────────────────

def _notify(message: str):
    system = platform.system()
    if system == "Darwin":
        subprocess.Popen(["say", message])
    elif system == "Linux":
        subprocess.run(["notify-send", "Forms Watcher", message], check=False)
    else:
        print(f"\a  *** {message} ***")


# ── Polling ──────────────────────────────────────────────────────────────────

def _check_form(client: httpx.Client, form: dict) -> tuple[bool, str]:
    tenant = form["tenant"]
    group = form["group"]
    fid = form["form_id"]
    url = f"{FORMS_BASE}/formapi/api/{tenant}/groups/{group}/light/runtimeFormsWithResponses('{fid}')"
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


def _poll(interval: int):
    forms = _load_forms()
    tokens = _load_tokens()
    notified: set[str] = set()

    print(f"Polling {len(forms)} forms every {interval}s")
    for f in forms:
        label = f.get("name", f["short"])
        print(f"  - {label} ({f['url']})")
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
                is_open, detail = _check_form(client, form)
                label = form.get("name", form["short"])
                print(f"  [{ts}] {label}: {detail}", flush=True)
                if is_open:
                    _notify(f"OPEN: {label}")
                    notified.add(fid)
                    print(f"  >>> {label} is OPEN! <<<", flush=True)
                elif detail == "already submitted":
                    notified.add(fid)
                    print(f"  (skipping {label} from now on)", flush=True)

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

    sub.add_parser("auth", help="Authenticate with your Microsoft account")

    add_p = sub.add_parser("add", help="Add form URLs to watch")
    add_p.add_argument("urls", nargs="+", help="Form URLs (forms.office.com/r/xxx)")
    add_p.add_argument("--name", action="append", help="Label for each form (prompted if omitted)")

    rm_p = sub.add_parser("remove", help="Remove a form by name or URL")
    rm_p.add_argument("target", help="Form name, short code, or URL")

    sub.add_parser("list", help="Show watched forms")
    sub.add_parser("clear", help="Remove all watched forms")

    poll_p = sub.add_parser("poll", help="Start polling (default)")
    poll_p.add_argument("--interval", type=int, default=5, help="Seconds between checks (default: 5)")

    args = parser.parse_args()

    if args.command == "auth":
        _device_code_auth("common")

    elif args.command == "add":
        tokens = _load_tokens()
        tokens = _ensure_fresh(tokens)
        existing = json.loads(FORMS_FILE.read_text()) if FORMS_FILE.exists() else []
        existing_urls = {f["url"] for f in existing}
        interactive = not args.name
        for i, url in enumerate(args.urls):
            if url in existing_urls:
                print(f"  Already watching: {url}")
                continue
            print(f"  Resolving: {url}...", end=" ", flush=True)
            form = _resolve_form(url, tokens["access_token"])
            print(f"OK (tenant={form['tenant'][:8]}... group={form['group'][:8]}...)")
            if args.name and i < len(args.name):
                form["name"] = args.name[i]
            elif interactive:
                name = input(f"  Name for {form['short']} (enter to skip): ").strip()
                if name:
                    form["name"] = name
            existing.append(form)
        _save_forms(existing)
        print(f"\n  Watching {len(existing)} forms.")

    elif args.command == "remove":
        if not FORMS_FILE.exists():
            print("No forms configured.")
            return
        forms = _load_forms()
        target = args.target
        before = len(forms)
        forms = [f for f in forms if target not in (f.get("name"), f.get("short"), f.get("url"))]
        if len(forms) == before:
            print(f"  No form matching: {target}")
            return
        _save_forms(forms)
        print(f"  Removed. {len(forms)} forms remaining.")

    elif args.command == "list":
        if not FORMS_FILE.exists():
            print("No forms configured.")
            return
        for f in _load_forms():
            label = f.get("name", f["short"])
            print(f"  {label}: {f['url']}")

    elif args.command == "clear":
        if FORMS_FILE.exists():
            FORMS_FILE.unlink()
        print("  Cleared all watched forms.")

    elif args.command == "poll" or args.command is None:
        _poll(getattr(args, "interval", 5))

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
