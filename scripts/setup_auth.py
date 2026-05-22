#!/usr/bin/env python3
"""One-time auth bootstrap for the Outlook / Microsoft 365 skill.

Runs the OAuth 2.0 device-code flow against login.microsoftonline.com and
writes ../.credentials/credentials.json with the access + refresh tokens that
outlook.py then maintains autonomously.

Prerequisites:
  1. An Azure app registration (personal Microsoft accounts supported) with
     the delegated Graph scopes below + "Allow public client flows" enabled.
  2. ../.credentials/config.json containing at least {"client_id": "..."}.

Usage:
  python setup_auth.py
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

import requests

SKILL_DIR = Path(__file__).resolve().parent.parent
CRED_DIR = SKILL_DIR / ".credentials"
CRED_FILE = CRED_DIR / "credentials.json"
CONF_FILE = CRED_DIR / "config.json"

AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0"
DEVICECODE_URL = f"{AUTHORITY}/devicecode"
TOKEN_URL = f"{AUTHORITY}/token"

SCOPE = (
    "offline_access User.Read "
    "Calendars.ReadWrite Calendars.ReadWrite.Shared "
    "Tasks.ReadWrite Tasks.ReadWrite.Shared"
)


def _load_client_id() -> str:
    if not CONF_FILE.exists():
        sys.exit(
            f"missing {CONF_FILE}\n"
            f"Copy .credentials/config.example.json to config.json and fill in "
            f"your Azure app client_id."
        )
    conf = json.loads(CONF_FILE.read_text())
    client_id = conf.get("client_id")
    if not client_id or client_id.startswith("your-"):
        sys.exit(f"set a real client_id in {CONF_FILE}")
    return client_id


def _request_device_code(client_id: str) -> dict:
    r = requests.post(
        DEVICECODE_URL,
        data={"client_id": client_id, "scope": SCOPE},
        timeout=20,
    )
    if r.status_code >= 400:
        sys.exit(f"devicecode request failed {r.status_code}: {r.text[:500]}")
    return r.json()


def _poll_for_token(client_id: str, device_code: str, interval: int,
                    expires_in: int) -> dict:
    deadline = time.time() + expires_in
    while time.time() < deadline:
        time.sleep(interval)
        r = requests.post(
            TOKEN_URL,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": client_id,
                "device_code": device_code,
            },
            timeout=20,
        )
        if r.status_code == 200:
            return r.json()
        err = r.json().get("error", "")
        if err == "authorization_pending":
            continue
        if err == "slow_down":
            interval += 5
            continue
        # authorization_declined, expired_token, bad_verification_code, ...
        sys.exit(f"auth failed: {err} — {r.json().get('error_description', '')[:300]}")
    sys.exit("device code expired before authorization completed")


def _save_creds(tok: dict) -> None:
    CRED_DIR.mkdir(parents=True, exist_ok=True)
    creds = {
        "access_token": tok["access_token"],
        "refresh_token": tok["refresh_token"],
        "expires_in": tok["expires_in"],
        "expires_at": int(time.time()) + int(tok["expires_in"]) - 60,
        "scope": tok.get("scope", SCOPE),
    }
    tmp = CRED_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(creds, indent=2))
    os.chmod(tmp, 0o600)
    tmp.replace(CRED_FILE)


def main() -> None:
    client_id = _load_client_id()
    dc = _request_device_code(client_id)
    print(dc["message"], flush=True)  # human-readable: go to URL, enter code
    tok = _poll_for_token(
        client_id, dc["device_code"], int(dc.get("interval", 5)),
        int(dc.get("expires_in", 900)),
    )
    if "refresh_token" not in tok:
        sys.exit(
            "no refresh_token in response — make sure 'offline_access' is in "
            "the requested scopes and granted on the app registration."
        )
    _save_creds(tok)
    print(f"\n✓ credentials written to {CRED_FILE}")
    print("  You can now run: python scripts/outlook.py status")


if __name__ == "__main__":
    main()
