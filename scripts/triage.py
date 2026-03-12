"""
Email Triage CLI — authenticates via browser and writes a Graph open extension.

Usage:
    uv run scripts/triage.py                        # interactive prompts
    uv run scripts/triage.py --message-id <id>      # skip message ID prompt

Azure AD prerequisite
─────────────────────
In portal.azure.com → your app registration → Authentication:
  Add platform → Mobile and desktop applications
  Custom redirect URI: http://localhost
"""

import argparse
import json
import sys
import webbrowser

import msal
import requests

CLIENT_ID      = "f28629c6-2f87-4afc-a6ff-1cbbd50166af"
AUTHORITY      = "https://login.microsoftonline.com/common"
SCOPES         = ["Mail.ReadWrite"]
EXTENSION_NAME = "com.rpapub.emailtriage"
REDIRECT_URI   = "http://localhost"

COUNTRIES = [
    "AUS","AUT","BEL","BRA","CAN","CHE","CHN","DEU","DNK","EGY",
    "ESP","FIN","FRA","GBR","IDN","IND","ITA","JPN","KOR","MEX",
    "NGA","NLD","NOR","POL","PRT","RUS","SAU","SWE","TUR","USA","ZAF",
]


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_token() -> str:
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    # Try silent first (cached session)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            print("Using cached credentials.")
            return result["access_token"]

    # Interactive — opens browser
    print("Opening browser for sign-in…")
    result = app.acquire_token_interactive(scopes=SCOPES)

    if "access_token" not in result:
        print("Authentication failed:")
        print(result.get("error_description", result.get("error", "unknown")))
        sys.exit(1)

    print("Signed in successfully.")
    return result["access_token"]


# ── Graph ─────────────────────────────────────────────────────────────────────

def list_recent_messages(token: str, count: int = 10) -> list[dict]:
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/me/messages",
        headers={"Authorization": f"Bearer {token}"},
        params={"$top": count, "$select": "id,subject,from,receivedDateTime", "$orderby": "receivedDateTime desc"},
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


def write_extension(token: str, message_id: str, data: dict) -> None:
    url  = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/extensions"
    body = {"@odata.type": "microsoft.graph.openTypeExtension", "extensionName": EXTENSION_NAME, **data}
    hdrs = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    resp = requests.post(url, headers=hdrs, json=body)

    if resp.status_code == 409:
        resp = requests.patch(f"{url}/{EXTENSION_NAME}", headers=hdrs, json=body)
        if resp.status_code not in (200, 204):
            resp.raise_for_status()
        print("Extension updated.")
    elif resp.status_code == 201:
        print("Extension created.")
    else:
        resp.raise_for_status()


# ── Prompts ───────────────────────────────────────────────────────────────────

def prompt(label: str, default: str = "") -> str:
    hint = f" [{default}]" if default else ""
    val = input(f"  {label}{hint}: ").strip()
    return val or default


def pick_message(token: str) -> str:
    print("\nFetching recent messages…")
    messages = list_recent_messages(token)
    if not messages:
        print("No messages found.")
        sys.exit(1)

    print()
    for i, m in enumerate(messages, 1):
        sender = m.get("from", {}).get("emailAddress", {}).get("address", "?")
        print(f"  {i:2}. {m['subject'][:60]:<60}  <{sender}>")

    while True:
        raw = input("\nSelect message number: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(messages):
            return messages[int(raw) - 1]["id"]
        print("  Invalid selection.")


def pick_country() -> str:
    print("\n  Countries:", "  ".join(COUNTRIES))
    while True:
        val = input("  Country (3-letter code): ").strip().upper()
        if val in COUNTRIES:
            return val
        print(f"  Unknown code. Choose from the list above.")


def pick_year() -> str:
    from datetime import date
    current = date.today().year
    years   = [str(y) for y in range(current, current - 6, -1)]
    while True:
        val = input(f"  Year [{current}]: ").strip() or str(current)
        if val in years:
            return val
        print(f"  Valid years: {', '.join(years)}")


def pick_quarter() -> str:
    from datetime import date
    current = f"Q{(date.today().month - 1) // 3 + 1}"
    while True:
        val = input(f"  Quarter (Q1–Q4) [{current}]: ").strip().upper() or current
        if val in ("Q1", "Q2", "Q3", "Q4"):
            return val
        print("  Enter Q1, Q2, Q3, or Q4.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description="Email Triage CLI")
    parser.add_argument("--message-id", help="Graph message ID (skip interactive selection)")
    args = parser.parse_args()

    token = get_token()

    message_id = args.message_id or pick_message(token)

    print("\nTriage fields (press Enter to accept default):\n")
    data = {
        "reference": prompt("Reference"),
        "category":  prompt("Category"),
        "priority":  prompt("Priority"),
        "owner":     prompt("Owner"),
        "notes":     prompt("Notes"),
        "country":   pick_country(),
        "year":      pick_year(),
        "quarter":   pick_quarter(),
    }

    print("\nPayload:")
    print(json.dumps(data, indent=2))
    confirm = input("\nSave to message extension? [Y/n]: ").strip().lower()
    if confirm in ("", "y", "yes"):
        write_extension(token, message_id, data)
    else:
        print("Aborted.")


if __name__ == "__main__":
    main()
