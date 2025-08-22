#!/usr/bin/env python3
"""
Scrape Outlook mailboxes defined in a paths file and store emails to SQLite.

Usage:
    python macos/run_mac_scraper.py /absolute/path/to/mailbox_paths.txt [--only-account "Exact Account Name"]

The paths file format is:
    Name: <Account Name>  |  Stream: <STREAM>
      - <Mailbox/Path>
      - <Mailbox/Path>

The script automatically maps account names from the file to available Outlook
accounts (case-insensitive and partial matching) and skips any unavailable ones.
"""

import sys
import os
import datetime
import re
import argparse
from typing import List, Tuple

# Add the parent directory to sys.path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mac_outlook_client import (
    get_outlook_accounts,
    get_n_most_recent_emails,
    clean_email_content,
    clean_email_subject,
)
from utils.sqlite_storage import init_db, insert_emails_bulk, get_all_emails


def validate_date(date_str: str) -> bool:
    """
    Validate that the date is not in the future and not today.
    Returns True if date is valid, False otherwise.
    """
    try:
        day, month, year = date_str.split('-')
        target_date = datetime.datetime(int(year), int(month), int(day))
        today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        if target_date >= today:
            print(f"Error: Cannot scrape emails for today ({today.strftime('%d-%m-%Y')}) or future dates.")
            return False
        return True

    except ValueError as e:
        print(f"Error: Invalid date format. Please use DD-MM-YYYY format. Error: {e}")
        return False

def parse_mailbox_paths(file_path: str) -> List[Tuple[str, str, List[str]]]:
    """
    Parse a mailbox paths file into (account_name, stream, [mailbox_paths]).
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Mailbox paths file not found: {file_path}")

    entries: List[Tuple[str, str, List[str]]] = []
    current_name: str = ""
    current_stream: str = ""
    current_mailboxes: List[str] = []

    header_regex = re.compile(r"^Name:\s*(.*?)\s*\|\s*Stream:\s*([A-Za-z]+)\s*$")
    mailbox_regex = re.compile(r"^\s*-\s*(.+?)\s*$")

    with open(file_path, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.rstrip("\n")

            if not line.strip():
                # Blank line ends current block
                if current_name and current_stream and current_mailboxes:
                    entries.append((current_name, current_stream, current_mailboxes))
                current_name, current_stream, current_mailboxes = "", "", []
                continue

            header_match = header_regex.match(line)
            if header_match:
                # Flush previous if any
                if current_name and current_stream and current_mailboxes:
                    entries.append((current_name, current_stream, current_mailboxes))
                current_name = header_match.group(1).strip()
                current_stream = header_match.group(2).strip()
                current_mailboxes = []
                continue

            mailbox_match = mailbox_regex.match(line)
            if mailbox_match:
                mailbox_path = mailbox_match.group(1).strip()
                if mailbox_path:
                    current_mailboxes.append(mailbox_path)

        # End of file flush
        if current_name and current_stream and current_mailboxes:
            entries.append((current_name, current_stream, current_mailboxes))

    return entries


def main():
    parser = argparse.ArgumentParser(description="Scrape Outlook mailboxes defined in a paths file and store emails in SQLite.")
    parser.add_argument(
        "paths_file",
        help="Absolute path to the mailbox paths definition file.",
    )
    parser.add_argument(
        "--only-account",
        dest="only_account",
        help="If provided, only scrape this specific Outlook account (must match the Name in the paths file and Outlook account name).",
        default=None,
    )
    args = parser.parse_args()

    # Parse definitions
    try:
        mailbox_definitions = parse_mailbox_paths(args.paths_file)
    except Exception as e:
        print(f"Failed to parse mailbox paths file: {e}")
        return

    if args.only_account:
        mailbox_definitions = [d for d in mailbox_definitions if d[0] == args.only_account]
        if not mailbox_definitions:
            print(f"No definitions found for account: {args.only_account}")
            return

    # Discover available Outlook accounts and filter by accessibility
    discovered_accounts = get_outlook_accounts()
    available_accounts_set = set(discovered_accounts)
    if not available_accounts_set:
        print("Warning: No Outlook accounts discovered via AppleScript. Proceeding to attempt delegated access using names from the paths file.")

    # Initialize DB once
    init_db()

    total_new_overall = 0

    # Helper: resolve a file account name to an actual Outlook account name
    def resolve_account_name(requested: str, available: List[str]) -> str | None:
        # Exact case-sensitive
        if requested in available:
            return requested
        # Exact case-insensitive
        lower_map = {a.lower(): a for a in available}
        if requested.lower() in lower_map:
            return lower_map[requested.lower()]
        # Heuristic: remove common suffixes like " Artsci"
        simplified = requested.replace(" Artsci", "").strip()
        if simplified.lower() in lower_map:
            return lower_map[simplified.lower()]
        # Partial contains (case-insensitive)
        candidates = [a for a in available if requested.lower() in a.lower() or simplified.lower() in a.lower()]
        if len(candidates) == 1:
            return candidates[0]
        # Allow pass-through for delegated accounts not listed in discovered accounts
        return requested

    for account_name, stream, mailbox_paths in mailbox_definitions:
        resolved_account = resolve_account_name(account_name, discovered_accounts)
        if not resolved_account:
            print(f"Skipping '{account_name}' — could not map to an available Outlook account on this machine.")
            continue

        print(f"\nAccount: {resolved_account} | Stream: {stream}")

        for mailbox in mailbox_paths:
            print(f"  Scraping mailbox: {mailbox}")
            emails = get_n_most_recent_emails(resolved_account, mailbox, 10000)
            if not emails:
                print("    No emails found.")
                continue

            # Only keep emails up to yesterday
            today_str = datetime.datetime.now().strftime("%Y-%m-%d")
            filtered_emails = [e for e in emails if e.received < today_str]
            if not filtered_emails:
                print("    No emails up to yesterday.")
                continue

            email_tuples = [
                (
                    clean_email_subject(e.subject),
                    clean_email_content(e.content),
                    e.received,
                    stream,
                    e.person_name,
                )
                for e in filtered_emails
            ]

            before_count = len(get_all_emails())
            insert_emails_bulk(email_tuples)
            after_count = len(get_all_emails())
            added = after_count - before_count
            total_new_overall += added
            if added > 0:
                print(f"    Added {added} new emails.")
            else:
                print("    No new emails.")

    print(f"\nTotal new emails added across all accounts: {total_new_overall}")


if __name__ == "__main__":
    main()