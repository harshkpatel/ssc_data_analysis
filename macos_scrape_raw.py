#!/usr/bin/env python3
"""
Scrape raw Outlook emails without filtering or cleaning, saving directly to CSV.

Usage:
    python macos_scrape_raw.py /absolute/path/to/mailbox_paths.txt [--only-account "Exact Account Name"]

The paths file format is:
    Name: <Account Name>  |  Stream: <STREAM>
      - <Mailbox/Path>
      - <Mailbox/Path>

This script creates raw, unfiltered CSV files for comparison purposes.
"""

import sys
import os
import datetime
import re
import argparse
import csv
import subprocess
from typing import List, Tuple, Optional


def run_applescript(script: str) -> str:
    """Execute AppleScript and return text output or an error signature."""
    process = subprocess.Popen(
        ['osascript', '-e', script],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True # Decode stdout/stderr as text
    )
    stdout, stderr = process.communicate()
    if process.returncode != 0:
        print(f"Error running AppleScript (return code {process.returncode}):\nSTDERR: {stderr.strip()}")
        return f"OSASCRIPT_ERROR: {stderr.strip()}"
    return stdout.strip()


def get_outlook_accounts() -> List[str]:
    """Return a list of Outlook account names on this Mac."""
    script = '''
    tell application "Microsoft Outlook"
        set accountList to {}
        try
            repeat with acct in exchange accounts
                set end of accountList to name of acct
            end repeat
        on error errMsg number errNum
            -- Using log can be problematic with osascript -e if not handled well,
            -- returning an error string might be better for Python to catch.
            -- log "No Exchange accounts or error: " & errMsg 
        end try
        return accountList
    end tell
    '''
    result = run_applescript(script)
    if not result or result.startswith("OSASCRIPT_ERROR:") or result.startswith("APPLE_SCRIPT_ERROR:"):
        print(f"Failed to get Outlook accounts: {result}")
        return []

    accounts = [account.strip() for account in result.split(',') if account.strip()]
    return accounts


def get_folder_navigation_applescript(mailbox_path: str, account_name: str) -> str:
    """Return AppleScript lines to select the folder path (e.g., 'Inbox/Sub') for the active account variable 'acct'."""
    if '/' not in mailbox_path:
        return f'''
        set mb to missing value
        set folderCount to count of mail folders of acct
        repeat with i from 1 to folderCount
            try
                set currentFolder to mail folder i of acct
                set currentName to name of currentFolder
                -- Trim trailing spaces for comparison
                set trimmedName to do shell script "echo " & quoted form of currentName & " | sed 's/[[:space:]]*$//'"
                if trimmedName is "{mailbox_path}" then
                    set mb to currentFolder
                    exit repeat
                end if
            on error
                -- Skip this folder if we can't access it
            end try
        end repeat
        if mb is missing value then
            error "Folder '{mailbox_path}' not found in account '{account_name}'"
        end if'''

    path_parts = mailbox_path.split('/')
    script_lines = []
    script_lines.append(f'''
    set mb to missing value
    set folderCount to count of mail folders of acct
    repeat with i from 1 to folderCount
        try
            set currentFolder to mail folder i of acct
            set currentName to name of currentFolder
            -- Trim trailing spaces for comparison
            set trimmedName to do shell script "echo " & quoted form of currentName & " | sed 's/[[:space:]]*$//'"
            if trimmedName is "{path_parts[0]}" then
                set mb to currentFolder
                exit repeat
            end if
        on error
            -- Skip this folder if we can't access it
        end try
    end repeat
    if mb is missing value then
        error "Top folder '{path_parts[0]}' not found in account '{account_name}'"
    end if''')

    for i in range(1, len(path_parts)):
        subfolder_name = path_parts[i]
        script_lines.append(f'''
        set subMb to missing value
        set subFolderCount to count of mail folders of mb
        repeat with j from 1 to subFolderCount
            try
                set currentSubFolder to mail folder j of mb
                set currentSubName to name of currentSubFolder
                -- Trim trailing spaces for comparison
                set trimmedSubName to do shell script "echo " & quoted form of currentSubName & " | sed 's/[[:space:]]*$//'"
                if trimmedSubName is "{subfolder_name}" then
                    set subMb to currentSubFolder
                    exit repeat
                end if
            on error
                -- Skip this subfolder if we can't access it
            end try
        end repeat
        if subMb is missing value then
            error "Subfolder '{subfolder_name}' not found in '{path_parts[0]}'"
        end if
        set mb to subMb''')

    return '\n        '.join(script_lines)


def is_meeting_or_booking_email(subject: str, content: str) -> bool:
    """Return True if the email looks like a meeting or Microsoft Bookings notification."""
    # Common patterns in Teams meeting emails
    teams_patterns = [
        r'Teams Meeting',
        r'Microsoft Teams',
        r'Teams meeting',
        r'teams\.microsoft\.com',
        r'Join Microsoft Teams Meeting',
        r'Meeting Details',
        r'Calendar Event',
        r'Meeting Invitation',
        r'Teams Video Call',
        r'Teams Audio Call',
        r'Teams Conference',
        r'Teams Webinar'
    ]

    # Common patterns in booking emails (Microsoft Bookings / one-on-ones)
    booking_patterns = [
        r'^New booking',
        r'New booking',
        r'Updated booking',
        r'Canceled:',
        r'Cancelled:',
        r'Canceled\s+',
        r'Cancelled\s+',
        r'Microsoft Bookings',
        r'Bookings',
        r'Booking Confirmation',
        r'Your booking is confirmed',
        r'Appointment Confirmed',
        r'Join your appointment',
        r'Reschedule',
        r'Cancel or reschedule',
        r'Meeting Confirmation',
        r'Calendar Invitation',
        r'Event Details',
        r'Meeting Details',
        r'Invitation to',
        r'has invited you to',
        r'One on One',
        r'One-on-One',
        r'Calendar Event',
        r'Outlook Calendar',
        r'Calendar Reminder',
        r'Event Reminder',
        r'Meeting Reminder',
        r'Appointment Reminder'
    ]

    # Check subject and content for patterns
    all_patterns = teams_patterns + booking_patterns
    
    # Be more lenient - only filter if we're very confident it's a meeting/booking
    subject_matches = 0
    content_matches = 0
    
    for pattern in all_patterns:
        if re.search(pattern, subject, re.IGNORECASE):
            subject_matches += 1
        if re.search(pattern, content, re.IGNORECASE):
            content_matches += 1
    
    # Only filter if we have multiple strong matches or a very clear subject match
    if subject_matches >= 2 or (subject_matches >= 1 and content_matches >= 1):
        return True
    
    # Check for very specific meeting indicators
    if re.search(r'teams\.microsoft\.com', content, re.IGNORECASE):
        return True
    
    if re.search(r'Join.*Meeting', subject, re.IGNORECASE):
        return True
    
    return False


def get_raw_emails(account_name: str, mailbox_name: str, n: int) -> List[Tuple[str, str, str, str]]:
    """Return up to n recent raw emails for a given account and mailbox path without cleaning."""
    script = f'''
    tell application "Microsoft Outlook"
        try
            -- Select the requested exchange account directly by name
            set acct to (first exchange account whose name is "{account_name}")

            -- Navigate to the target mailbox using the selected account
            {get_folder_navigation_applescript(mailbox_name, account_name)}

            set msgs to (messages of mb)
            if (count of msgs) is 0 then
                return ""
            end if

            set emailList to {{}}
            set currentCount to 0
            repeat with msg in msgs
                if currentCount is {n} then
                    exit repeat
                end if
                
                set msgID to id of msg
                set msgSubject to subject of msg
                set msgContent to plain text content of msg
                set msgTime to time received of msg

                -- Handle null or missing content gracefully
                if msgContent is missing value then
                    set msgContent to ""
                end if
                
                if msgSubject is missing value then
                    set msgSubject to ""
                end if

                set msgYear to year of msgTime as string
                set msgMonth to (month of msgTime as integer) as string
                if (count of msgMonth) is 1 then set msgMonth to "0" & msgMonth
                set msgDay to day of msgTime as string
                if (count of msgDay) is 1 then set msgDay to "0" & msgDay
                set dateOnly to msgYear & "-" & msgMonth & "-" & msgDay

                set msgInfo to msgID & "|||DELIM|||" & msgSubject & "|||DELIM|||" & msgContent & "|||DELIM|||" & dateOnly
                set end of emailList to msgInfo
                set currentCount to currentCount + 1
            end repeat

            set emailText to ""
            repeat with i from 1 to count of emailList
                set emailText to emailText & item i of emailList
                if i < (count of emailList) then set emailText to emailText & "|||EMAIL|||"
            end repeat
            return emailText
            
        on error errMsg number errNum
            log "Error in get_raw_emails: " & errMsg & " (" & errNum & ")"
            return "ERROR_SCRIPT: " & errMsg & " (" & errNum & ")"
        end try
    end tell
    '''

    result = run_applescript(script)
    if not result:
        return []
    
    # Check for error responses
    if result.startswith("ERROR_"):
        print(f"    AppleScript error: {result}")
        return []

    emails = []
    for line in result.split('|||EMAIL|||'):
        if "|||DELIM|||" in line:
            parts = line.split("|||DELIM|||", 3)
            if len(parts) >= 4:
                msg_id = parts[0].strip()
                subject = parts[1].strip()
                content = parts[2].strip()
                received = parts[3].strip()

                # Skip meeting/booking emails since they aren't real emails
                if is_meeting_or_booking_email(subject, content):
                    continue

                # Keep all other emails, even empty ones, for raw comparison
                emails.append((msg_id, subject, content, received))

    return emails


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


def export_to_csv(emails: List[Tuple[str, str, str, str, str, str]], filename: str):
    """Export emails to CSV file."""
    # Ensure csv_files directory exists
    csv_dir = 'csv_files'
    if not os.path.exists(csv_dir):
        os.makedirs(csv_dir)
    
    filepath = os.path.join(csv_dir, filename)
    with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['message_id', 'subject', 'content', 'received', 'stream', 'person_name'])
        for msg_id, subject, content, received, stream, person_name in emails:
            writer.writerow([msg_id, subject, content, received, stream, person_name])
    
    print(f"Exported {len(emails)} raw emails to {filepath}")


def main():
    parser = argparse.ArgumentParser(description="Scrape raw Outlook emails without filtering and save to CSV.")
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
    parser.add_argument(
        "--account-name",
        dest="account_name",
        help="Your name for the CSV filename (e.g., 'harsh' for harsh_raw_emails.csv).",
        default=None,
    )
    args = parser.parse_args()

    # Get account name for filename
    account_name = args.account_name
    if not account_name:
        account_name = input("Enter your name (for the CSV filename): ").strip()
    account_clean = account_name.replace(' ', '-').replace('/', '--')

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

    total_emails = 0
    all_raw_emails = []

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
            emails = get_raw_emails(resolved_account, mailbox, 10000)
            if not emails:
                print("    No emails found.")
                continue

            # Keep all emails for raw comparison (no filtering)
            print(f"    Found {len(emails)} raw emails.")
            total_emails += len(emails)
            
            # Add stream information and person name to emails
            person_name = account_name.split()[0] if account_name else ""
            for msg_id, subject, content, received in emails:
                all_raw_emails.append((msg_id, subject, content, received, stream, person_name))

    if all_raw_emails:
        # Export all raw emails to CSV
        filename = f"{account_clean}_raw_emails.csv"
        export_to_csv(all_raw_emails, filename)
        print(f"\nTotal raw emails exported: {total_emails}")
    else:
        print("\nNo emails found to export.")


if __name__ == "__main__":
    main()
