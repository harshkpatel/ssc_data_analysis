#!/usr/bin/env python3
"""
Helpers for interacting with Outlook on macOS via AppleScript and cleaning email content.
"""

import subprocess
import re
from typing import List, Optional
from models.common_models import Email
try:
    from email_reply_parser import EmailReplyParser
    HAS_EMAIL_REPLY_PARSER = True
except Exception:
    HAS_EMAIL_REPLY_PARSER = False

def clean_email_subject(subject: str) -> str:
    """Return trimmed subject or empty string if missing."""
    if not subject:
        return ""
    return subject.strip()


def clean_email_content(content: str) -> str:
    """Return cleaned body text: strip quoted replies, HTML, warnings, and separators."""
    if not content:
        return ""

    content = parse_visible_reply_text(content)

    # Clean HTML Tags
    content = re.sub(r'<[^>]*>', '', content)
    content = re.sub(r'&nbsp;', ' ', content)
    content = re.sub(r'&amp;', '&', content)
    content = re.sub(r'&lt;', '<', content)
    content = re.sub(r'&gt;', '>', content)
    content = re.sub(r'&quot;', '"', content)
    content = re.sub(r'&#39;', "'", content)

    # Remove Outlook security warnings
    content = re.sub(
        r"You don't often get email from .+?Learn why this is important",
        '',
        content,
        flags=re.IGNORECASE | re.DOTALL
    )

    content = remove_remaining_multilingual_separators(content)

    return content.strip()


def remove_remaining_multilingual_separators(content: str) -> str:
    """
    Remove content after multilingual email separators.
    Simple patterns for common non-English email headers.
    """
    separator_patterns = [
        r'在 \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}，.+?写道：',  # Chinese Simplified
        r'於 \d{4}年\d{1,2}月\d{1,2}日 .+?寫道：',  # Chinese Traditional
        r'Am .+? schrieb .+?:',  # German
        r'Le .+? a écrit :',  # French
        r'El .+? escribió:',  # Spanish
    ]

    for pattern in separator_patterns:
        match = re.search(pattern, content)
        if match:
            # Keep everything before the separator
            content = content[:match.start()].strip()
            break

    return content


def parse_visible_reply_text(content: str) -> str:
    """
    Return only the visible (new) portion of an email reply.
    Uses email_reply_parser when available, otherwise falls back to a
    lightweight heuristic that removes quoted blocks and common separators.
    """
    if not content:
        return ""

    if HAS_EMAIL_REPLY_PARSER:
        try:
            return EmailReplyParser.parse_reply(content).strip()
        except Exception:
            # Fall back to heuristic if the library fails unexpectedly
            pass

    lines = content.splitlines()
    kept_lines: List[str] = []

    separator_patterns = [
        r"^On .+ wrote:\s*$",
        r"^From:\s*.+$",
        r"^-+\s*Original Message\s*-+$",
        r"^_{2,}$",
    ]
    compiled_separators = [re.compile(pat, re.IGNORECASE) for pat in separator_patterns]

    for line in lines:
        # Stop at common reply separators
        if any(pat.search(line) for pat in compiled_separators):
            break
        # Skip quoted lines
        if line.strip().startswith('>'):
            continue
        kept_lines.append(line)

    return "\n".join(kept_lines).strip()


def get_folder_navigation_applescript(mailbox_path: str) -> str:
    """Return AppleScript lines to select the folder path (e.g., 'Inbox/Sub')."""
    if '/' not in mailbox_path:
        # Simple case: top-level folder
        return f'set mb to (first mail folder of acct whose name is "{mailbox_path}")'

    # Complex case: navigate path step by step
    path_parts = mailbox_path.split('/')
    script_lines = []

    # Start with the top-level folder
    script_lines.append(f'set mb to (first mail folder of acct whose name is "{path_parts[0]}")')

    # Navigate through each subfolder
    for i in range(1, len(path_parts)):
        subfolder_name = path_parts[i]
        script_lines.append(f'set mb to (first mail folder of mb whose name is "{subfolder_name}")')

    return '\n        '.join(script_lines)

def select_from_list(items: List[str], prompt: str) -> Optional[str]:
    """Deprecated: retained for compatibility."""
    return items[0] if items else None

def select_upper_and_lower_bound(items: List[str], prompt: str) -> Optional[List[str]]:
    """Deprecated: retained for compatibility."""
    return items if items else None

def select_stream_classification() -> Optional[str]:
    """Deprecated: stream is provided by the paths file."""
    return None

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
        # Print a more detailed error message, including the script that failed for easier debugging
        # print(f"--- Failing AppleScript ---\n{script}\n--------------------------")
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

def get_n_most_recent_emails(account_name: str, mailbox_name: str, n: int) -> List[Email]:
    """Return up to n recent emails for a given account and mailbox path."""
    script = f'''
    tell application "Microsoft Outlook"
        set acct to (first exchange account whose name is "{account_name}")
        {get_folder_navigation_applescript(mailbox_name)}

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
    end tell
    '''

    result = run_applescript(script)
    if not result:
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

                # Skip meeting/booking emails
                if is_meeting_or_booking_email(subject, content):
                    continue

                # Clean the content
                cleaned_content = clean_email_content(content)

                email = Email(
                    subject=subject,
                    content=cleaned_content,
                    received=received
                )
                emails.append(email)

    return emails

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
        r'Meeting Invitation'
    ]

    # Common patterns in booking emails (Microsoft Bookings / one-on-ones)
    booking_patterns = [
        r'^New booking',
        r'New booking',
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
    ]

    # Check subject and content for patterns
    all_patterns = teams_patterns + booking_patterns
    for pattern in all_patterns:
        if re.search(pattern, subject, re.IGNORECASE) or re.search(pattern, content, re.IGNORECASE):
            return True

    return False