#!/usr/bin/env python3
"""
Helpers for interacting with Outlook on macOS via AppleScript and cleaning email content.
"""

import subprocess
import re
from typing import List, Optional
from models.common_models import Email
from email_reply_parser import EmailReplyParser

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

    # Remove Outlook security warnings and Gmail reaction patterns
    content = re.sub(
        r"You don't often get email from .+?Learn why this is important",
        '',
        content,
        flags=re.IGNORECASE | re.DOTALL
    )

    content = re.sub(
        r"Some people who received this message don't often get email from .+?Learn why this is important",
        '',
        content,
        flags=re.IGNORECASE | re.DOTALL
    )

    
    # Microsoft sender identification links
    content = re.sub(r'\[ at https://aka\.ms/LearnAboutSenderIdentification \]','', content)
    
    # Remove Gmail Reactions
    content = re.sub(r'.+?reacted via Gmail','', content, flags=re.IGNORECASE)
    content = re.sub(r'.+?已通过 Gmail\s+做出回应', '', content)
    content = re.sub(r'.+?님이 Gmail\s+을 통해 반응함', '', content)

    # Dates
    content = re.sub(r'\d{4}年\d{1,2}月\d{1,2}日 \d{2}:\d{2}，.+?写道：','', content)
    content = re.sub(r'At \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}, ".+?" wrote:','', content)

    content = remove_remaining_multilingual_separators(content)
    return content.strip()


def remove_remaining_multilingual_separators(content: str) -> str:
    """
    Remove content after multilingual email separators.
    Comprehensive patterns for common non-English email headers and various reply formats.
    """
    separator_patterns = [
        # Chinese patterns
        r'在 \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}，.+?写道：',  # Chinese Simplified
        r'於 \d{4}年\d{1,2}月\d{1,2}日 .+?寫道：',  # Chinese Traditional
        r'\d{4}年\d{1,2}月\d{1,2}日 \d{2}:\d{2}，.+?写道：',  # Chinese with time
        r'\d{4}年\d{1,2}月\d{1,2}日 \d{2}:\d{2}:\d{2}，.+?写道：',  # Chinese with seconds
        
        # German patterns
        r'Am .+? schrieb .+?:',  # German
        r'Am .+? schrieb .+?$',  # German without colon
        
        # French patterns
        r'Le .+? a écrit :',  # French
        r'Le .+? a écrit$',  # French without colon
        
        # Spanish patterns
        r'El .+? escribió:',  # Spanish
        r'El .+? escribió$',  # Spanish without colon
        
        # Japanese patterns
        r'.+?が .+? に返信しました：',  # Japanese reply format
        r'.+?が .+? に返信しました$',  # Japanese without colon
        
        # Korean patterns
        r'.+?님이 .+?에게 답장했습니다：',  # Korean reply format
        r'.+?님이 .+?에게 답장했습니다$',  # Korean without colon
        
        # English patterns with more flexibility
        r'On .+? wrote:',  # Standard English
        r'On .+? wrote$',  # Without colon
        r'At .+?, .+? wrote:',  # With timestamp
        r'At .+?, .+? wrote$',  # Without colon
        r'From: .+? Sent: .+? To: .+? Subject:',  # Email headers
        r'From: .+? Date: .+? To: .+? Subject:',  # Alternative header format
        
        # Generic patterns
        r'_{3,}',  # Multiple underscores
        r'-{3,}',  # Multiple dashes
        r'={3,}',  # Multiple equals
        r'\*{3,}',  # Multiple asterisks
        
        # Outlook/Teams specific
        r'.+?reacted to your message:',  # Teams reactions
        r'.+?reacted to your message$',  # Without colon
        r'Original Message',  # Outlook original message marker
        r'Forwarded message',  # Forwarded message marker
    ]

    for pattern in separator_patterns:
        match = re.search(pattern, content, re.IGNORECASE | re.MULTILINE)
        if match:
            # Keep everything before the separator
            content = content[:match.start()].strip()
            break

    return content


def parse_visible_reply_text(content: str) -> str:
    """
    Return only the visible (new) portion of an email reply.
    Uses email_reply_parser for main parsing, then applies additional regex cleanup
    for any reply headers or artifacts the library might miss.
    """
    if not content:
        return ""

    # Normalize non-breaking spaces to regular spaces
    content = content.replace('\u00A0', ' ')

    # Use the library to get the visible reply
    content = EmailReplyParser.parse_reply(content).strip()

    # Always do additional cleanup with regex patterns to catch what the library misses
    lines = content.splitlines()
    kept_lines: List[str] = []
    
    # Track if we've seen content that looks like a real message
    has_real_content = False
    consecutive_empty_lines = 0
    max_empty_lines = 3  # Allow some empty lines but not too many

    # Stop markers that indicate the start of quoted content or previous thread
    separator_patterns = [
        # Common English reply header - more flexible
        r"^[>\s]*On .{0,300}?wrote\s*$",
        r"^[>\s]*At .{0,300}?wrote\s*$",
        
        # Standard email headers that often start quoted blocks
        r"^From:\s*.+$",
        r"^Sent:\s*.+$",
        r"^To:\s*.+$",
        r"^Subject:\s*.+$",
        r"^Date:\s*.+$",
        r"^Cc:\s*.+$",
        r"^Bcc:\s*.+$",
        
        # Typical separators
        r"^-+\s*Original Message\s*-+$",
        r"^-+\s*Forwarded Message\s*-+$",
        r"^_{2,}\s*$",
        r"^-{2,}\s*$",
        r"^={2,}\s*$",
        
        # Outlook/Teams reactions or artifacts
        r".+reacted to your message\s*[_:]+$",
        r".+reacted via .+$",
        
        # Gmail specific patterns
        r".+?reacted via Gmail",
        r".+?已通过 Gmail\s+做出回应",
        r".+?님이 Gmail 을 통해 반응함",
        
        # Microsoft sender identification
        r"\[ at https://aka\.ms/LearnAboutSenderIdentification \]",
        
        # Multilingual patterns (shorter versions for line-by-line detection)
        r"^.+?写道：\s*$",
        r"^.+?寫道：\s*$",
        r"^.+?schrieb\s*$",
        r"^.+?a écrit\s*$",
        r"^.+?escribió\s*$",
    ]
    compiled_separators = [re.compile(pat, re.IGNORECASE) for pat in separator_patterns]

    for line in lines:
        line_stripped = line.strip()
        
        # Stop at common reply separators
        if any(pat.search(line) for pat in compiled_separators):
            break
            
        # Skip quoted lines but be less strict
        if line_stripped.startswith('>') and len(line_stripped) > 1:
            # Only skip if it's clearly a quoted line with substantial content
            if len(line_stripped) > 10:  # Skip long quoted lines
                continue
            # For short quoted lines, check if they look like actual content
            if not re.match(r'^>\s*[A-Za-z]', line_stripped):
                continue
        
        # Track empty lines
        if not line_stripped:
            consecutive_empty_lines += 1
            if consecutive_empty_lines > max_empty_lines:
                # Too many consecutive empty lines, might be end of content
                break
        else:
            consecutive_empty_lines = 0
            # Check if this looks like real content (not just metadata)
            if len(line_stripped) > 5 and not re.match(r'^[A-Z][a-z]+:\s*', line_stripped):
                has_real_content = True
        
        kept_lines.append(line)

    result = "\n".join(kept_lines).strip()
    
    # If we didn't find any real content, be more lenient and return more lines
    if not has_real_content and len(kept_lines) < len(lines) // 2:
        # Return more content, maybe the parsing was too aggressive
        return content
    
    return result


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
            log "Error in get_n_most_recent_emails: " & errMsg & " (" & errNum & ")"
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

                # Skip completely empty emails
                if not subject and not content:
                    continue

                # Skip meeting/booking emails
                if is_meeting_or_booking_email(subject, content):
                    continue

                # Clean the content - be more lenient with short content
                if len(content) < 50:
                    # Very short content, minimal cleaning
                    cleaned_content = content.strip()
                else:
                    cleaned_content = clean_email_content(content)

                # Only add if we have some meaningful content
                if cleaned_content or subject:
                    # Extract person name from account name (first word)
                    person_name = account_name.split()[0] if account_name else ""
                    
                    email = Email(
                        subject=subject,
                        content=cleaned_content,
                        received=received,
                        person_name=person_name
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