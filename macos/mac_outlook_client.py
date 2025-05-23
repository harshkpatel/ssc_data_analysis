#!/usr/bin/env python3
import subprocess
import re
import html
from typing import List, Dict, Optional
from datetime import datetime, timedelta
from common_models import Email


def clean_email_content(content: str) -> str:
    """
    Clean email content by removing quoted/forwarded content, greetings, signatures,
    and non-printable characters. Normalizes whitespace for better textual analysis.
    """
    # Remove quoted/forwarded content
    lines = content.split('\n')
    cleaned_lines = []
    in_quoted_content = False
    for line in lines:
        if re.match(r'On .*wrote:', line.strip()) or \
           re.match(r'From:.*Sent:.*', line.strip()) or \
           re.match(r'^>.*', line.strip()):
            in_quoted_content = True
            continue
        if in_quoted_content:
            continue
        cleaned_lines.append(line)
    
    # Remove leading/trailing whitespace and empty lines
    cleaned_lines = [l.strip() for l in cleaned_lines if l.strip()]
    if not cleaned_lines:
        return ''

    # Remove greeting (first line if it's a greeting)
    greetings = [
        'hi', 'hello', 'dear', 'hey', 'greetings', 'good morning', 'good afternoon', 'good evening'
    ]
    if cleaned_lines[0].lower().split(',')[0] in greetings or \
       any(cleaned_lines[0].lower().startswith(g + ' ') for g in greetings):
        cleaned_lines = cleaned_lines[1:]

    # Remove signature (lines after a signature word)
    signature_keywords = [
        'thanks', 'thank you', 'regards', 'best', 'cheers', 'sincerely', 'sent from my', 'yours truly', 'warm regards', 'kind regards', 'respectfully', 'with appreciation', 'with gratitude'
    ]
    main_body = []
    for line in cleaned_lines:
        # If the line is a signature keyword or starts with one, stop here
        if any(line.lower().startswith(word) for word in signature_keywords):
            break
        main_body.append(line)
    
    # Join lines and clean up extra whitespace
    cleaned_content = '\n'.join(main_body)
    cleaned_content = re.sub(r'\n\s*\n\s*\n', '\n\n', cleaned_content)  # Remove excessive newlines
    cleaned_content = re.sub(r'[^\x20-\x7E\n]', '', cleaned_content)  # Remove non-printable characters
    cleaned_content = re.sub(r'\s+', ' ', cleaned_content)  # Normalize whitespace
    return cleaned_content.strip()


def run_applescript(script: str) -> str:
    """Execute AppleScript and return the result."""
    process = subprocess.Popen(
        ['osascript', '-e', script],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    stdout, stderr = process.communicate()
    if process.returncode != 0:
        print(f"Error running AppleScript: {stderr.decode('utf-8')}")
        return ""
    return stdout.decode('utf-8').strip()


def get_outlook_accounts() -> List[str]:
    """Get a list of all accounts in Outlook."""
    script = '''
    tell application "Microsoft Outlook"
        set accountList to {}
        repeat with acct in exchange accounts
            set end of accountList to name of acct
        end repeat
        repeat with acct in pop accounts
            set end of accountList to name of acct
        end repeat
        repeat with acct in IMAP accounts
            set end of accountList to name of acct
        end repeat
        return accountList
    end tell
    '''
    result = run_applescript(script)
    if not result:
        return []

    # Parse the comma-separated list
    accounts = [account.strip() for account in result.split(',')]
    return accounts


def get_mailboxes_for_account(account_name: str) -> List[str]:
    """Get all mailboxes for a specific account."""
    script = f'''
    tell application "Microsoft Outlook"
        set mailboxList to {{}}
        set acct to (first exchange account whose name is "{account_name}")
        repeat with mb in mail folders of acct
            set end of mailboxList to name of mb
        end repeat
        return mailboxList
    end tell
    '''
    result = run_applescript(script)
    if not result:
        return []

    # Parse the comma-separated list
    mailboxes = [mailbox.strip() for mailbox in result.split(',')]
    return mailboxes


def get_emails_from_date(account_name: str, mailbox_name: str, target_date: str) -> List[Email]:
    """
    Get emails from a specific date for a given account and mailbox.

    Args:
        account_name: The name of the Outlook account
        mailbox_name: The name of the mailbox to scrape
        target_date: Date in format "DD-MM-YYYY"

    Returns:
        List of Email objects
    """
    # Parse the target date
    day, month, year = target_date.split("-")
    target_date_obj = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
    target_day_name = target_date_obj.strftime("%A")  # Get day name (e.g., "Monday")

    print(f"Target date: {target_date} (which is a {target_day_name})")
    print(f"Looking for emails from: {target_date_obj.strftime('%Y-%m-%d')}")

    # First, get message IDs that match our target date
    script = f'''
    tell application "Microsoft Outlook"
        set allData to {{}}
        set acct to (first exchange account whose name is "{account_name}")
        set mb to (first mail folder of acct whose name is "{mailbox_name}")

        set allMsgs to messages of mb

        repeat with msg in allMsgs
            try
                set msgID to id of msg
                set msgDate to time received of msg
                set msgSubject to subject of msg

                -- Get both the full date string and the components
                set dateStr to msgDate as string
                set msgYear to year of msgDate as string
                set msgMonth to (month of msgDate as integer) as string
                if (count of msgMonth) is 1 then set msgMonth to "0" & msgMonth
                set msgDay to day of msgDate as string
                if (count of msgDay) is 1 then set msgDay to "0" & msgDay
                set fullDate to msgYear & "-" & msgMonth & "-" & msgDay

                -- Add to our list with a special delimiter
                set msgInfo to msgID & "|||ID|||" & dateStr & "|||ID|||" & msgSubject & "|||ID|||" & fullDate
                set end of allData to msgInfo
            on error errMsg
                -- Skip problematic messages
                log "Error with message: " & errMsg
            end try
        end repeat

        set emailText to ""
        repeat with i from 1 to count of allData
            set emailText to emailText & item i of allData
            if i < (count of allData) then set emailText to emailText & "|||EMAIL|||"
        end repeat
        return emailText
    end tell
    '''

    result = run_applescript(script)
    if not result:
        print("No results returned from AppleScript")
        return []

    # Parse the results to identify matching messages
    matching_ids = []
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    last_week = today - timedelta(days=7)

    print("\nProcessing messages...")
    for line in result.split('|||EMAIL|||'):
        if "|||ID|||" in line:
            parts = line.split("|||ID|||")
            if len(parts) >= 4:  # Now expecting 4 parts: ID, date string, subject, full date
                msg_id = parts[0].strip()
                date_str = parts[1].strip()
                subject = parts[2].strip()
                full_date = parts[3].strip()

                print(f"\nChecking message: {subject}")
                print(f"Date string from Outlook: {date_str}")
                print(f"Full date from Outlook: {full_date}")

                # First try the full date we got directly from Outlook
                if full_date == target_date_obj.strftime("%Y-%m-%d"):
                    matching_ids.append(msg_id)
                    print(f"✓ Added (Full date matches target date)")
                    continue

                # Handle abbreviated dates like "Monday", "Today", "Yesterday"
                if date_str == "Today":
                    if today.strftime("%Y-%m-%d") == target_date_obj.strftime("%Y-%m-%d"):
                        matching_ids.append(msg_id)
                        print(f"✓ Added (Today matches target date)")
                elif date_str == "Yesterday":
                    if yesterday.strftime("%Y-%m-%d") == target_date_obj.strftime("%Y-%m-%d"):
                        matching_ids.append(msg_id)
                        print(f"✓ Added (Yesterday matches target date)")
                elif date_str == target_day_name:  # e.g., "Monday"
                    # This is trickier - we need to find the most recent day with this name
                    days_to_subtract = 1
                    while days_to_subtract < 8:  # Look back up to a week
                        check_date = today - timedelta(days=days_to_subtract)
                        if check_date.strftime("%A") == target_day_name:
                            if check_date.strftime("%Y-%m-%d") == target_date_obj.strftime("%Y-%m-%d"):
                                matching_ids.append(msg_id)
                                print(f"✓ Added (Recent {target_day_name} matches target date)")
                            break
                        days_to_subtract += 1
                # Handle "Last Week" or other relative references
                elif "Last Week" in date_str:
                    # Check if target date is within last week
                    if last_week <= target_date_obj <= today:
                        matching_ids.append(msg_id)
                        print(f"✓ Added (Last Week includes target date)")
                else:
                    # Try parsing the full date
                    try:
                        # Try various date formats
                        date_formats = [
                            "%A, %B %d, %Y at %I:%M:%S %p",  # Monday, May 12, 2025 at 2:09:54 PM
                            "%A, %B %d, %Y at %I:%M %p",  # Monday, May 12, 2025 at 2:09 PM
                            "%A, %d %B %Y %H:%M:%S",  # Monday, 12 May 2025 14:09:54
                            "%m/%d/%Y %I:%M:%S %p",  # 05/12/2025 2:09:54 PM
                            "%Y-%m-%d %H:%M:%S",  # 2025-05-12 14:09:54
                            "%B %d, %Y at %I:%M:%S %p",  # May 12, 2025 at 2:09:54 PM
                            "%B %d, %Y at %I:%M %p",  # May 12, 2025 at 2:09 PM
                            "%d %B %Y %H:%M:%S"  # 12 May 2025 14:09:54
                        ]

                        msg_date = None
                        for fmt in date_formats:
                            try:
                                # Try parsing with this format
                                msg_date = datetime.strptime(date_str, fmt)
                                if msg_date.strftime("%Y-%m-%d") == target_date_obj.strftime("%Y-%m-%d"):
                                    matching_ids.append(msg_id)
                                    print(f"✓ Added (Date format {fmt} matches target date)")
                                break
                            except ValueError:
                                continue

                        # If no standard format worked, try extracting components
                        if not msg_date:
                            # Look for patterns like "Monday, May 12, 2025 at 2:09 PM"
                            match = re.search(r'(\w+), (\w+) (\d{1,2}), (\d{4})', date_str)
                            if match:
                                weekday, month_name, day_num, year_num = match.groups()
                                month_map = {
                                    'January': 1, 'February': 2, 'March': 3, 'April': 4,
                                    'May': 5, 'June': 6, 'July': 7, 'August': 8,
                                    'September': 9, 'October': 10, 'November': 11, 'December': 12
                                }
                                month_num = month_map.get(month_name, 1)
                                check_date = datetime(int(year_num), month_num, int(day_num))
                                if check_date.strftime("%Y-%m-%d") == target_date_obj.strftime("%Y-%m-%d"):
                                    matching_ids.append(msg_id)
                                    print(f"✓ Added (Extracted components match target date)")
                    except Exception as e:
                        print(f"Error parsing date '{date_str}': {e}")

    print(f"\nFound {len(matching_ids)} matching messages")

    # Now get the full content for matching messages
    emails = []
    for msg_id in matching_ids:
        script = f'''
        tell application "Microsoft Outlook"
            set acct to (first exchange account whose name is "{account_name}")
            set mb to (first mail folder of acct whose name is "{mailbox_name}")
            set msg to (first message of mb whose id is "{msg_id}")

            set msgSubject to subject of msg
            set msgContent to plain text content of msg
            set msgDate to time received of msg

            set msgYear to year of msgDate as string
            set msgMonth to (month of msgDate as integer) as string
            if (count of msgMonth) is 1 then set msgMonth to "0" & msgMonth
            set msgDay to day of msgDate as string
            if (count of msgDay) is 1 then set msgDay to "0" & msgDay
            set dateOnly to msgYear & "-" & msgMonth & "-" & msgDay

            return msgSubject & "|||DELIM|||" & msgContent & "|||DELIM|||" & dateOnly
        end tell
        '''

        result = run_applescript(script)
        if result and "|||DELIM|||" in result:
            parts = result.split("|||DELIM|||", 2)
            if len(parts) >= 3:
                # Clean the email content before creating the Email object
                cleaned_content = clean_email_content(parts[1].strip())
                email = Email(
                    subject=parts[0].strip(),
                    content=cleaned_content,
                    received=parts[2].strip()
                )
                emails.append(email)

    return emails


def list_emails_in_mailbox(account_name: str, mailbox_name: str, limit: int = 10) -> List[Dict[str, str]]:
    """List recent emails in mailbox for debugging."""
    script = f'''
    tell application "Microsoft Outlook"
        set emailList to {{}}
        set acct to (first exchange account whose name is "{account_name}")
        set mb to (first mail folder of acct whose name is "{mailbox_name}")

        set allMsgs to messages of mb
        set msgCount to count of allMsgs
        set limitCount to {limit}
        if msgCount < limitCount then set limitCount to msgCount

        repeat with i from 1 to limitCount
            set msg to item i of allMsgs
            set msgSubject to subject of msg
            set msgDate to time received of msg
            
            -- Get both the display date and the full date components
            set displayDate to msgDate as string
            set msgYear to year of msgDate as string
            set msgMonth to (month of msgDate as integer) as string
            if (count of msgMonth) is 1 then set msgMonth to "0" & msgMonth
            set msgDay to day of msgDate as string
            if (count of msgDay) is 1 then set msgDay to "0" & msgDay
            set fullDate to msgYear & "-" & msgMonth & "-" & msgDay

            set msgInfo to msgSubject & " | " & displayDate & " | " & fullDate
            set end of emailList to msgInfo
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
    print(f"Raw AppleScript result: {result}")
    if not result:
        return []

    emails_info = []
    # Use a unique delimiter between emails
    # The AppleScript returns a comma-separated list, so join with a unique delimiter in AppleScript
    # Let's fix this by joining with '|||EMAIL|||'
    # But first, check if the delimiter is present
    if '|||EMAIL|||' in result:
        email_lines = result.split('|||EMAIL|||')
    else:
        email_lines = result.split(',')  # fallback for old runs
    for line in email_lines:
        line = line.strip()
        print(f"Parsing line: {line}")
        parts = [p.strip() for p in line.split('|')]
        print(f"  Number of fields: {len(parts)}; Fields: {parts}")
        if len(parts) >= 3:
            subject = parts[0]
            display_date = parts[1]
            full_date = parts[2]
            emails_info.append({
                "subject": subject,
                "display_date": display_date,
                "full_date": full_date
            })

    return emails_info


def get_email_with_attachments(account_name: str, mailbox_name: str, msg_id: str) -> tuple:
    """
    Get email content and information about attachments/images.
    Returns a tuple of (cleaned_content, attachments_info)
    """
    script = f'''
    tell application "Microsoft Outlook"
        set acct to (first exchange account whose name is "{account_name}")
        set mb to (first mail folder of acct whose name is "{mailbox_name}")
        set msg to (first message of mb whose id is "{msg_id}")

        set msgContent to plain text content of msg
        set attachmentInfo to {{}}
        
        -- Get information about attachments
        repeat with att in attachments of msg
            set attName to name of att
            set attSize to size of att
            set attType to content type of att
            
            -- Check if it's an image
            if attType starts with "image/" then
                set end of attachmentInfo to "IMAGE:" & attName & " (" & attSize & " bytes)"
            else
                set end of attachmentInfo to "ATTACHMENT:" & attName & " (" & attSize & " bytes)"
            end if
        end repeat
        
        -- Get information about embedded images
        set embeddedImages to {{}}
        try
            set htmlContent to HTML content of msg
            set imageCount to count of (every paragraph of htmlContent where it contains "<img")
            if imageCount > 0 then
                set end of embeddedImages to "EMBEDDED_IMAGES:" & imageCount & " images found"
            end if
        end try
        
        -- Combine all information
        set allInfo to msgContent & "|||ATTACHMENTS|||" & attachmentInfo & "|||EMBEDDED|||" & embeddedImages
        return allInfo
    end tell
    '''
    
    result = run_applescript(script)
    if not result or "|||ATTACHMENTS|||" not in result:
        return "", []  # Return empty string instead of undefined content
        
    parts = result.split("|||ATTACHMENTS|||")
    content = parts[0].strip()
    
    attachments_info = []
    if len(parts) > 1:
        attachments_part = parts[1].split("|||EMBEDDED|||")
        if attachments_part[0].strip():
            attachments_info.extend(attachments_part[0].strip().split(", "))
        if len(attachments_part) > 1 and attachments_part[1].strip():
            attachments_info.extend(attachments_part[1].strip().split(", "))
    
    # Clean the content
    cleaned_content = clean_email_content(content)
    
    # Add attachment information to the content
    if attachments_info:
        cleaned_content += "\n\n[Attachments and Images:\n" + "\n".join(attachments_info) + "]"
    
    return cleaned_content, attachments_info


def get_most_recent_email(account_name: str, mailbox_name: str) -> Optional[Email]:
    """
    Get the most recent email from a specific account and mailbox.

    Args:
        account_name: The name of the Outlook account
        mailbox_name: The name of the mailbox to scrape

    Returns:
        Most recent Email object or None if no emails found
    """
    script = f'''
    tell application "Microsoft Outlook"
        set acct to (first exchange account whose name is "{account_name}")
        set mb to (first mail folder of acct whose name is "{mailbox_name}")

        set msgs to (messages of mb)
        if (count of msgs) is 0 then
            return ""
        end if

        set latestMsg to item 1 of msgs
        repeat with msg in msgs
            if time received of msg > time received of latestMsg then
                set latestMsg to msg
            end if
        end repeat

        set msgID to id of latestMsg
        set msgSubject to subject of latestMsg
        set msgTime to time received of latestMsg

        set msgYear to year of msgTime as string
        set msgMonth to (month of msgTime as integer) as string
        if (count of msgMonth) is 1 then set msgMonth to "0" & msgMonth
        set msgDay to day of msgTime as string
        if (count of msgDay) is 1 then set msgDay to "0" & msgDay
        set dateOnly to msgYear & "-" & msgMonth & "-" & msgDay

        return msgID & "|||DELIM|||" & msgSubject & "|||DELIM|||" & dateOnly
    end tell
    '''

    result = run_applescript(script)
    if not result or "|||DELIM|||" not in result:
        return None

    parts = result.split("|||DELIM|||", 2)
    if len(parts) >= 3:
        msg_id = parts[0].strip()
        subject = parts[1].strip()
        received = parts[2].strip()
        
        # Get content with attachment information
        content, _ = get_email_with_attachments(account_name, mailbox_name, msg_id)
        
        return Email(
            subject=subject,
            content=content,
            received=received
        )

    return None


def select_from_list(items: List[str], prompt: str) -> Optional[str]:
    """Present a menu for user to select an item from a list."""
    if not items:
        print("No items available.")
        return None

    while True:
        print(prompt)
        print("Enter -1 to exit")
        for i, item in enumerate(items, 1):
            print(f"{i}. {item}")

        try:
            choice = int(input("Enter your choice (number): "))
            if choice == -1:
                print("Exiting program.")
                return None
            if 1 <= choice <= len(items):
                return items[choice - 1]
            else:
                print("Invalid choice. Please try again.")
        except ValueError:
            print("Please enter a valid number.")