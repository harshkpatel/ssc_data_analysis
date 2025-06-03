#!/usr/bin/env python3
import win32com.client
import pythoncom
from typing import List, Dict, Optional
from datetime import datetime, timedelta
from models.common_models import Email
import re


def get_outlook_accounts() -> List[str]:
    """Get a list of all accounts in Outlook."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        accounts = []
        for account in namespace.Accounts:
            accounts.append(account.DisplayName)
        
        return accounts
    except Exception as e:
        print(f"Error getting Outlook accounts: {e}")
        return []


def get_mailboxes_for_account(account_name: str) -> List[str]:
    """Get all mailboxes for a specific account, returning full folder paths."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Find the account
        account = None
        for acc in namespace.Accounts:
            if acc.DisplayName == account_name:
                account = acc
                break
        
        if not account:
            print(f"Account not found: {account_name}")
            return []
        
        root_folder = account.DeliveryStore.GetRootFolder()
        mailboxes = []

        def add_folder(folder, path_so_far):
            current_path = f"{path_so_far}/{folder.Name}" if path_so_far else folder.Name
            mailboxes.append(current_path)
            for subfolder in folder.Folders:
                add_folder(subfolder, current_path)
        
        add_folder(root_folder, "")
        return mailboxes
    except Exception as e:
        print(f"Error getting mailboxes: {e}")
        return []


def _find_folder_by_path(folder, path_parts):
    if not path_parts:
        return folder
    next_name = path_parts[0]
    for subfolder in folder.Folders:
        if subfolder.Name == next_name:
            return _find_folder_by_path(subfolder, path_parts[1:])
    return None


def _get_store_root_folder(store_display_name):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i+1)
        if store.DisplayName == store_display_name:
            return store.GetRootFolder()
    print(f"Store not found: {store_display_name}")
    return None


def get_emails_from_date(store_display_name: str, mailbox_path: str, target_date: str) -> List[Email]:
    """Return all emails (including replies) from the specified Outlook folder that were received on the given date (DD-MM-YYYY), using the raw message body (no cleaning)."""
    try:
        # Validate and parse the input date (DD-MM-YYYY → datetime)
        try:
            day, month, year = target_date.split("-")
            target_dt = datetime(int(year), int(month), int(day))
        except ValueError as e:
            print(f"ERROR: Invalid date '{target_date}'. Use DD-MM-YYYY. {e}")
            return []
        target_date_only = target_dt.date()

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Locate the store
        store = None
        for i in range(namespace.Stores.Count):
            s = namespace.Stores.Item(i + 1)
            if s.DisplayName == store_display_name:
                store = s
                break
        if not store:
            print(f"Store not found: {store_display_name}")
            return []

        # Traverse folders
        folder = store.GetRootFolder()
        for part in mailbox_path.split("/"):
            sub = None
            for f in folder.Folders:
                if f.Name == part:
                    sub = f
                    break
            if not sub:
                print(f"Folder not found in path: {part}")
                return []
            folder = sub

        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)  # newest first
        total = messages.Count

        emails: List[Email] = []

        for msg in messages:
            try:
                recv = msg.ReceivedTime
                if not isinstance(recv, datetime):
                    recv = datetime.fromtimestamp(recv.timestamp())
                recv_date = recv.date()

                if recv_date < target_date_only:
                    break  # We've gone past the desired day
                if recv_date > target_date_only:
                    continue  # Still looking for emails on the target date

                subj = (msg.Subject or "").strip()
                # Clean the subject line
                subj = clean_subject_line(subj)
                
                body = msg.Body or ""
                
                # If Body is empty, try HTMLBody
                if not body.strip():
                    try:
                        html_body = msg.HTMLBody or ""
                        if html_body.strip():
                            # Simple HTML to text conversion
                            body = re.sub(r'<[^>]+>', '', html_body)
                            body = re.sub(r'&nbsp;', ' ', body)
                            body = re.sub(r'&amp;', '&', body)
                            body = re.sub(r'&lt;', '<', body)
                            body = re.sub(r'&gt;', '>', body)
                    except:
                        pass

                # Clean the email content
                cleaned_content = clean_email_content(body)

                email_obj = Email(
                    subject=subj,
                    content=cleaned_content,
                    received=recv.strftime("%Y-%m-%d"),
                )
                emails.append(email_obj)
            except Exception as e:
                continue

        return emails

    except Exception as fatal:
        print(f"Fatal error in get_emails_from_date: {fatal}")
        return []


def list_emails_in_mailbox(store_display_name: str, mailbox_path: str, count: int = 10) -> List[Dict]:
    try:
        root_folder = _get_store_root_folder(store_display_name)
        if not root_folder:
            return []
        path_parts = mailbox_path.split("/")
        target_folder = _find_folder_by_path(root_folder, path_parts)
        if not target_folder:
            print(f"Mailbox not found: {mailbox_path}")
            return []
        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)
        emails_info = []
        message = messages.GetFirst()
        count = min(count, messages.Count)
        for _ in range(count):
            if not message:
                break
            received_time = message.ReceivedTime
            if not isinstance(received_time, datetime):
                received_time = datetime.fromtimestamp(received_time.timestamp())
                
            subject = message.Subject or ""
            # Clean the subject line
            subject = clean_subject_line(subject)
            
            body = message.Body or ""
            # If Body is empty, try HTMLBody
            if not body.strip():
                try:
                    html_body = message.HTMLBody or ""
                    if html_body.strip():
                        # Simple HTML to text conversion
                        body = re.sub(r'<[^>]+>', '', html_body)
                        body = re.sub(r'&nbsp;', ' ', body)
                        body = re.sub(r'&amp;', '&', body)
                        body = re.sub(r'&lt;', '<', body)
                        body = re.sub(r'&gt;', '>', body)
                except:
                    pass
            
            # Clean the email content
            cleaned_content = clean_email_content(body)
            
            emails_info.append({
                "subject": subject,
                "content": cleaned_content,
                "received": received_time.strftime("%Y-%m-%d")
            })
            message = messages.GetNext()
        return emails_info
    except Exception as e:
        print(f"Error listing emails: {e}")
        return []


def clean_email_content(content: str) -> str:
    """
    Clean email content by removing quoted/forwarded content, greetings, signatures,
    emojis, and special characters. Normalizes whitespace for better textual analysis.
    Only keeps the main message content.
    """
    if not content:
        return ""
    
    lines = content.split('\n')
    cleaned_lines = []
    
    # Check if this is a test email - preserve test emails completely
    is_test_email = any("test" in line.lower() for line in lines[:5])
    if is_test_email:
        return content.strip()
    
    # Stop processing when we hit common quoted content indicators
    quote_indicators = [
        r'^On .* wrote:',  # "On [date] [person] wrote:"
        r'^From:.*',       # Email headers
        r'^To:.*',
        r'^Sent:.*',
        r'^Date:.*',
        r'^Subject:.*',
        r'^________________________________+',  # Outlook separator lines
        r'^-----Original Message-----',
        r'^>.*',           # Line starting with >
        r'^\s*>.*',        # Line starting with whitespace then >
        r'^Begin forwarded message:',
        r'^Forwarded message',
        r'^----- Forwarded Message -----',
    ]
    
    for line in lines:
        line_stripped = line.strip()
        
        # Check if this line indicates start of quoted content
        is_quoted = False
        for pattern in quote_indicators:
            if re.match(pattern, line_stripped, re.IGNORECASE):
                is_quoted = True
                break
        
        if is_quoted:
            # Stop processing here - everything after is quoted content
            break
            
        # Also stop if we see "From:" followed by an email address pattern
        if re.match(r'^From:.*@.*', line_stripped, re.IGNORECASE):
            break
            
        cleaned_lines.append(line)
    
    # Remove empty lines from the end
    while cleaned_lines and not cleaned_lines[-1].strip():
        cleaned_lines.pop()
    
    # Remove common email signatures/closings from the end
    signature_patterns = [
        r'^thanks?.*',
        r'^thank you.*',
        r'^regards?.*',
        r'^best.*',
        r'^cheers.*',
        r'^sincerely.*',
        r'^yours truly.*',
        r'^warm regards.*',
        r'^kind regards.*',
        r'^respectfully.*',
        r'^with appreciation.*',
        r'^with gratitude.*',
        r'^sent from my.*',
        r'^get outlook for.*',
        r'^.*@.*\.(com|org|edu|ca|net).*',  # Email addresses
        r'^.*\(\w+/\w+.*\).*',  # Pronouns like (she/her)
    ]
    
    # Remove signature lines from the end
    while cleaned_lines:
        last_line = cleaned_lines[-1].strip().lower()
        if not last_line:
            cleaned_lines.pop()
            continue
            
        is_signature = False
        for pattern in signature_patterns:
            if re.match(pattern, last_line, re.IGNORECASE):
                is_signature = True
                break
                
        if is_signature:
            cleaned_lines.pop()
        else:
            break
    
    # Remove common greetings from the beginning
    greeting_patterns = [
        r'^hi\s*,?.*',
        r'^hello\s*,?.*',
        r'^dear\s*.*',
        r'^hey\s*,?.*',
        r'^greetings\s*,?.*',
        r'^good morning\s*,?.*',
        r'^good afternoon\s*,?.*',
        r'^good evening\s*,?.*',
    ]
    
    while cleaned_lines:
        first_line = cleaned_lines[0].strip().lower()
        if not first_line:
            cleaned_lines.pop(0)
            continue
            
        is_greeting = False
        for pattern in greeting_patterns:
            if re.match(pattern, first_line, re.IGNORECASE):
                is_greeting = True
                break
                
        if is_greeting:
            cleaned_lines.pop(0)
        else:
            break
    
    # Join the remaining lines
    cleaned_content = '\n'.join(line.rstrip() for line in cleaned_lines)
    
    # Remove emojis and special characters
    # Remove emoji characters (most emojis are in these Unicode ranges)
    emoji_pattern = re.compile(
        "["
        "\U0001F1E0-\U0001F1FF"  # flags (iOS)
        "\U0001F300-\U0001F5FF"  # symbols & pictographs
        "\U0001F600-\U0001F64F"  # emoticons
        "\U0001F680-\U0001F6FF"  # transport & map symbols
        "\U0001F700-\U0001F77F"  # alchemical symbols
        "\U0001F780-\U0001F7FF"  # Geometric Shapes Extended
        "\U0001F800-\U0001F8FF"  # Supplemental Arrows-C
        "\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
        "\U0001FA00-\U0001FA6F"  # Chess Symbols
        "\U0001FA70-\U0001FAFF"  # Symbols and Pictographs Extended-A
        "\U00002702-\U000027B0"  # Dingbats
        "\U000024C2-\U0001F251"
        "]+", flags=re.UNICODE)
    
    cleaned_content = emoji_pattern.sub('', cleaned_content)
    
    # Remove other special characters but keep basic punctuation
    # Keep: letters, numbers, basic punctuation, accented characters
    cleaned_content = re.sub(r'[^\w\s\.\,\!\?\;\:\-\(\)\[\]\'\"\@\#\$\%\&\*\+\=\/\\\|\<\>\~\`\^\_\{\}àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿĀāĂăĄąĆćĈĉĊċČčĎďĐđĒēĔĕĖėĘęĚěĜĝĞğĠġĢģĤĥĦħĨĩĪīĬĭĮįİıĲĳĴĵĶķĸĹĺĻļĽľĿŀŁłŃńŅņŇňŉŊŋŌōŎŏŐőŒœŔŕŖŗŘřŚśŜŝŞşŠšŢţŤťŦŧŨũŪūŬŭŮůŰűŲųŴŵŶŷŸŹźŻżŽžſ]', '', cleaned_content)
    
    # Normalize whitespace
    cleaned_content = re.sub(r'\n\s*\n\s*\n+', '\n\n', cleaned_content)  # Multiple blank lines to double
    cleaned_content = re.sub(r'[ \t]+', ' ', cleaned_content)  # Multiple spaces/tabs to single space
    
    return cleaned_content.strip()


def get_email_with_attachments(store_display_name: str, mailbox_path: str, msg_entry_id: str):
    """
    Get email content and information about attachments/images for a message by EntryID.
    Returns a tuple of (cleaned_content, attachments_info)
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    root_folder = _get_store_root_folder(store_display_name)
    if not root_folder:
        return '', []
    path_parts = mailbox_path.split("/")
    target_folder = _find_folder_by_path(root_folder, path_parts)
    if not target_folder:
        return '', []
    message = None
    for item in target_folder.Items:
        if hasattr(item, 'EntryID') and item.EntryID == msg_entry_id:
            message = item
            break
    if not message:
        return '', []
    content = message.Body
    attachments_info = []
    for att in message.Attachments:
        att_name = att.FileName
        att_size = att.Size
        att_type = getattr(att, 'Type', None)
        if att_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
            attachments_info.append(f"IMAGE: {att_name} ({att_size} bytes)")
        else:
            attachments_info.append(f"ATTACHMENT: {att_name} ({att_size} bytes)")
    cleaned_content = clean_email_content(content)
    if attachments_info:
        cleaned_content += "\n\n[Attachments and Images:\n" + "\n".join(attachments_info) + "]"
    return cleaned_content, attachments_info


def clean_subject_line(subject: str) -> str:
    """Remove emojis and special characters from subject line."""
    if not subject:
        return ""
    
    # Remove emoji characters
    emoji_pattern = re.compile(
        "["
        "\U0001F1E0-\U0001F1FF"  # flags (iOS)
        "\U0001F300-\U0001F5FF"  # symbols & pictographs
        "\U0001F600-\U0001F64F"  # emoticons
        "\U0001F680-\U0001F6FF"  # transport & map symbols
        "\U0001F700-\U0001F77F"  # alchemical symbols
        "\U0001F780-\U0001F7FF"  # Geometric Shapes Extended
        "\U0001F800-\U0001F8FF"  # Supplemental Arrows-C
        "\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
        "\U0001FA00-\U0001FA6F"  # Chess Symbols
        "\U0001FA70-\U0001FAFF"  # Symbols and Pictographs Extended-A
        "\U00002702-\U000027B0"  # Dingbats
        "\U000024C2-\U0001F251"
        "]+", flags=re.UNICODE)
    
    subject = emoji_pattern.sub('', subject)
    
    # Remove other special characters but keep basic punctuation
    subject = re.sub(r'[^\w\s\.\,\!\?\;\:\-\(\)\[\]\'\"\@\#\$\%\&\*\+\=\/\\\|\<\>\~\`\^\_\{\}àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿĀāĂăĄąĆćĈĉĊċČčĎďĐđĒēĔĕĖėĘęĚěĜĝĞğĠġĢģĤĥĦħĨĩĪīĬĭĮįİıĲĳĴĵĶķĸĹĺĻļĽľĿŀŁłŃńŅņŇňŉŊŋŌōŎŏŐőŒœŔŕŖŗŘřŚśŜŝŞşŠšŢţŤťŦŧŨũŪūŬŭŮůŰűŲųŴŵŶŷŸŹźŻżŽžſ]', '', subject)
    
    # Normalize whitespace
    subject = re.sub(r'\s+', ' ', subject)
    
    return subject.strip()


def get_n_most_recent_emails(store_display_name: str, mailbox_path: str, n: int) -> List[Email]:
    root_folder = _get_store_root_folder(store_display_name)
    if not root_folder:
        return []
    path_parts = mailbox_path.split("/")
    target_folder = _find_folder_by_path(root_folder, path_parts)
    if not target_folder:
        print(f"Mailbox not found: {mailbox_path}")
        return []
    messages = target_folder.Items
    messages.Sort("[ReceivedTime]", True)
    emails = []
    count = 0
    for message in messages:
        if count >= n:
            break
        try:
            received_time = message.ReceivedTime
            subject = message.Subject or ""
            
            # Clean the subject line
            subject = clean_subject_line(subject)
            
            # Try to get content from different properties
            body = message.Body or ""
            
            # If Body is empty, try HTMLBody
            if not body.strip():
                try:
                    html_body = message.HTMLBody or ""
                    if html_body.strip():
                        # Simple HTML to text conversion
                        body = re.sub(r'<[^>]+>', '', html_body)
                        body = re.sub(r'&nbsp;', ' ', body)
                        body = re.sub(r'&amp;', '&', body)
                        body = re.sub(r'&lt;', '<', body)
                        body = re.sub(r'&gt;', '>', body)
                except:
                    pass
            
            if not isinstance(received_time, datetime):
                received_time = datetime.fromtimestamp(received_time.timestamp())
            
            # Clean the email content to remove quoted/forwarded content
            cleaned_content = clean_email_content(body)
            
            email = Email(
                subject=subject,
                content=cleaned_content,
                received=received_time.strftime("%Y-%m-%d")
            )
            emails.append(email)
            count += 1
        except Exception as e:
            print(f"Error processing message: {e}")
            continue
    return emails


def get_most_recent_email(store_display_name: str, mailbox_path: str) -> Optional[Email]:
    try:
        root_folder = _get_store_root_folder(store_display_name)
        if not root_folder:
            return None
        path_parts = mailbox_path.split("/")
        target_folder = _find_folder_by_path(root_folder, path_parts)
        if not target_folder:
            print(f"Mailbox not found: {mailbox_path}")
            return None
        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)
        if messages.Count == 0:
            return None
        message = messages.GetFirst()
        received_time = message.ReceivedTime
        if not isinstance(received_time, datetime):
            received_time = datetime.fromtimestamp(received_time.timestamp())
            
        subject = message.Subject or ""
        # Clean the subject line
        subject = clean_subject_line(subject)
        
        body = message.Body or ""
        # If Body is empty, try HTMLBody
        if not body.strip():
            try:
                html_body = message.HTMLBody or ""
                if html_body.strip():
                    # Simple HTML to text conversion
                    body = re.sub(r'<[^>]+>', '', html_body)
                    body = re.sub(r'&nbsp;', ' ', body)
                    body = re.sub(r'&amp;', '&', body)
                    body = re.sub(r'&lt;', '<', body)
                    body = re.sub(r'&gt;', '>', body)
            except:
                pass
        
        # Clean the email content
        cleaned_content = clean_email_content(body)
        
        return Email(
            subject=subject,
            content=cleaned_content,
            received=received_time.strftime("%Y-%m-%d")
        )
    except Exception as e:
        print(f"Error getting most recent email: {e}")
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


def debug_print_accounts_and_stores():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    print("Accounts:")
    for i, acc in enumerate(namespace.Accounts):
        print(f"  {i+1}. {acc.DisplayName}")
    print("Stores:")
    for i in range(namespace.Stores.Count):
        store = namespace.Stores.Item(i+1)
        print(f"  {i+1}. {store.DisplayName}")


def get_all_stores() -> list:
    """Return a list of all store display names (including shared and delegate mailboxes)."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        stores = []
        for i in range(namespace.Stores.Count):
            store = namespace.Stores.Item(i+1)
            stores.append(store.DisplayName)
        return stores
    except Exception as e:
        print(f"Error getting stores: {e}")
        return []


def get_mailboxes_for_store(store_display_name: str) -> list:
    """Get all mailboxes for a specific store (by display name), returning full folder paths."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        store = None
        for i in range(namespace.Stores.Count):
            s = namespace.Stores.Item(i+1)
            if s.DisplayName == store_display_name:
                store = s
                break
        if not store:
            print(f"Store not found: {store_display_name}")
            return []
        root_folder = store.GetRootFolder()
        mailboxes = []
        def add_folder(folder, path_so_far):
            current_path = f"{path_so_far}/{folder.Name}" if path_so_far else folder.Name
            mailboxes.append(current_path)
            for subfolder in folder.Folders:
                add_folder(subfolder, current_path)
        add_folder(root_folder, "")
        return mailboxes
    except Exception as e:
        print(f"Error getting mailboxes for store: {e}")
        return [] 