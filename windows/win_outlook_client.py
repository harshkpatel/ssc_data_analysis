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
    try:
        day, month, year = target_date.split("-")
        target_date_obj = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
        target_day_name = target_date_obj.strftime("%A")
        print(f"Target date: {target_date} (which is a {target_day_name})")
        print(f"Looking for emails from: {target_date_obj.strftime('%Y-%m-%d')}")
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
        matching_emails = []
        for message in messages:
            try:
                received_time = message.ReceivedTime
                subject = message.Subject
                if not isinstance(received_time, datetime):
                    received_time = datetime.fromtimestamp(received_time.timestamp())
                if received_time.strftime("%Y-%m-%d") == target_date_obj.strftime("%Y-%m-%d"):
                    body = message.Body
                    email = Email(
                        subject=subject,
                        content=body,
                        received=received_time.strftime("%Y-%m-%d")
                    )
                    matching_emails.append(email)
            except Exception as e:
                print(f"Error processing message: {e}")
                continue
        print(f"\nFound {len(matching_emails)} matching messages")
        return matching_emails
    except Exception as e:
        print(f"Error getting emails: {e}")
        return []


def list_emails_in_mailbox(store_display_name: str, mailbox_path: str, limit: int = 10) -> List[Dict[str, str]]:
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
        count = 0
        for message in messages:
            if count >= limit:
                break
            try:
                received_time = message.ReceivedTime
                subject = message.Subject
                if not isinstance(received_time, datetime):
                    received_time = datetime.fromtimestamp(received_time.timestamp())
                emails_info.append({
                    "subject": subject,
                    "display_date": received_time.strftime("%A, %B %d, %Y at %I:%M:%S %p"),
                    "full_date": received_time.strftime("%Y-%m-%d")
                })
                count += 1
            except Exception as e:
                print(f"Error processing message: {e}")
                continue
        return emails_info
    except Exception as e:
        print(f"Error listing emails: {e}")
        return []


def clean_email_content(content: str) -> str:
    """
    Clean email content by removing quoted/forwarded content, greetings, signatures,
    and non-printable characters. Normalizes whitespace for better textual analysis.
    """
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
    cleaned_lines = [l.strip() for l in cleaned_lines if l.strip()]
    if not cleaned_lines:
        return ''
    greetings = [
        'hi', 'hello', 'dear', 'hey', 'greetings', 'good morning', 'good afternoon', 'good evening'
    ]
    if cleaned_lines[0].lower().split(',')[0] in greetings or \
       any(cleaned_lines[0].lower().startswith(g + ' ') for g in greetings):
        cleaned_lines = cleaned_lines[1:]
    signature_keywords = [
        'thanks', 'thank you', 'regards', 'best', 'cheers', 'sincerely', 'sent from my', 'yours truly', 'warm regards', 'kind regards', 'respectfully', 'with appreciation', 'with gratitude'
    ]
    main_body = []
    for line in cleaned_lines:
        if any(line.lower().startswith(word) for word in signature_keywords):
            break
        main_body.append(line)
    cleaned_content = '\n'.join(main_body)
    cleaned_content = re.sub(r'\n\s*\n\s*\n', '\n\n', cleaned_content)
    cleaned_content = re.sub(r'[^\x20-\x7E\n]', '', cleaned_content)
    cleaned_content = re.sub(r'\s+', ' ', cleaned_content)
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
            subject = message.Subject
            if not isinstance(received_time, datetime):
                received_time = datetime.fromtimestamp(received_time.timestamp())
            content, _ = get_email_with_attachments(store_display_name, mailbox_path, message.EntryID)
            email = Email(
                subject=subject,
                content=content,
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
        content, _ = get_email_with_attachments(store_display_name, mailbox_path, message.EntryID)
        return Email(
            subject=message.Subject,
            content=content,
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