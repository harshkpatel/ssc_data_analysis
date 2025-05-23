#!/usr/bin/env python3
import win32com.client
import pythoncom
from typing import List, Dict, Optional
from datetime import datetime, timedelta
from models.common_models import Email


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
    """Get all mailboxes for a specific account."""
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
        
        # Get root folder and its subfolders
        root_folder = account.DeliveryStore.GetRootFolder()
        mailboxes = []
        
        def add_folder(folder):
            mailboxes.append(folder.Name)
            for subfolder in folder.Folders:
                add_folder(subfolder)
        
        add_folder(root_folder)
        return mailboxes
    except Exception as e:
        print(f"Error getting mailboxes: {e}")
        return []


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
    try:
        # Parse the target date
        day, month, year = target_date.split("-")
        target_date_obj = datetime.strptime(f"{year}-{month}-{day}", "%Y-%m-%d")
        target_day_name = target_date_obj.strftime("%A")

        print(f"Target date: {target_date} (which is a {target_day_name})")
        print(f"Looking for emails from: {target_date_obj.strftime('%Y-%m-%d')}")

        # Initialize Outlook
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

        # Find the mailbox folder
        def find_folder(folder, target_name):
            if folder.Name == target_name:
                return folder
            for subfolder in folder.Folders:
                result = find_folder(subfolder, target_name)
                if result:
                    return result
            return None

        root_folder = account.DeliveryStore.GetRootFolder()
        target_folder = find_folder(root_folder, mailbox_name)
        
        if not target_folder:
            print(f"Mailbox not found: {mailbox_name}")
            return []

        # Get all messages
        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        matching_emails = []
        today = datetime.now()
        yesterday = today - timedelta(days=1)
        last_week = today - timedelta(days=7)

        print("\nProcessing messages...")
        for message in messages:
            try:
                received_time = message.ReceivedTime
                subject = message.Subject
                
                # Convert received_time to datetime if it's not already
                if not isinstance(received_time, datetime):
                    received_time = datetime.fromtimestamp(received_time.timestamp())
                
                print(f"\nChecking message: {subject}")
                print(f"Date from Outlook: {received_time}")

                # Check if the message date matches our target date
                if received_time.strftime("%Y-%m-%d") == target_date_obj.strftime("%Y-%m-%d"):
                    # Get the email content
                    body = message.Body
                    
                    # Create Email object
                    email = Email(
                        subject=subject,
                        content=body,
                        received=received_time.strftime("%Y-%m-%d")
                    )
                    matching_emails.append(email)
                    print(f"✓ Added (Date matches target date)")
            except Exception as e:
                print(f"Error processing message: {e}")
                continue

        print(f"\nFound {len(matching_emails)} matching messages")
        return matching_emails

    except Exception as e:
        print(f"Error getting emails: {e}")
        return []


def list_emails_in_mailbox(account_name: str, mailbox_name: str, limit: int = 10) -> List[Dict[str, str]]:
    """List recent emails in mailbox for debugging."""
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

        # Find the mailbox folder
        def find_folder(folder, target_name):
            if folder.Name == target_name:
                return folder
            for subfolder in folder.Folders:
                result = find_folder(subfolder, target_name)
                if result:
                    return result
            return None

        root_folder = account.DeliveryStore.GetRootFolder()
        target_folder = find_folder(root_folder, mailbox_name)
        
        if not target_folder:
            print(f"Mailbox not found: {mailbox_name}")
            return []

        # Get messages
        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        emails_info = []
        count = 0
        
        for message in messages:
            if count >= limit:
                break
                
            try:
                received_time = message.ReceivedTime
                subject = message.Subject
                
                # Convert received_time to datetime if it's not already
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


def get_most_recent_email(account_name: str, mailbox_name: str) -> Optional[Email]:
    """
    Get the most recent email from a specific account and mailbox.

    Args:
        account_name: The name of the Outlook account
        mailbox_name: The name of the mailbox to scrape

    Returns:
        Most recent Email object or None if no emails found
    """
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
            return None

        # Find the mailbox folder
        def find_folder(folder, target_name):
            if folder.Name == target_name:
                return folder
            for subfolder in folder.Folders:
                result = find_folder(subfolder, target_name)
                if result:
                    return result
            return None

        root_folder = account.DeliveryStore.GetRootFolder()
        target_folder = find_folder(root_folder, mailbox_name)
        
        if not target_folder:
            print(f"Mailbox not found: {mailbox_name}")
            return None

        # Get messages
        messages = target_folder.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
        
        if messages.Count == 0:
            return None

        # Get the most recent message
        message = messages.GetFirst()
        received_time = message.ReceivedTime
        
        # Convert received_time to datetime if it's not already
        if not isinstance(received_time, datetime):
            received_time = datetime.fromtimestamp(received_time.timestamp())

        return Email(
            subject=message.Subject,
            content=message.Body,
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