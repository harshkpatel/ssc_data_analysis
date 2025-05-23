#!/usr/bin/env python3
import sys
import os
import argparse
from datetime import datetime

# Add the parent directory to sys.path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.common_models import Email
from utils.csv_storage import save_to_csv
from mac_outlook_client import (
    get_outlook_accounts,
    get_mailboxes_for_account,
    get_most_recent_email,
    select_from_list
)


def get_csv_filename(mailbox_name: str, suffix: str = "latest") -> str:
    """Create a standardized filename for the CSV file."""
    # Replace spaces with underscores in mailbox name
    mailbox_clean = mailbox_name.replace(' ', '_')
    return f"csv_files/{mailbox_clean}_{suffix}.csv"


def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Test Outlook email scraper')
    parser.add_argument('--output', type=str, help='Output CSV file (optional)')
    args = parser.parse_args()

    # Get accounts
    accounts = get_outlook_accounts()
    if not accounts:
        print("No Outlook accounts found.")
        return

    # Let user select account
    selected_account = select_from_list(accounts, "Select an Outlook account:")
    if not selected_account:
        return

    # Get mailboxes for selected account
    mailboxes = get_mailboxes_for_account(selected_account)
    if not mailboxes:
        print(f"No mailboxes found for account: {selected_account}")
        return

    # Let user select mailbox
    selected_mailbox = select_from_list(mailboxes, "Select a mailbox:")
    if not selected_mailbox:
        return

    # Get the most recent email
    print(f"Getting the most recent email from {selected_account}/{selected_mailbox}...")
    print("This may take a moment...")
    email = get_most_recent_email(selected_account, selected_mailbox)

    if not email:
        print(f"No emails found in the mailbox")
        return

    # Display the most recent email
    print("\nMost recent email:")
    print(f"Subject: {email.subject}")
    print(f"Received: {email.received}")

    # Show a reasonably sized content preview
    content_preview = email.content[:200] + ("..." if len(email.content) > 200 else "")
    print(f"Content preview: {content_preview}")

    # Set output filename
    output_file = args.output if args.output else get_csv_filename(selected_mailbox)

    # Save to CSV
    save_to_csv([email], output_file)
    print(f"\nEmail saved to {output_file}")

    # Suggest next steps
    print("\nNext steps:")
    print("1. Check the CSV file to ensure data is formatted correctly")
    print("2. If successful, try the full script with: python run_mac_scraper.py")


if __name__ == "__main__":
    main()