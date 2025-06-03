#!/usr/bin/env python3
import sys
import os
import datetime
import argparse
import re
from typing import List, Optional

# Add the parent directory to sys.path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.common_models import Email
from utils.csv_storage import save_to_csv
from mac_outlook_client import (
    get_outlook_accounts,
    get_mailboxes_for_account,
    get_emails_from_date,
    get_most_recent_email,
    select_from_list,
    list_emails_in_mailbox,
    get_n_most_recent_emails
)


def validate_date(date_str: str) -> bool:
    """
    Validate that the date is not in the future and not today.
    Returns True if date is valid, False otherwise.
    """
    try:
        # Parse the input date
        day, month, year = date_str.split('-')
        target_date = datetime.datetime(int(year), int(month), int(day))
        
        # Get today's date (without time)
        today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        # Check if date is today or in the future
        if target_date >= today:
            print(f"Error: Cannot scrape emails for today ({today.strftime('%d-%m-%Y')}) or future dates.")
            return False
            
        return True
    except ValueError as e:
        print(f"Error: Invalid date format. Please use DD-MM-YYYY format. Error: {e}")
        return False


def get_csv_filename(account_name: str, mailbox_name: str, date_str: str) -> str:
    """
    Create a standardized filename for the CSV file.
    Format: account_mailbox--submailbox_date.csv (with slashes replaced by double dashes)
    Files saved in csv_files directory
    """
    # Replace spaces with dashes in account name and slashes with double dashes in mailbox name
    account_clean = account_name.replace(' ', '-')
    mailbox_clean = mailbox_name.replace(' ', '-').replace('/', '--')

    # Get the absolute path to the csv_files directory
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    csv_dir = os.path.join(project_root, 'csv_files')
    
    # Ensure the directory exists
    if not os.path.exists(csv_dir):
        os.makedirs(csv_dir)

    # For latest/count files
    if date_str in ["latest", "count", "test_latest"]:
        return os.path.join(csv_dir, f"{account_clean}_{mailbox_clean}_latest.csv")

    # For date-based files, convert DD-MM-YYYY to YYYY-MM-DD
    if re.match(r'\d{2}-\d{2}-\d{4}', date_str):
        day, month, year = date_str.split('-')
        date_clean = f"{year}-{month}-{day}"
    else:
        date_clean = date_str

    return os.path.join(csv_dir, f"{account_clean}_{mailbox_clean}_{date_clean}.csv")


def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Scrape Outlook emails on macOS')
    parser.add_argument('--date', type=str, help='Date to scrape emails from (DD-MM-YYYY)')
    parser.add_argument('--output', type=str, help='Output CSV file (overwrites the default naming)')
    parser.add_argument('--latest', action='store_true', help='Get only the most recent email (for testing)')
    parser.add_argument('--debug', action='store_true', help='Show recent emails in the selected mailbox')
    parser.add_argument('--verbose', action='store_true', help='Show detailed processing information')
    parser.add_argument('--count', type=int, help='Number of recent emails to parse')
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

    # Debug mode - list recent emails
    if args.debug:
        print(f"\nListing recent emails in {selected_account}/{selected_mailbox} for debugging:")
        emails = list_emails_in_mailbox(selected_account, selected_mailbox, 20)
        if not emails:
            print("No emails found in this mailbox.")
            return

        print("\nRecent emails:")
        for i, email in enumerate(emails, 1):
            print(f"{i}. Subject: {email['subject']}")
            print(f"   Display Date: {email['display_date']}")
            print(f"   Full Date: {email['full_date']}")
            print()

        # If in verbose mode, also display the current date for reference
        if args.verbose:
            today = datetime.datetime.now()
            print(f"\nToday is: {today.strftime('%A, %B %d, %Y')}")
            last_week = today - datetime.timedelta(days=7)
            print(f"Last week: {last_week.strftime('%A, %B %d, %Y')} to {today.strftime('%A, %B %d, %Y')}")

        return

    if args.latest:
        # Get the most recent email (for testing)
        print(f"Getting the most recent email from {selected_account}/{selected_mailbox}...")
        email = get_most_recent_email(selected_account, selected_mailbox)

        if not email:
            print("No emails found.")
            return

        # Set default output filename if not specified
        output_file = args.output if args.output else get_csv_filename(selected_account, selected_mailbox, "test_latest")

        # Save to CSV
        save_to_csv([email], output_file)

        # Display preview
        print("\nMost recent email:")
        print(f"Subject: {email.subject}")
        print(f"Received: {email.received}")
        content_preview = email.content[:100] + ("..." if len(email.content) > 100 else "")
        print(f"Content preview: {content_preview}")
        print(f"\nSaved to: {output_file}")

    elif args.count:
        # Get the specified number of recent emails
        print(f"Getting the last {args.count} emails from {selected_account}/{selected_mailbox}...")
        emails = get_n_most_recent_emails(selected_account, selected_mailbox, args.count)
        if not emails:
            print("No emails found.")
            return

        # Set default output filename if not specified
        output_file = args.output if args.output else get_csv_filename(selected_account, selected_mailbox, "latest")

        # Save to CSV
        save_to_csv(emails, output_file)

        # Display summary
        print(f"\nFound {len(emails)} emails.")
        for i, email in enumerate(emails[:3], 1):  # Show preview of first 3
            print(f"\nEmail {i}:")
            print(f"Subject: {email.subject}")
            print(f"Received: {email.received}")
            content_preview = email.content[:100] + ("..." if len(email.content) > 100 else "")
            print(f"Content preview: {content_preview}")

        if len(emails) > 3:
            print(f"\n... and {len(emails) - 3} more emails.")

        print(f"\nSaved to: {output_file}")

    else:
        # Get date to scrape
        target_date = args.date
        if not target_date:
            # Default to yesterday
            yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
            target_date = yesterday.strftime("%d-%m-%Y")
            print(f"No date specified. Using yesterday: {target_date}")
        
        # Validate the date
        if not validate_date(target_date):
            return

        # Set default output filename if not specified
        output_file = args.output if args.output else get_csv_filename(selected_account, selected_mailbox, target_date)

        # Get emails
        print(f"Scraping emails from {selected_account}/{selected_mailbox} on {target_date}...")
        # Show verbose output only if requested
        old_stdout = sys.stdout
        if not args.verbose:
            sys.stdout = open(os.devnull, 'w')

        emails = get_emails_from_date(selected_account, selected_mailbox, target_date)

        # Restore stdout
        if not args.verbose:
            sys.stdout = old_stdout

        if not emails:
            print(f"No emails found for the date: {target_date}")
            return

        # Save to CSV
        save_to_csv(emails, output_file)

        # Display summary
        print(f"\nFound {len(emails)} emails for {target_date}")
        for i, email in enumerate(emails[:3], 1):  # Show preview of first 3
            print(f"\nEmail {i}:")
            print(f"Subject: {email.subject}")
            print(f"Received: {email.received}")
            content_preview = email.content[:100] + ("..." if len(email.content) > 100 else "")
            print(f"Content preview: {content_preview}")

        if len(emails) > 3:
            print(f"\n... and {len(emails) - 3} more emails.")

        print(f"\nSaved to: {output_file}")


if __name__ == "__main__":
    main()