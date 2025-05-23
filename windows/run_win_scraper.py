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
from win_outlook_client import (
    get_outlook_accounts,
    get_mailboxes_for_account,
    get_emails_from_date,
    get_most_recent_email,
    select_from_list,
    list_emails_in_mailbox,
    debug_print_accounts_and_stores,
    get_all_stores,
    get_mailboxes_for_store,
    get_n_most_recent_emails
)


def get_csv_filename(account_name: str, mailbox_name: str, date_str: str) -> str:
    """
    Create a standardized filename for the CSV file.
    Format: account_mailbox_date.csv (with spaces replaced by underscores)
    """
    # Replace spaces with underscores in account and mailbox names
    account_clean = account_name.replace(' ', '_')
    mailbox_clean = mailbox_name.replace(' ', '_')

    # If date_str is in DD-MM-YYYY format, convert to YYYY-MM-DD
    if re.match(r'\d{2}-\d{2}-\d{4}', date_str):
        day, month, year = date_str.split('-')
        date_clean = f"{year}-{month}-{day}"
    else:
        date_clean = date_str

    # Get the absolute path to the global csv_files folder
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    csv_dir = os.path.join(project_root, 'csv_files')
    
    # Ensure the directory exists
    if not os.path.exists(csv_dir):
        os.makedirs(csv_dir)

    return os.path.join(csv_dir, f"{account_clean}_{mailbox_clean}_{date_clean}.csv")


def main():
    # Print all accounts and stores for debugging
    debug_print_accounts_and_stores()

    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Scrape Outlook emails on Windows')
    parser.add_argument('--date', type=str, help='Date to scrape emails from (DD-MM-YYYY)')
    parser.add_argument('--output', type=str, help='Output CSV file (overwrites the default naming)')
    parser.add_argument('--latest', action='store_true', help='Get only the most recent email')
    parser.add_argument('--debug', action='store_true', help='Show recent emails in the selected mailbox')
    parser.add_argument('--verbose', action='store_true', help='Show detailed processing information')
    parser.add_argument('--count', type=int, help='Number of recent emails to parse')
    args = parser.parse_args()

    # Let user select store (not just account)
    stores = get_all_stores()
    if not stores:
        print("No Outlook stores found.")
        return
    selected_store = select_from_list(stores, "Select an Outlook store (account or shared mailbox):")
    if not selected_store:
        return

    # Get mailboxes for selected store
    mailboxes = get_mailboxes_for_store(selected_store)
    if not mailboxes:
        print(f"No mailboxes found for store: {selected_store}")
        return

    # Let user select mailbox (now full path)
    selected_mailbox = select_from_list(mailboxes, "Select a mailbox (full path):")
    if not selected_mailbox:
        return

    # Remove store name prefix from mailbox path for lookup
    if selected_mailbox.startswith(selected_store):
        mailbox_path = selected_mailbox[len(selected_store):]
        if mailbox_path.startswith("/"):
            mailbox_path = mailbox_path[1:]
    else:
        mailbox_path = selected_mailbox

    # Debug mode - list recent emails
    if args.debug:
        print(f"\nListing recent emails in {selected_store}/{mailbox_path} for debugging:")
        emails = list_emails_in_mailbox(selected_store, mailbox_path, 20)
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
        # Get the most recent email
        print(f"Getting the most recent email from {selected_store}/{mailbox_path}...")
        email = get_most_recent_email(selected_store, mailbox_path)

        if not email:
            print("No emails found.")
            return

        # Set default output filename if not specified
        output_file = args.output if args.output else get_csv_filename(selected_store, mailbox_path, "latest")

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
        print(f"Getting the last {args.count} emails from {selected_store}/{mailbox_path}...")
        emails = get_n_most_recent_emails(selected_store, mailbox_path, args.count)
        if not emails:
            print("No emails found.")
            return

        # Set default output filename if not specified
        output_file = args.output if args.output else get_csv_filename(selected_store, mailbox_path, "recent")

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

        # Set default output filename if not specified
        output_file = args.output if args.output else get_csv_filename(selected_store, mailbox_path, target_date)

        # Get emails
        print(f"Scraping emails from {selected_store}/{mailbox_path} on {target_date}...")
        # Show verbose output only if requested
        old_stdout = sys.stdout
        if not args.verbose:
            sys.stdout = open(os.devnull, 'w')

        emails = get_emails_from_date(selected_store, mailbox_path, target_date)

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