#!/usr/bin/env python3
import sys
import os
import datetime
import re

# Add the parent directory to sys.path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.common_models import Email
from utils.csv_storage import save_to_csv
from mac_outlook_client import (
    get_outlook_accounts,
    select_from_list,
    select_upper_and_lower_bound,
    get_n_most_recent_emails,
    clean_email_content,
    clean_email_subject
)
from utils.sqlite_storage import init_db, insert_emails_bulk, get_all_emails


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
    # Get accounts
    accounts = get_outlook_accounts()
    if not accounts:
        print("No Outlook accounts found.")
        return

    # Let user select account
    selected_accounts = select_upper_and_lower_bound(accounts, "Select an Outlook account:")
    if not selected_accounts:
        return

    # Default behavior: scrape all emails from the three mailboxes and store in SQLite DB
    mailboxes_to_scrape = [
        "Inbox/Awaiting Information",
        "Inbox/Referral Made",
        "Inbox/Resolved"
    ]
    init_db()  # Ensure DB is initialized
    total_new = 0
    for selected_account in selected_accounts:
        for mailbox in mailboxes_to_scrape:
            print(f"\nScraping mailbox: {mailbox}")
            emails = get_n_most_recent_emails(selected_account, mailbox, 10000)  # Large number to get all
            if not emails:
                print(f"No emails found in {mailbox}.")
                continue
            # Only keep emails up to yesterday
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            filtered_emails = [e for e in emails if e.received < today]
            if not filtered_emails:
                print(f"No emails up to yesterday in {mailbox}.")
                continue
            # Prepare for DB insert, cleaning subject and content first
            email_tuples = [(
                clean_email_subject(e.subject),
                clean_email_content(e.content),
                e.received
            ) for e in filtered_emails]
            before_count = len(get_all_emails())
            insert_emails_bulk(email_tuples)
            after_count = len(get_all_emails())
            added = after_count - before_count
            total_new += added
            print(f"Added {added} new emails from {mailbox}.")
        print(f"\nTotal new emails added: {total_new}")


if __name__ == "__main__":
    main()