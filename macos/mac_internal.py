from mac_outlook_client import (
    get_mailboxes_for_account,
    get_emails_from_date,
    get_most_recent_email,
    select_from_list,
    list_emails_in_mailbox,
    get_n_most_recent_emails
)
import sys
import os
import argparse

parser = argparse.ArgumentParser(description='Scrape Outlook emails on macOS')
parser.add_argument('--date', type=str, help='Date to scrape emails from (DD-MM-YYYY)')
parser.add_argument('--output', type=str, help='Output CSV file (overwrites the default naming)')
parser.add_argument('--latest', action='store_true', help='Get only the most recent email (for testing)')
parser.add_argument('--debug', action='store_true', help='Show recent emails in the selected mailbox')
parser.add_argument('--verbose', action='store_true', help='Show detailed processing information')
parser.add_argument('--count', type=int, help='Number of recent emails to parse')
args = parser.parse_args()

# If any CLI param is provided, prompt for mailbox selection
if args.debug or args.latest or args.count or args.date:
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
    return