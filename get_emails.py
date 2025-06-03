import sys
import os
import datetime
import argparse
import csv
from utils.sqlite_storage import get_all_emails

def parse_args():
    parser = argparse.ArgumentParser(description='Export emails from the local database for a specific date or all dates to a CSV file.')
    parser.add_argument('--date', type=str, help='Date to export emails from (DD-MM-YYYY). Defaults to yesterday if not provided.')
    parser.add_argument('--account', type=str, help='Your first name (for the CSV filename).')
    parser.add_argument('--all', action='store_true', help='Export all emails up to yesterday.')
    args = parser.parse_args()
    return args

def export_to_csv(emails, filename):
    # Ensure csv_files directory exists
    csv_dir = 'csv_files'
    if not os.path.exists(csv_dir):
        os.makedirs(csv_dir)
    filepath = os.path.join(csv_dir, filename)
    with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['subject', 'content', 'received'])
        for subject, content, received in emails:
            writer.writerow([subject, content, received])
    print(f"Exported {len(emails)} emails to {filepath}")

def main():
    args = parse_args()
    account_name = args.account
    if not account_name:
        account_name = input("Enter your first name (for the CSV filename): ").strip()
    account_clean = account_name.replace(' ', '-').replace('/', '--')
    emails = get_all_emails()
    if args.all:
        # Export all emails up to yesterday
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        filtered = [e for e in emails if e[2] < today]
        if not filtered:
            print("No emails found in the database up to yesterday.")
            return
        filename = f"{account_clean}_all.csv"
        export_to_csv(filtered, filename)
    else:
        # Export for a specific date (default: yesterday)
        if args.date:
            try:
                day, month, year = args.date.split('-')
                date_obj = datetime.datetime(int(year), int(month), int(day))
            except Exception:
                print("Invalid date format. Please use DD-MM-YYYY.")
                sys.exit(1)
        else:
            date_obj = datetime.datetime.now() - datetime.timedelta(days=1)
        target_date = date_obj.strftime("%Y-%m-%d")
        found = [e for e in emails if e[2] == target_date]
        if not found:
            print(f"No emails found for {target_date}.")
            return
        # Format date as dd-mm-yyyy for filename
        y, m, d = target_date.split('-')
        date_for_filename = f"{d}-{m}-{y}"
        filename = f"{account_clean}_{date_for_filename}.csv"
        export_to_csv(found, filename)

if __name__ == "__main__":
    main() 