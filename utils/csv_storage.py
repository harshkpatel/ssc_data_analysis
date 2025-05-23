#!/usr/bin/env python3
import csv
import os
import re
from typing import List, Dict, Optional
from models.common_models import Email


def ensure_directory_exists(file_path: str) -> None:
    """Ensure the directory for the given file path exists."""
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)


def clean_text_for_csv(text: str) -> str:
    """Clean text to make it suitable for CSV."""
    if not text:
        return ""

    # Replace newlines with spaces
    text = re.sub(r'\n+', ' ', text)

    # Replace multiple spaces with one
    text = re.sub(r'\s+', ' ', text)

    # Strip leading/trailing whitespace
    return text.strip()


def save_to_csv(emails: List[Email], output_file: str) -> None:
    """Save email data to a CSV file."""
    if not emails:
        print("No emails to save.")
        return

    try:
        ensure_directory_exists(output_file)

        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['subject', 'content', 'received']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            writer.writeheader()
            for email in emails:
                # Clean the text fields before saving
                writer.writerow({
                    'subject': clean_text_for_csv(email.subject),
                    'content': clean_text_for_csv(email.content),
                    'received': clean_text_for_csv(email.received)
                })

        print(f"Successfully saved {len(emails)} emails to {output_file}")
    except Exception as e:
        print(f"Error saving to CSV: {e}")


def read_from_csv(csv_file: str) -> List[Email]:
    """Read emails from a CSV file."""
    emails = []

    if not os.path.exists(csv_file):
        print(f"File not found: {csv_file}")
        return emails

    try:
        with open(csv_file, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                email = Email(
                    subject=row.get('subject', ''),
                    content=row.get('content', ''),
                    received=row.get('received', '')
                )
                emails.append(email)

        print(f"Successfully read {len(emails)} emails from {csv_file}")
        return emails
    except Exception as e:
        print(f"Error reading from CSV: {e}")
        return []