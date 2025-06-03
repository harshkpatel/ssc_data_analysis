from windows.win_outlook_client import list_emails_in_mailbox, clean_email_content
import csv
import os

# Ensure test_scraping_data directory exists
os.makedirs('test_scraping_data', exist_ok=True)

# Get the last 3 emails from the inbox (with metadata)
emails_info = list_emails_in_mailbox('harshk.patel@mail.utoronto.ca', 'Inbox', 3)

results = []
for info in emails_info:
    subject = info['subject']
    content = info['content']  # Get content directly from the info
    full_date = info['full_date']
    content = clean_email_content(content)
    results.append((subject, content, full_date))

with open('test_scraping_data/emails.csv', 'w', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['subject', 'content', 'received_date'])
    for row in results:
        writer.writerow(row)

print(f'Wrote {len(results)} emails to test_scraping_data/emails.csv') 