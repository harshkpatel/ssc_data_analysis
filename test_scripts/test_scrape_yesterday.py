from macos.mac_outlook_client import list_emails_in_mailbox, run_applescript, clean_email_content
import csv
import os

# Ensure test_scraping_data directory exists
os.makedirs('test_scraping_data', exist_ok=True)

# Get the last 3 emails from the inbox (with metadata)
emails_info = list_emails_in_mailbox('harshk.patel@mail.utoronto.ca', 'Inbox', 3)

results = []
for info in emails_info:
    subject = info['subject']
    full_date = info['full_date']
    # Fetch the full content for this subject and date
    script = f'''
    tell application "Microsoft Outlook"
        set acct to (first exchange account whose name is "harshk.patel@mail.utoronto.ca")
        set mb to (first mail folder of acct whose name is "Inbox")
        set msg to (first message of mb whose subject is "{subject}")
        set msgContent to plain text content of msg
        return msgContent
    end tell
    '''
    content = run_applescript(script)
    content = clean_email_content(content)
    results.append((subject, content, full_date))

with open('test_scraping_data/emails.csv', 'w') as f:
    writer = csv.writer(f)
    writer.writerow(['subject', 'content', 'received_date'])
    for row in results:
        writer.writerow(row)

print(f'Wrote {len(results)} emails to test_scraping_data/emails.csv') 