# SSC Data Analysis

This project helps analyze emails from Microsoft Outlook by scraping and processing them into a structured format.

## Important Requirements

### Outlook Version Requirements
- **For macOS Users**: This tool requires the **legacy version of Microsoft Outlook for Mac**. The new Outlook for Mac (v16.XX+) has limited AppleScript support and will not work with this tool.
- **For Windows Users**: This tool requires the **classic Microsoft Outlook desktop app** (part of Microsoft 365/Office 2016/2019/2021). The Microsoft Store version ("Outlook (new)") is **not supported**.

## Quick Start

### For macOS Users

1. **Install Python**
   - Visit [Python for macOS](https://www.python.org/downloads/macos/)
   - Download and install the latest version (3.8 or higher)
   - Run the installer package

2. **Install required packages**
   - Open Terminal
   - Run these commands:
     ```bash
     pip install emoji
     ```

### For Windows Users

1. **Install Python**
   - Visit [Python for Windows](https://www.python.org/downloads/windows/)
   - Download the latest version (3.8 or higher)
   - Run the installer
   - **Important**: Check "Add Python to PATH" during installation

2. **Install required packages**
   - Open Command Prompt
   - Run these commands:
     ```cmd
     pip install emoji
     pip install pywin32
     ```

## Usage

### For macOS Users

1. **Run the scraper**
   ```bash
   python3 macos/run_mac_scraper.py [options]
   ```

   **Default behavior (no options):**
   - You will be prompted to select an Outlook account.
   - The script will automatically scrape all emails from these three mailboxes for that account **for yesterday only**:
     - Inbox/Awaiting Information
     - Inbox/Referral Made
     - Inbox/Resolved
   - Only emails from yesterday are added to the local SQLite database (`emails.db`).
   - Duplicate emails (same subject, content, and received date) are automatically ignored.

   **Options:**
   - `--date DD-MM-YYYY`: Scrape emails from a specific date (you will select a mailbox)
   - `--count N`: Get the N most recent emails (you will select a mailbox)
   - `--debug`: Show recent emails in the selected mailbox

2. **Export emails to CSV**
   - Use the `get_emails.py` script to export emails from the database to CSV.
   - **Default behavior (no options):** Only emails from yesterday are exported.
   - Use `--all` to export all emails in the database up to yesterday.

   **Options:**
   - `--date DD-MM-YYYY`: Export emails from a specific date
   - `--all`: Export all emails in the database up to yesterday
   - `--account FirstName`: Use your first name in the CSV filename (otherwise you will be prompted)

   **Example commands:**
   ```bash
   # Export emails from yesterday (will prompt for your first name)
   python3 get_emails.py

   # Export all emails up to yesterday, using your name in the filename
   python3 get_emails.py --all --account Harsh

   # Export emails from a specific date
   python3 get_emails.py --date 31-05-2025 --account Harsh
   ```

### For Windows Users

1. **Run the scraper**
   ```cmd
   python windows/run_win_scraper.py [options]
   ```

   Options:
   - `--date DD-MM-YYYY`: Scrape emails from a specific date
   - `--count N`: Get the N most recent emails
   - `--debug`: Show recent emails in the selected mailbox

2. **Example commands**
   ```cmd
   # Get emails from yesterday
   python windows/run_win_scraper.py

   # Get emails from a specific date
   python windows/run_win_scraper.py --date 01-01-2024

   # Get the 5 most recent emails
   python windows/run_win_scraper.py --count 5
   ```

## Output

- **Default mode (no options):**
  - Emails from yesterday are stored in a local SQLite database (`emails.db`) in the project directory.
  - Each email is uniquely identified by its subject, content, and received date (deduplication is automatic).
- **With get_emails.py:**
  - Emails are saved in CSV format in the `csv_files` directory
  - Files are named using the format: `FirstName_dd-mm-yyyy.csv` for a specific date (or for yesterday, if no date is given), or `FirstName_all.csv` for all emails
- **With --date or --count in run_mac_scraper.py:**
  - Emails are saved in CSV format in the `csv_files` directory
  - Files are named using the format: `account_mailbox--submailbox_date.csv` (slashes in folder paths become double dashes)
  - For count-based scraping, files are named: `account_mailbox--submailbox_latest.csv`

## Troubleshooting

### Common Issues

1. **Python not found**
   - Make sure Python is installed and added to PATH
   - Try using `python3` instead of `python` on macOS

2. **Package installation fails**
   - Make sure you're in the correct directory
   - Try running the command with administrator privileges
   - **Windows Users**: Make sure to install pywin32 after other packages: `pip install pywin32`

3. **Outlook Version Issues**
   - **macOS**: If you see AppleScript errors, make sure you're using the legacy version of Outlook for Mac
   - **Windows**: If you see COM or "Invalid class string" errors, you're likely using the Microsoft Store version of Outlook
   - Both platforms require the legacy/classic version of Outlook for automation

## Contributing

Feel free to submit issues and enhancement requests!

