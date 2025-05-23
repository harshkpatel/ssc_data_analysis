# Mac Outlook Email Scraper

This tool allows you to scrape emails from Microsoft Outlook on macOS using AppleScript. It extracts the subject, cleaned content (main body only), and received date of emails.

## Prerequisites

- **Legacy Outlook for Mac:** This scraper is designed for the legacy version of Microsoft Outlook for Mac. The new Outlook for Mac (v16.XX+) has limited AppleScript support, so this tool may not work as expected on newer versions.
- **Python 3.6 or higher:** This tool is compatible with Python 3.6 and above. It has been tested on Python 3.13, but should work on any recent Python 3 release.

## Installation

1. Clone this repository:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Scraping Emails

To scrape emails from your Outlook inbox, run:

```bash
python3 macos/run_mac_scraper.py
```

This script allows you to:
- Scrape emails from a specific date (e.g., `--date 23-05-2025`).
- Scrape the most recent email (use `--latest`).
- Parse a specified number of recent emails (e.g., `--count 5`).
- Debug by listing recent emails (use `--debug`).
- Show detailed processing information (use `--verbose`).

To use any of these features, type in the original command followed by one of the features. (e.g. `python3 macos/run_mac_scraper.py --date 23-05-2025`)

### Customizing the Scraper

- **Account and Mailbox:** The script will prompt you to select an account and mailbox.
- **Output File:** By default, the script saves the results to a CSV file in the `csv_files` directory. You can specify a custom output file using `--output`.

## Caveats

- **Legacy Outlook Required:** This tool is designed for the legacy version of Microsoft Outlook for Mac. The new Outlook for Mac (v16.XX+) has limited AppleScript support, so this tool may not work as expected on newer versions.
- **Performance:** Scraping large inboxes can be slow. Consider limiting the number of emails processed for better performance.

## Windows Usage

### Requirements
- **Classic Microsoft Outlook desktop app** (part of Microsoft 365/Office 2016/2019/2021). 
  - The Microsoft Store version ("Outlook (new)") is **not supported**. You must use the full-featured legacy Outlook desktop app, just as the macOS version requires legacy Outlook.
- Python 3.8+
- All dependencies in `requirements.txt` (install with `pip install -r requirements.txt`)

### Running the Windows Scraper

1. **Open classic Outlook and ensure all accounts and shared mailboxes you want to access are visible in the left pane.**
2. In your terminal, run:
   ```sh
   python windows/run_win_scraper.py --latest
   # or for a specific date:
   python windows/run_win_scraper.py --date 12-05-2025
   # or for N most recent emails:
   python windows/run_win_scraper.py --count 10
   ```
3. **Follow the prompts** to select the Outlook store (account or shared mailbox) and the mailbox (e.g., Inbox).
4. The script will save results to the `csv_files/` directory.

### Notes
- The Windows script uses the same CSV and parsing logic as the macOS version for consistency.
- The only major difference is that on Windows, you select the Outlook store (account or shared mailbox) first, then the mailbox (e.g., Inbox). On macOS, you select the account and mailbox directly.
- Both platforms require the legacy/classic version of Outlook for automation.

### Troubleshooting
- If you do not see your shared mailbox or Inbox, ensure it is added as a full account or store in classic Outlook.
- If you see an error about COM or "Invalid class string", you are likely using the Microsoft Store version of Outlook, which is not supported.

