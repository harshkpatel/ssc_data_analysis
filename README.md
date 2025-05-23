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

