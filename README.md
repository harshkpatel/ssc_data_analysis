# SSC Data Analysis

This project helps analyze emails from Microsoft Outlook by scraping and processing them into a structured format.

## Important Requirements

### Outlook Version Requirements
- **For macOS Users**: This tool requires the **legacy version of Microsoft Outlook for Mac**. The new Outlook for Mac (v16.XX+) has limited AppleScript support and will not work with this tool.
- **For Windows Users**: This tool requires the **classic Microsoft Outlook desktop app** (part of Microsoft 365/Office 2016/2019/2021). The Microsoft Store version ("Outlook (new)") is **not supported**.

## Prerequisites

Before you begin, you'll need to install some basic tools. Follow the instructions for your operating system:

### For macOS Users

1. **Install Git**
   - Visit [Git for macOS](https://git-scm.com/download/mac)
   - Download and install the latest version
   - To verify installation, open Terminal and type:
     ```bash
     git --version
     ```

2. **Install Python**
   - Visit [Python for macOS](https://www.python.org/downloads/macos/)
   - Download the latest version (3.8 or higher)
   - Run the installer package
   - To verify installation, open Terminal and type:
     ```bash
     python3 --version
     ```

3. **Install pip** (Python package manager)
   - pip comes with Python installation
   - To verify installation, open Terminal and type:
     ```bash
     pip3 --version
     ```

### For Windows Users

1. **Install Git**
   - Visit [Git for Windows](https://git-scm.com/download/win)
   - Download and run the installer
   - Use default settings during installation
   - To verify installation, open Command Prompt and type:
     ```cmd
     git --version
     ```

2. **Install Python**
   - Visit [Python for Windows](https://www.python.org/downloads/windows/)
   - Download the latest version (3.8 or higher)
   - Run the installer
   - **Important**: Check "Add Python to PATH" during installation
   - To verify installation, open Command Prompt and type:
     ```cmd
     python --version
     ```

3. **Install pip** (Python package manager)
   - pip comes with Python installation
   - To verify installation, open Command Prompt and type:
     ```cmd
     pip --version
     ```

## Installation

1. **Clone the repository**
   - Open Terminal (macOS) or Command Prompt (Windows)
   - Navigate to where you want to install the project
   - Run:
     ```bash
      git clone https://github.com/harshkpatel/ssc_data_analysis.git
     ```
     ```bash
     cd ssc_data_analysis
     ```

2. **Install required packages**
   - For macOS:
     ```bash
     pip3 install -r requirements.txt
     ```
   - For Windows:
     ```cmd
     pip install -r requirements.txt
     ```
     ```bash
     pip install pywin32
     ```

## Usage

### For macOS Users

1. **Run the scraper**
   ```bash
   python3 macos/run_mac_scraper.py [options]
   ```

   Options:
   - `--date DD-MM-YYYY`: Scrape emails from a specific date
   - `--count N`: Get the N most recent emails
   - `--debug`: Show recent emails in the selected mailbox

2. **Example commands**
   ```bash
   # Get emails from yesterday
   python3 macos/run_mac_scraper.py

   # Get emails from a specific date
   python3 macos/run_mac_scraper.py --date 01-01-2024

   # Get the 5 most recent emails
   python3 macos/run_mac_scraper.py --count 5
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

- Emails are saved in CSV format in the `csv_files` directory
- Files are named using the format: `account_mailbox--submailbox_date.csv` (slashes in folder paths become double dashes)
- For count-based scraping, files are named: `account_mailbox--submailbox_latest.csv`

## Troubleshooting

### Common Issues

1. **Python not found**
   - Make sure Python is installed and added to PATH
   - Try using `python3` instead of `python` on macOS

2. **Git not found**
   - Ensure Git is properly installed
   - Restart your terminal/command prompt after installation

3. **Package installation fails**
   - Make sure you're in the correct directory
   - Try running the command with administrator privileges
   - **Windows Users**: Make sure to install pywin32 after requirements: `pip install pywin32`

4. **Outlook Version Issues**
   - **macOS**: If you see AppleScript errors, make sure you're using the legacy version of Outlook for Mac
   - **Windows**: If you see COM or "Invalid class string" errors, you're likely using the Microsoft Store version of Outlook
   - Both platforms require the legacy/classic version of Outlook for automation


## Contributing

Feel free to submit issues and enhancement requests!

