# GradCAS Applicant PDF Downloader

Automates downloading Application PDFs from upitt-gradcas.admissionsbyliaison.com

## Disclaimer

**USE AT YOUR OWN RISK**

This software is provided "as is", without warranty of any kind. The author is not responsible for:
- Any data loss or corruption
- Any violations of university policies or terms of service
- Any misuse of this software
- Any consequences resulting from the use of this software

Users are responsible for ensuring they have proper authorization to access and download applicant data. Use of this software must comply with all applicable university policies, FERPA regulations, and terms of service of the GradCAS system.

## Setup

1. **Activate the virtual environment:**
   ```bash
   source venv/bin/activate
   ```

2. **Install dependencies (if not already installed):**
   ```bash
   pip install playwright openpyxl
   playwright install chromium
   ```

3. **Edit CONFIG in `download_gradcas.py`** to match your Excel file. I recommend starting with a smaller file to test.

4. **Run the script:**
   ```bash
   python download_gradcas.py
   # OR if venv is not activated:
   ./venv/bin/python download_gradcas.py
   ```

5. **IMPORTANT - When the browser opens:**
   - Log in via Pitt SSO
   - **Navigate to: Contacts > Applicants > All**
     (You should see the list page with a search box in the top-right)
   - A browser alert will remind you to navigate to this page
   - **Press Enter in the terminal** only after you're on the Contacts > Applicants > All page

## Usage Notes

- The script uses logging to `playwright_debug.log` for debugging
- Already downloaded PDFs are skipped automatically
- Check the log file if you encounter issues

## Git

This repository is now version controlled. Always activate the venv before running Python commands.

## Acknowledgment
This script was developed with AI assistance (Claude, Anthropic). The workflow, testing, and site-specific debugging were done by the author.
