# GradCAS Applicant PDF Downloader

Automates downloading Application PDFs from upitt-gradcas.admissionsbyliaison.com

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

3. **Edit CONFIG in `download_gradcas.py`** to match your Excel file.

4. **Run the script:**
   ```bash
   python download_gradcas.py
   # OR if venv is not activated:
   ./venv/bin/python download_gradcas.py
   ```

5. When the browser opens:
   - Log in via Pitt SSO
   - Navigate to: Contacts > Applicants > All
   - Press Enter in the terminal

## Usage Notes

- The script uses logging to `playwright_debug.log` for debugging
- Already downloaded PDFs are skipped automatically
- Check the log file if you encounter issues

## Git

This repository is now version controlled. Always activate the venv before running Python commands.
