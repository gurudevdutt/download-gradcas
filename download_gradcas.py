"""
GradCAS Applicant PDF Downloader
=================================
Automates downloading Application PDFs from upitt-gradcas.admissionsbyliaison.com

Full click path per applicant:
  [List page] Search last name â click row ">"
  â [Profile] click "Applications" in sidebar â click application row ">"
  â [Application] click "ATTACHMENTS" tab
  â click "APPLICATION PDF" section to expand
  â click PDF.js Save/Download button
  â file saved to DOWNLOAD_DIR/LastName_FirstName.pdf

Usage:
  1. Install dependencies (one time):
         pip install playwright openpyxl
         playwright install chromium

  2. Edit CONFIG below to match your Excel file.

  3. Run:
         python download_gradcas.py

  4. When the browser opens, log in via SSO, navigate to the
     Contacts / Applicants / All list page (search box in top-right),
     then press Enter in the terminal.
"""

import asyncio
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

########################################################
# CONFIG
########################################################
EXCEL_PATH     = "try_3applicants.xlsx"          # path to your Excel file
FIRST_NAME_COL = "First Name"               # exact column header for first name
LAST_NAME_COL  = "Last Name"                # exact column header for last name
DOWNLOAD_DIR   = Path("gradcas_downloads")  # folder where PDFs will be saved
GRADCAS_URL    = "https://upitt-gradcas.admissionsbyliaison.com/"

TIMEOUT_MS     = 20_000   # ms to wait for page elements; increase if connection is slow

# Logging configuration
# Set to logging.DEBUG for detailed debugging, logging.INFO for normal operation
LOG_LEVEL      = logging.DEBUG  # Change to logging.INFO once code is working
LOG_FILE       = "playwright_debug.log"  # Log file name\

########################################################
# LOGGING SETUP
########################################################

def setup_logging():
    """Set up logging with configurable level."""
    log_file = Path(LOG_FILE)
    
    # Clear any existing handlers to avoid duplicates
    root_logger = logging.getLogger()
    root_logger.handlers = []
    
    # Set up handlers
    file_handler = logging.FileHandler(log_file, mode='a')  # Append mode
    console_handler = logging.StreamHandler(sys.stdout)
    
    # Set formatter
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Set levels
    file_handler.setLevel(logging.DEBUG)  # Always log everything to file
    console_handler.setLevel(LOG_LEVEL)   # Console level is configurable
    
    # Configure root logger
    root_logger.setLevel(logging.DEBUG)  # Root must be at lowest level
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)
    
    logger = logging.getLogger(__name__)
    logger.info(f"Logging initialized: Console level={logging.getLevelName(LOG_LEVEL)}, File={log_file}")
    return logger

logger = setup_logging()

# ------------------------------------------------------------------------------

async def go_back_to_list(page):
    logger.info(f"🔙 go_back_to_list: Current URL = {page.url}")
    try:
        logger.debug("  Clicking people icon to navigate back...")
        await page.locator("i[data-name='people']").click(timeout=TIMEOUT_MS)
        await wait_and_settle(page, ms=2000)
        
        # Clear any lingering search by clicking the cancel button
        logger.debug("  Looking for cancel button (X) in search field...")
        try:
            # Use JavaScript to find and click the cancel button (more reliable with aria-hidden)
            clicked = await page.evaluate("""
                () => {
                    const icons = Array.from(document.querySelectorAll('i.material-icons'));
                    for (let icon of icons) {
                        if (icon.textContent.trim() === 'cancel') {
                            const style = window.getComputedStyle(icon);
                            if (style.display !== 'none' && style.visibility !== 'hidden') {
                                icon.click();
                                return true;
                            }
                        }
                    }
                    return false;
                }
            """)
            if clicked:
                logger.info("  🖱️  Clicked cancel button to clear search...")
                await wait_and_settle(page, ms=1000)
                logger.info("  ✅ Search cleared")
            else:
                logger.debug("  No visible cancel button found (search may already be clear)")
        except Exception as e:
            logger.debug(f"  Could not click cancel button (search may already be clear): {e}")
        
        logger.info(f"  ✅ go_back_to_list complete. New URL = {page.url}")
    except PWTimeout:
        logger.error("  ❌ Timeout: Could not find people/contacts sidebar icon")
        print("  Could not find people/contacts sidebar icon")
def load_applicants(excel_path, first_col, last_col):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    try:
        fi = headers.index(first_col)
        li = headers.index(last_col)
    except ValueError as e:
        sys.exit(f"Column not found in Excel: {e}\nFound columns: {headers}")

    applicants = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        first = str(row[fi]).strip() if row[fi] else ""
        last  = str(row[li]).strip() if row[li] else ""
        if first or last:
            applicants.append({"first": first, "last": last})

    print(f"Loaded {len(applicants)} applicants from {excel_path}")
    return applicants


def safe_filename(first, last):
    def clean(s):
        return re.sub(r'[^\w\-]', '_', s)
    return f"{clean(last)}_{clean(first)}.pdf"


async def wait_and_settle(page, ms=1500):
    logger.debug(f"  ⏳ wait_and_settle: waiting {ms}ms...")
    await page.wait_for_timeout(ms)
    logger.debug(f"  ✅ wait_and_settle: done ({ms}ms elapsed)")


async def search_and_open_applicant(page, first, last):
    logger.info(f"🔍 search_and_open_applicant: Searching for '{first} {last}'")
    logger.debug(f"  Current URL = {page.url}")
    try:
        # Wait for the SPA to fully load â spinner disappears and search icon appears
        # Wait for page to be ready first
        logger.debug("  Waiting for page to be ready...")
        await page.wait_for_load_state("networkidle", timeout=10_000)
        await page.wait_for_timeout(1000)  # Additional wait for SPA to initialize
        
        await page.locator("i.search-button").wait_for(state="visible", timeout=30_000)
        # Click the magnifying glass icon to reveal the search input
        await page.locator("i.search-button").click(timeout=TIMEOUT_MS)
        logger.debug("  Search button clicked, waiting 800ms...")
        await page.wait_for_timeout(800)
        search_input = page.locator("input[type='text'], input[type='search']").last
        input_count = await page.locator("input[type='text'], input[type='search']").count()
        logger.debug(f"  Found {input_count} search input(s), using last one")
        logger.debug("  Waiting for search input to be visible...")
        await search_input.wait_for(state="visible", timeout=TIMEOUT_MS)
        logger.info("  ✅ Search input is visible")
        logger.info("  🖱️  Clicking search input...")
        await search_input.click(timeout=TIMEOUT_MS)
        # logger.debug("  Triple-clicking to select all...")
        # await search_input.triple_click()
        logger.info(f"  ⌨️  Typing last name: '{last}'")
        await search_input.fill(last)
        logger.info("  ⌨️  Pressing Enter to submit search...")
        await search_input.press("Enter")
        logger.debug("  Waiting 3000ms for search results to update...")
        await page.wait_for_timeout(3000)
        logger.info("  ✅ Search completed")
    except PWTimeout as e:
        logger.error(f"  ❌ Timeout in search_and_open_applicant: {e}")
        logger.debug(f"  Current URL = {page.url}")
        print(f"  Could not find search box")
        return False
    except Exception as e:
        logger.error(f"  ❌ Unexpected error in search_and_open_applicant: {e}")
        print(f"  Error searching: {e}")
        return False

    logger.debug("  Looking for table rows...")
    rows = page.locator("tbody tr")
    count = await rows.count()
    logger.info(f"  Found {count} row(s) in results table")

    if count == 0:
        logger.warning(f"  ⚠️  No results for last name '{last}'")
        print(f"  No results for last name '{last}'")
        return False

    target_row = None
    logger.debug("  Searching for matching row (matching both first and last name)...")
    for i in range(count):
        row = rows.nth(i)
        text = (await row.text_content() or "").strip()
        logger.debug(f"    Row {i+1}: {text[:100]}...")
        if first.lower() in text.lower() and last.lower() in text.lower():
            logger.info(f"  ✅ Found matching row {i+1}")
            target_row = row
            break

    if target_row is None:
        if count == 1:
            logger.info("  Using single result row")
            target_row = rows.first
        else:
            logger.warning(f"  ⚠️  '{first} {last}' not uniquely identified among {count} rows")
            print(f"  '{first} {last}' not uniquely identified among {count} rows")
            return False

    try:
        logger.info("  🖱️  Clicking on target row (last td)...")
        await target_row.locator("td").last.click(timeout=8000)
        logger.info("  ✅ Row clicked successfully")
    except PWTimeout:
        logger.debug("  Last td click failed, trying to click entire row...")
        await target_row.click(timeout=8000)
        logger.info("  ✅ Row clicked successfully (fallback)")

    logger.debug("  Waiting for page to settle after row click...")
    await wait_and_settle(page)
    logger.info(f"  ✅ Navigation complete. New URL = {page.url}")
    return True


async def click_applications_sidebar(page):
    logger.info("📋 click_applications_sidebar")
    logger.debug(f"  Current URL = {page.url}")
    try:
        # Try multiple selectors to find the Applications link
        logger.debug("  Looking for 'Applications' sidebar link...")
        apps_link = None
        
        # Strategy 1: Target the specific div with class "flex-1" containing "Applications"
        try:
            flex_div = page.locator("div.flex-1:has-text('Applications')")
            count = await flex_div.count()
            logger.debug(f"  div.flex-1 with 'Applications': Found {count} element(s)")
            if count > 0:
                # Filter to find the one that actually contains "Applications" text
                for i in range(count):
                    div = flex_div.nth(i)
                    text = await div.text_content()
                    if text and "Applications" in text.strip():
                        is_visible = await div.is_visible(timeout=2000)
                        if is_visible:
                            apps_link = div
                            logger.info(f"  ✅ Found visible Applications div.flex-1 (index {i})")
                            break
                        else:
                            # Try clicking the parent element
                            logger.debug(f"  Div at index {i} not visible, trying parent...")
                            parent = div.locator("..")
                            parent_count = await parent.count()
                            if parent_count > 0:
                                is_visible = await parent.first.is_visible(timeout=2000)
                                if is_visible:
                                    apps_link = parent.first
                                    logger.info(f"  ✅ Found visible parent of Applications div")
                                    break
        except Exception as e:
            logger.debug(f"  div.flex-1 strategy failed: {e}")
        
        # Strategy 2: Try the icon first (fallback)
        if apps_link is None:
            try:
                icon_selector = "i[data-name='applications'], i[data-name='application']"
                icon_count = await page.locator(icon_selector).count()
                logger.debug(f"  Icon selector '{icon_selector}': Found {icon_count} element(s)")
                if icon_count > 0:
                    icon = page.locator(icon_selector).first
                    is_visible = await icon.is_visible(timeout=2000)
                    if is_visible:
                        apps_link = icon
                        logger.info(f"  ✅ Found visible Applications icon")
                    else:
                        logger.debug("  Icon found but not visible, trying parent...")
                        parent = icon.locator("..")
                        parent_count = await parent.count()
                        if parent_count > 0:
                            is_visible = await parent.first.is_visible(timeout=2000)
                            if is_visible:
                                apps_link = parent.first
                                logger.info(f"  ✅ Found visible parent of Applications icon")
            except Exception as e:
                logger.debug(f"  Icon strategy failed: {e}")
        
        # Strategy 3: Try finding parent link/li that contains the div
        if apps_link is None:
            try:
                parent_selector = "a:has(div.flex-1:has-text('Applications')), li:has(div.flex-1:has-text('Applications'))"
                parent_count = await page.locator(parent_selector).count()
                logger.debug(f"  Parent selector with div.flex-1: Found {parent_count} element(s)")
                if parent_count > 0:
                    parent = page.locator(parent_selector).first
                    is_visible = await parent.is_visible(timeout=2000)
                    if is_visible:
                        apps_link = parent
                        logger.info(f"  ✅ Found visible Applications parent element")
                    else:
                        # Try scrolling into view
                        logger.debug("  Element not visible, trying to scroll into view...")
                        await parent.scroll_into_view_if_needed(timeout=2000)
                        is_visible = await parent.is_visible(timeout=2000)
                        if is_visible:
                            apps_link = parent
                            logger.info(f"  ✅ Scrolled Applications element into view")
            except Exception as e:
                logger.debug(f"  Parent strategy failed: {e}")
        
        if apps_link is None:
            logger.error("  ❌ Could not find Applications link with any selector")
            print("  Could not find 'Applications' sidebar link")
            return False
        
        # Click the found element
        logger.info("  🖱️  Clicking 'Applications' sidebar link...")
        await apps_link.click(timeout=TIMEOUT_MS)
        logger.info("  ✅ Click successful")
        await wait_and_settle(page, ms=2000)
        logger.info(f"  ✅ Navigation complete. New URL = {page.url}")
        return True
    except PWTimeout as e:
        logger.error(f"  ❌ Timeout: Could not find or click 'Applications' sidebar link: {e}")
        print("  Could not find 'Applications' sidebar link")
        return False
    except Exception as e:
        logger.error(f"  ❌ Error in click_applications_sidebar: {e}")
        return False


async def click_application_row(page):
    logger.info("📄 click_application_row")
    logger.debug(f"  Current URL = {page.url}")
    try:
        logger.debug("  Looking for application link (td a, tbody tr a)...")
        app_link = page.locator("td a, tbody tr a").first
        count = await page.locator("td a, tbody tr a").count()
        logger.debug(f"  Found {count} application link(s)")
        
        logger.info("  🖱️  Clicking first application link...")
        await app_link.click(timeout=TIMEOUT_MS)
        logger.info("  ✅ Click successful")
        await wait_and_settle(page)
        logger.info(f"  ✅ Navigation complete. New URL = {page.url}")
        return True
    except PWTimeout:
        logger.debug("  First method failed, trying fallback (click last td of first row)...")
        try:
            first_row = page.locator("tbody tr").first
            logger.info("  🖱️  Clicking last td of first row (fallback)...")
            await first_row.locator("td").last.click(timeout=8000)
            logger.info("  ✅ Click successful (fallback)")
            await wait_and_settle(page)
            logger.info(f"  ✅ Navigation complete. New URL = {page.url}")
            return True
        except PWTimeout as e:
            logger.error(f"  ❌ Could not click into application row: {e}")
            print("  Could not click into application row")
            return False
    except Exception as e:
        logger.error(f"  ❌ Error in click_application_row: {e}")
        return False


async def click_attachments_tab(page):
    try:
        await page.locator("text=ATTACHMENTS").click(timeout=TIMEOUT_MS)
        await wait_and_settle(page, ms=2000)
        return True
    except PWTimeout:
        print("  Could not find ATTACHMENTS tab")
        return False


async def expand_application_pdf(page):
    try:
        await page.locator("text=APPLICATION PDF").click(timeout=TIMEOUT_MS)
        await wait_and_settle(page, ms=3000)
        return True
    except PWTimeout:
        print("  Could not find 'APPLICATION PDF' section")
        return False


async def download_pdf(page, download_dir, filename):
    dest = download_dir / filename
    try:
        async with page.expect_download(timeout=30_000) as dl_info:
            try:
                frame = page.frame_locator("iframe").first
                await frame.locator("#downloadButton").click(timeout=10_000)
            except Exception:
                try:
                    frame = page.frame_locator("iframe").first
                    await frame.locator("button[title='Save'], button[title='Download']").first.click(timeout=8000)
                except Exception:
                    await page.locator("button:has-text('Save'), button:has-text('Download')").first.click(timeout=8000)

        download = await dl_info.value
        await download.save_as(dest)
        size_kb = dest.stat().st_size // 1024
        print(f"  Saved: {dest.name}  ({size_kb} KB)")
        return True

    except Exception as e:
        print(f"  Download failed: {e}")
        return False


async def main():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    applicants = load_applicants(EXCEL_PATH, FIRST_NAME_COL, LAST_NAME_COL)

    succeeded = []
    failed    = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False)
        context = await browser.new_context(
            accept_downloads=True,
            viewport={"width": 1400, "height": 900},
        )
        page = await context.new_page()
        
        # Set up event listeners for detailed logging
        def log_click(event):
            # Event handlers must be synchronous, so we just log coordinates
            try:
                logger.debug(f"🖱️  MOUSE CLICK: at ({event.x}, {event.y})")
            except Exception as e:
                logger.debug(f"Error logging click: {e}")
        
        def log_navigation(event):
            try:
                logger.info(f"🌐 NAVIGATION: {event.url}")
            except Exception as e:
                logger.debug(f"Error logging navigation: {e}")
        
        def log_console(msg):
            try:
                logger.debug(f"📝 CONSOLE: {msg.text}")
            except Exception as e:
                logger.debug(f"Error logging console: {e}")
        
        page.on("click", log_click)
        page.on("framenavigated", log_navigation)
        page.on("console", log_console)
        
        try:
            logger.info(f"🚀 Navigating to {GRADCAS_URL}")
            await page.goto(GRADCAS_URL)
            logger.info(f"✅ Initial page loaded. URL = {page.url}")
        except Exception as e:
            logger.error(f"❌ Error navigating to initial page: {e}")
            print(f"Error loading page: {e}")
            print("Browser will remain open. Please navigate manually.")

        print("\n" + "="*60)
        print("Browser is open. Please:")
        print("  1. Log in via Pitt SSO")
        print("  2. Navigate to: Contacts > Applicants > All")
        print("     (the list page with the search box in the top-right)")
        print("  3. Return here and press Enter")
        print("="*60 + "\n")
        
        try:
            input("Press Enter when ready > ")
        except (EOFError, KeyboardInterrupt) as e:
            logger.warning(f"Input interrupted: {e}")
            print("\n⚠️  Input was interrupted. Browser will remain open.")
            print("You can close it manually when done.")
            return

        try:
            list_page_url = page.url
            logger.info(f"📌 List page URL saved: {list_page_url}")
            page_title = await page.title()
            logger.info(f"📌 Current page title: {page_title}")
        except Exception as e:
            logger.error(f"❌ Error getting page info: {e}")
            print(f"Warning: Could not get page info: {e}")
            list_page_url = page.url  # Fallback to current URL

        for i, applicant in enumerate(applicants, 1):
            first = applicant["first"]
            last  = applicant["last"]
            fname = safe_filename(first, last)
            dest  = DOWNLOAD_DIR / fname

            print(f"\n[{i}/{len(applicants)}] {first} {last}")

            if dest.exists():
                print(f"  Already downloaded, skipping")
                succeeded.append(f"{first} {last}")
                continue

            try:
                logger.info("="*60)
                logger.info(f"Starting processing for: {first} {last}")
                logger.info("="*60)
                
                await go_back_to_list(page)

                if not await search_and_open_applicant(page, first, last):
                    failed.append(f"{first} {last} - not found"); continue

                if not await click_applications_sidebar(page):
                    failed.append(f"{first} {last} - Applications sidebar missing"); continue

                if not await click_application_row(page):
                    failed.append(f"{first} {last} - could not open application"); continue

                if not await click_attachments_tab(page):
                    failed.append(f"{first} {last} - ATTACHMENTS tab missing"); continue

                if not await expand_application_pdf(page):
                    failed.append(f"{first} {last} - APPLICATION PDF section missing"); continue

                if await download_pdf(page, DOWNLOAD_DIR, fname):
                    succeeded.append(f"{first} {last}")
                else:
                    failed.append(f"{first} {last} - download failed")

            except Exception as e:
                print(f"  Unexpected error: {e}")
                failed.append(f"{first} {last} - error: {e}")

        await browser.close()

    print("\n" + "="*60)
    print(f"COMPLETE: {len(succeeded)} succeeded, {len(failed)} failed")
    if failed:
        print("\nFailed (can retry by re-running â already downloaded will be skipped):")
        for name in failed:
            print(f"  - {name}")
    print(f"\nFiles saved to: {DOWNLOAD_DIR.resolve()}")
    print("="*60)


if __name__ == "__main__":
    asyncio.run(main())
