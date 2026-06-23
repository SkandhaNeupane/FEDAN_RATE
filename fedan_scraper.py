#!/usr/bin/env python3
"""
FEDAN USD Rate Scraper & WhatsApp Notifier
==========================================
Monitors: http://fedan.com.np/today-foreign-rate.aspx
Windows : 10:00–10:59 AM  and  2:00–2:59 PM (NST)
Interval: Every 1 minute within each window
Output  : C:\\fedan_rate\\fedan_rates.xlsx
Alert   : WhatsApp Web via Selenium (Edge)
Skips   : Saturday
"""

import os
import sys
import time
import shutil
import datetime
import logging
import urllib.parse

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
    
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False
    


# ─────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────

BASE_DIR         = r"C:\fedan_rate"
EXCEL_FILE       = os.path.join(BASE_DIR, "fedan_rates.xlsx")
LOG_FILE         = os.path.join(BASE_DIR, "fedan_scraper.log")
CHROME_PROFILE   = os.path.join(BASE_DIR, "chrome_profile")

FEDAN_URL        = "http://fedan.com.np/today-foreign-rate.aspx"
CHECK_INTERVAL   = 60        # seconds between scrape checks
MORNING_HOUR     = 10        # 10:00 AM – 10:59 AM
EVENING_HOUR     = 14        # 2:00 PM  – 2:59 PM

RATES_SHEET      = "Rates"
RECIPIENTS_SHEET = "Recipients"
#GROUP_LINK       = "https://chat.whatsapp.com/D1SzqAW2zzh6pQ88x9ZuUG"
GROUP_NAME       = "FEDANTEST"

# ─────────────────────────────────────────────────────────────
# Excel column layout
# ─────────────────────────────────────────────────────────────
# Col 1=SN   2=Date
#     3=MR1  4=MD1   5=MR2  6=MD2   7=MR3  8=MD3
#     9=ER1  10=ED1  11=ER2 12=ED2  13=ER3 14=ED3
# ─────────────────────────────────────────────────────────────

HEADERS = [
    "S.N.", "Date",
    "Morning Rate 1", "Message Delivered",
    "Morning Rate 2", "Message Delivered",
    "Morning Rate 3", "Message Delivered",
    "Evening Rate 1", "Message Delivered",
    "Evening Rate 2", "Message Delivered",
    "Evening Rate 3", "Message Delivered",
]

SESSION_CFG = {
    "morning": {"rates": [3, 5, 7],   "delivered": [4, 6, 8],    "label": "10AM"},
    "evening": {"rates": [9, 11, 13], "delivered": [10, 12, 14], "label": "2PM"},
}


# ─────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────

os.makedirs(BASE_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────
# In-memory state  (per session, reset each day)
# ─────────────────────────────────────────────────────────────

_state = {
    "morning": {"date": None, "last_scraped": None},
    "evening": {"date": None, "last_scraped": None},
}

_driver = None   # Persistent Selenium Chrome driver


# ═════════════════════════════════════════════════════════════
# Excel helpers
# ═════════════════════════════════════════════════════════════

def load_or_create_workbook():
    """Load existing workbook or create a new one with headers."""
    if os.path.exists(EXCEL_FILE):
        return openpyxl.load_workbook(EXCEL_FILE)

    wb = openpyxl.Workbook()

    # ── Rates sheet ──────────────────────────────────────────
    ws = wb.active
    ws.title = RATES_SHEET
    ws.append(HEADERS)

    header_fill = PatternFill("solid", fgColor="002060")
    for cell in ws[1]:
        cell.font      = Font(bold=True, color="FFFFFF", size=10)
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    col_widths = [6, 12, 14, 20, 14, 20, 14, 20, 14, 20, 14, 20, 14, 20]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Recipients sheet ─────────────────────────────────────
    wr = wb.create_sheet(RECIPIENTS_SHEET)
    wr.append(["Name", "WhatsApp Number"])
    wr.append(["Recipient 1", "9801079561"])
    wr.append(["Recipient 2", "9802079139"])
    for cell in wr[1]:
        cell.font = Font(bold=True)
    wr.column_dimensions["A"].width = 22
    wr.column_dimensions["B"].width = 20

    wb.save(EXCEL_FILE)
    log.info(f"Excel created: {EXCEL_FILE}")
    return wb


def get_today_row(ws):
    """Return row number for today's date, or None if not found."""
    today = datetime.date.today()
    for row in ws.iter_rows(min_row=2):
        val = row[1].value
        row_date = val.date() if isinstance(val, datetime.datetime) else val
        if row_date == today:
            return row[0].row
    return None


def create_today_row(ws):
    """Append a new row for today and return its row number."""
    last_row = ws.max_row
    last_sn  = ws.cell(row=last_row, column=1).value if last_row > 1 else 0
    sn       = (last_sn or 0) + 1
    new_row  = last_row + 1
    ws.cell(row=new_row, column=1, value=sn)
    ws.cell(row=new_row, column=2, value=datetime.date.today())
    return new_row


def load_recipients(wb):
    """Return list of phone number strings from the Recipients sheet."""
    if RECIPIENTS_SHEET not in wb.sheetnames:
        log.warning("Recipients sheet not found")
        return []
    return [
        str(row[1]).strip()
        for row in wb[RECIPIENTS_SHEET].iter_rows(min_row=2, values_only=True)
        if row[1]
    ]


# ═════════════════════════════════════════════════════════════
# Scraping
# ═════════════════════════════════════════════════════════════

def scrape_usd_rate(session: str = "morning"):
    """
    Fetch and parse the USD buying rate from FEDAN.

    The page has TWO separate tables, each inside its own div:
      - Morning (10 AM): the div whose header contains span id="ContentPlaceHolder1_lbl1opm"
      - Evening  (2 PM): the div whose header contains span id="ContentPlaceHolder1_lbl2pm"

    Each table has columns: Currency | Unit | Buying Rate(Average)
    We read tds[2] (Buying Rate) from the correct table only.

    Returns float or None on failure.
    """
    LABEL_ID = {
        "morning": "ContentPlaceHolder1_lbl1opm",
        "evening": "ContentPlaceHolder1_lbl2pm",
    }
    label_id = LABEL_ID[session]

    try:
        resp = requests.get(
            FEDAN_URL,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"},
            timeout=15,
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        # Find the span that identifies this session's block, then walk up to
        # its containing div and find the table inside it.
        label_span = soup.find("span", id=label_id)
        if not label_span:
            log.warning(f"[{session}] Label span '{label_id}' not found — rate may not be published yet")
            return None

        # The span is inside the header div; the table is a sibling inside the same parent div
        container = label_span.find_parent("div", class_="bg-white")
        if not container:
            log.warning(f"[{session}] Could not find container div for label '{label_id}'")
            return None

        table = container.find("table")
        if not table:
            log.warning(f"[{session}] No table found inside container")
            return None

        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) >= 3 and "USD" in tds[0].get_text():
                raw = tds[2].get_text(strip=True)
                try:
                    rate = float(raw)
                    if rate > 0:
                        log.debug(f"[{session}] Parsed rate: {rate:.2f}")
                        return rate
                except ValueError:
                    log.warning(f"[{session}] Could not parse rate value: '{raw}'")
                    return None

        log.warning(f"[{session}] USD row not found in table")
        return None

    except requests.RequestException as exc:
        log.error(f"Network error: {exc}")
        return None


# ═════════════════════════════════════════════════════════════
# WhatsApp via Selenium / WhatsApp Web
# ═════════════════════════════════════════════════════════════

def _wdm_cache_dir():
    """Directory webdriver-manager caches downloaded drivers in."""
    return os.environ.get("WDM_LOCAL", os.path.join(os.path.expanduser("~"), ".wdm"))


def _clear_wdm_cache():
    """
    Wipe webdriver-manager's cached driver metadata.

    webdriver-manager caches the resolved driver for a little while so it
    doesn't have to hit the network on every run. If Chrome auto-updates
    in that window, the cached entry can point at a driver build that no
    longer matches the installed Chrome version. Clearing it forces the
    next install() call to look up and download the correct build fresh.
    """
    cache_dir = _wdm_cache_dir()
    try:
        if os.path.isdir(cache_dir):
            shutil.rmtree(cache_dir, ignore_errors=True)
            log.info(f"Cleared driver cache: {cache_dir}")
    except Exception as exc:
        log.warning(f"Could not clear driver cache ({cache_dir}): {exc}")


def _build_chrome_options():
    opts = Options()
    # Dedicated Chrome profile stored alongside the Excel file.
    # First run: Chrome opens, scan QR at web.whatsapp.com — stays logged in forever.
    opts.add_argument(f"--user-data-dir={CHROME_PROFILE}")
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--no-first-run")
    # detach=True removed — caused orphaned Chrome to lock the profile on restart
    return opts


def get_driver():
    """Return (or create) the persistent Chrome Selenium driver.

    Resolves a ChromeDriver build matching the locally installed Chrome
    version automatically (downloading a new one if needed), so a Chrome
    auto-update can no longer crash the script with a driver-version
    mismatch. Three layers of fallback are tried in order:

      1. Normal webdriver-manager resolve (uses its cache if still valid).
      2. Clear the cache and re-resolve — covers the case where Chrome
         updated but the cached driver entry is now stale/mismatched.
      3. Selenium's own built-in driver manager (Selenium 4.6+) as a
         last resort, in case webdriver-manager itself can't reach its
         download source.
    """
    global _driver

    if _driver:
        try:
            _ = _driver.current_url   # Alive check
            return _driver
        except Exception:
            _driver = None

    if not SELENIUM_OK:
        log.error("Selenium not installed. Run: pip install selenium webdriver-manager")
        return None

    # Kill any orphaned Chrome processes locking the profile directory.
    # This prevents "DevToolsActivePort file doesn't exist" crashes on restart.
    try:
        import subprocess
        subprocess.call(
            ["taskkill", "/F", "/IM", "chrome.exe", "/T"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(2)  # Give Windows time to release the profile lock
    except Exception:
        pass

    opts = _build_chrome_options()

    strategies = ("webdriver_manager", "webdriver_manager_fresh", "selenium_manager")
    last_exc   = None

    for attempt, strategy in enumerate(strategies, start=1):
        try:
            if strategy == "webdriver_manager":
                svc = Service(ChromeDriverManager().install())

            elif strategy == "webdriver_manager_fresh":
                log.warning("Driver init failed — clearing cache and downloading a fresh matching ChromeDriver...")
                _clear_wdm_cache()
                svc = Service(ChromeDriverManager().install())

            else:  # selenium_manager — let Selenium itself resolve/download the driver
                log.warning("webdriver-manager failed twice — falling back to Selenium's built-in driver manager...")
                svc = Service()

            _driver = webdriver.Chrome(service=svc, options=opts)
            log.info(f"Chrome driver started successfully (strategy: {strategy})")
            return _driver

        except Exception as exc:
            last_exc = exc
            log.error(f"Chrome init failed [{strategy}] (attempt {attempt}/{len(strategies)}): {exc}")
            _driver = None
            time.sleep(2)

    log.error(f"All driver init strategies failed — giving up. Last error: {last_exc}")
    return None


def send_whatsapp(phone, message):
    """Send one WhatsApp message via WhatsApp Web. Returns True on success."""
    drv = get_driver()
    if not drv:
        return False

    try:
        encoded = urllib.parse.quote(message)
        url     = f"https://web.whatsapp.com/send?phone={phone}&text={encoded}"
        drv.get(url)

        # Check if QR login is required (first run)
        time.sleep(5)
        page = drv.page_source.lower()
        if "scan" in page and "qr" in page:
            log.warning("WhatsApp Web not logged in — waiting for QR scan")
            print("\n>>> Please scan the QR code in the Chrome window, then press Enter here.")
            input()
            time.sleep(3)

        # Check if 'Use Here' button overlay exists (e.g. if open in another window)
        try:
            use_here_btn = drv.find_element(By.XPATH, "//div[@role='button'][contains(., 'Use Here')]")
            log.info("WhatsApp Web 'Use Here' dialog detected. Clicking 'Use Here'...")
            use_here_btn.click()
            time.sleep(3)
        except Exception:
            pass

        wait = WebDriverWait(drv, 60)  # increased from 30 to 60 seconds
        short_wait = WebDriverWait(drv, 5)

        # Step 1: Wait for the message input box to confirm chat is fully loaded
        input_box_locators = [
            (By.CSS_SELECTOR, 'div[contenteditable="true"][data-tab="10"]'),
            (By.CSS_SELECTOR, 'div[contenteditable="true"][title="Type a message"]'),
            (By.XPATH,        '//div[@contenteditable="true"][@data-tab="10"]'),
        ]
        input_ready = False
        for locator in input_box_locators:
            try:
                wait.until(EC.presence_of_element_located(locator))
                input_ready = True
                log.debug(f"Input box ready for {phone}")
                break
            except Exception:
                continue

        if not input_ready:
            log.error(f"Chat input box not found for {phone} — chat may not have loaded")
            return False

        # Small buffer after input box appears before clicking send
        time.sleep(2)

        # Step 2: Try clicking the Send button via multiple selectors
        send_clicked = False
        locators = [
            (By.CSS_SELECTOR, 'button[aria-label="Send"]'),
            (By.XPATH,        '//button[@aria-label="Send"]'),
            (By.XPATH,        '//span[@data-icon="send"]/ancestor::button[1]'),
            (By.CSS_SELECTOR, '[data-tab="11"]'),
        ]

        for locator in locators:
            try:
                btn = short_wait.until(EC.element_to_be_clickable(locator))
                time.sleep(0.5)
                btn.click()
                send_clicked = True
                break
            except Exception:
                continue

        # Step 3: If button click failed, press Enter on the input box as fallback
        if not send_clicked:
            log.warning(f"Send button not found for {phone} — trying Enter key fallback")
            try:
                from selenium.webdriver.common.keys import Keys
                for locator in input_box_locators:
                    try:
                        box = drv.find_element(*locator)
                        box.send_keys(Keys.ENTER)
                        send_clicked = True
                        log.info(f"    Sent via Enter key to {phone}")
                        break
                    except Exception:
                        continue
            except Exception as e:
                log.error(f"Enter key fallback failed for {phone}: {e}")

        if not send_clicked:
            log.error(f"All send attempts failed for {phone}")
            return False

        # Step 4: Verify message was sent — input box should be empty after send
        time.sleep(3)
        try:
            for locator in input_box_locators:
                try:
                    box = drv.find_element(*locator)
                    if box.text.strip() == "" or box.get_attribute("innerHTML").strip() in ("", "<br>"):
                        log.info(f"    Sent to {phone}")
                        return True
                    else:
                        log.warning(f"Input box not empty after send for {phone} — message may not have sent")
                        return False
                except Exception:
                    continue
        except Exception:
            pass

        # If we can't verify, assume success since click/enter worked
        log.info(f"    Sent to {phone} (unverified)")
        return True

    except Exception as exc:
        log.error(f"WhatsApp error ({phone}): {exc}")
        try:
            error_screenshot = os.path.join(BASE_DIR, "whatsapp_error.png")
            drv.save_screenshot(error_screenshot)
            log.info(f"Error screenshot saved to {error_screenshot}")
        except Exception as se:
            log.error(f"Failed to save error screenshot: {se}")
        return False


def broadcast(wb, message):
    """
    Send message to the WhatsApp group by searching for it by name in WhatsApp Web.
    Returns True if delivery succeeded.
    """
    from selenium.webdriver.common.keys import Keys

    drv = get_driver()
    if not drv:
        log.error("Driver unavailable — skipping WhatsApp")
        return False

    try:
        # Go to WhatsApp Web main page
        if "web.whatsapp.com" not in drv.current_url:
            drv.get("https://web.whatsapp.com")

        # Check if QR login is required (splash screen)
        time.sleep(4)
        page = drv.page_source.lower()
        if "scan" in page and "qr" in page:
            log.warning("WhatsApp Web not logged in — waiting for QR scan")
            print("\n>>> Please scan the QR code in the Edge window, then press Enter here.")
            input()
            time.sleep(3)

        # Check if 'Use Here' button overlay exists (e.g. if open in another window)
        try:
            use_here_btn = drv.find_element(By.XPATH, "//div[@role='button'][contains(., 'Use Here')]")
            log.info("WhatsApp Web 'Use Here' dialog detected. Clicking 'Use Here'...")
            use_here_btn.click()
            time.sleep(3)
        except Exception:
            pass

        wait = WebDriverWait(drv, 60)
        short_wait = WebDriverWait(drv, 5)

        # Wait for WhatsApp Web chat UI to fully load (past splash screen)
        log.info("Step 1: Waiting for WhatsApp Web UI to load...")
        chat_ui_locators = [
            (By.CSS_SELECTOR, '#pane-side'),
            (By.CSS_SELECTOR, 'div[aria-label="Chat list"]'),
            (By.CSS_SELECTOR, 'div[data-testid="chat-list"]'),
            (By.XPATH,        '//div[@role="grid"]'),
        ]
        for locator in chat_ui_locators:
            try:
                wait.until(EC.presence_of_element_located(locator))
                log.info("Step 1: WhatsApp Web UI loaded")
                break
            except Exception:
                continue

        time.sleep(2)

        # Find and click the search box
        search_locators = [
            (By.CSS_SELECTOR, 'input[data-tab="3"]'),
            (By.CSS_SELECTOR, 'input[aria-label="Search or start new chat"]'),
            (By.CSS_SELECTOR, 'div[contenteditable="true"][data-tab="3"]'),
            (By.CSS_SELECTOR, 'div[title="Search input textbox"]'),
            (By.XPATH,        '//div[@contenteditable="true"][@data-tab="3"]'),
            (By.XPATH,        '//input[@data-tab="3"]'),
        ]

        log.info("Step 2: Looking for search box...")
        search_ready = False
        for locator in search_locators:
            try:
                search_box = short_wait.until(EC.element_to_be_clickable(locator))
                search_box.click()
                time.sleep(1)
                
                # Clear any existing text
                search_box.send_keys(Keys.CONTROL + "a")
                search_box.send_keys(Keys.BACKSPACE)
                
                search_box.send_keys(GROUP_NAME)
                search_ready = True
                log.info(f"Step 2: Typed '{GROUP_NAME}' in search box")
                break
            except Exception as e:
                log.info(f"Search locator failed: {locator[1]} — {e}")
                continue

        if not search_ready:
            log.error("Search box not found in WhatsApp Web")
            return False

        time.sleep(3)  # Wait for search results
        log.info("Step 3: Looking for group in search results...")

        # Click the first search result (the group)
        result_locators = [
            (By.XPATH, f'//span[@title="{GROUP_NAME}"]'),
            (By.CSS_SELECTOR, 'div[aria-label="Search results."] div[role="listitem"]:first-child'),
            (By.XPATH, '(//div[@role="listitem"])[1]'),
        ]

        group_found = False
        for locator in result_locators:
            try:
                result = short_wait.until(EC.element_to_be_clickable(locator))
                result.click()
                group_found = True
                log.info(f"Step 4: Group clicked successfully")
                time.sleep(3)
                break
            except Exception as e:
                log.info(f"Group locator failed: {locator[1]} — {e}")
                continue

        if not group_found:
            log.error(f"Group '{GROUP_NAME}' not found in search results")
            return False

        # Wait for input box
        input_box_locators = [
            (By.CSS_SELECTOR, 'div[contenteditable="true"][data-tab="10"]'),
            (By.CSS_SELECTOR, 'div[contenteditable="true"][title="Type a message"]'),
            (By.XPATH,        '//div[@contenteditable="true"][@data-tab="10"]'),
        ]

        input_ready = False
        for locator in input_box_locators:
            try:
                short_wait.until(EC.presence_of_element_located(locator))
                input_ready = True
                break
            except Exception:
                continue

        if not input_ready:
            log.error("Group chat input box not found")
            return False

        time.sleep(2)

        # Type message
        for locator in input_box_locators:
            try:
                box = drv.find_element(*locator)
                box.click()
                box.send_keys(message)
                time.sleep(1)
                break
            except Exception:
                continue

        # Try send button
        send_clicked = False
        send_locators = [
            (By.CSS_SELECTOR, 'button[aria-label="Send"]'),
            (By.XPATH,        '//button[@aria-label="Send"]'),
            (By.XPATH,        '//span[@data-icon="send"]/ancestor::button[1]'),
            (By.CSS_SELECTOR, '[data-tab="11"]'),
        ]

        for locator in send_locators:
            try:
                btn = short_wait.until(EC.element_to_be_clickable(locator))
                time.sleep(0.5)
                btn.click()
                send_clicked = True
                break
            except Exception:
                continue

        # Enter key fallback
        if not send_clicked:
            log.warning("Send button not found — trying Enter key fallback")
            for locator in input_box_locators:
                try:
                    box = drv.find_element(*locator)
                    box.send_keys(Keys.ENTER)
                    send_clicked = True
                    break
                except Exception:
                    continue

        if not send_clicked:
            log.error("All send attempts failed for group")
            return False

        time.sleep(3)
        log.info(f"    Sent to group: {GROUP_NAME}")
        return True

    except Exception as exc:
        log.error(f"Group send error: {exc}")
        try:
            error_screenshot = os.path.join(BASE_DIR, "whatsapp_error.png")
            drv.save_screenshot(error_screenshot)
            log.info(f"Error screenshot saved to {error_screenshot}")
        except Exception as se:
            log.error(f"Failed to save error screenshot: {se}")
        return False


# ═════════════════════════════════════════════════════════════
# Core session logic
# ═════════════════════════════════════════════════════════════

def process_session(session: str):
    """
    Scrape rate, compare with stored values, update Excel, send WhatsApp.
    Handles up to 3 rate changes per session per day.
    """
    today  = datetime.date.today()
    state  = _state[session]
    cfg    = SESSION_CFG[session]
    label  = cfg["label"]
    r_cols = cfg["rates"]
    d_cols = cfg["delivered"]

    # Reset in-memory state on new day
    if state["date"] != today:
        state["date"]         = today
        state["last_scraped"] = None

    # ── Scrape ───────────────────────────────────────────────
    rate = scrape_usd_rate(session)
    if rate is None:
        return

    # Fast path: same rate as last minute's scrape — skip Excel I/O
    if rate == state["last_scraped"]:
        log.debug(f"[{label}] Rate {rate:.2f} unchanged since last check")
        return

    state["last_scraped"] = rate
    log.info(f"[{label}] Scraped rate: {rate:.2f}")

    # ── Open workbook ────────────────────────────────────────
    wb  = load_or_create_workbook()
    ws  = wb[RATES_SHEET]
    row = get_today_row(ws) or create_today_row(ws)

    stored    = [ws.cell(row=row, column=c).value for c in r_cols]
    delivered = [ws.cell(row=row, column=c).value for c in d_cols]

    # Find last stored value and next usable slot (0-based).
    # A slot is usable if empty OR has a rate but was never delivered (retry failed send).
    last_stored_val = None
    next_idx        = None

    for i, val in enumerate(stored):
        if val is not None:
            last_stored_val = val
            if delivered[i] is None:
                log.info(f"[{label}] Slot {i+1} has undelivered rate {val} — retrying")
                next_idx = i
                break
        else:
            next_idx = i
            break

    if next_idx is None:
        log.info(f"[{label}] All 3 slots full — ignoring further changes today")
        return

    # If slot > 0 and rate equals last successfully delivered value → no real change
    prev_delivered = delivered[next_idx - 1] if next_idx > 0 else None
    if next_idx > 0 and prev_delivered is not None and rate == last_stored_val:
        log.debug(f"[{label}] Rate matches last delivered value ({rate:.2f}) — no new entry")
        return

    date_str = today.strftime("%B %d, %Y")
    if next_idx == 0:
        msg = f"FEDAN({label}) - {date_str}: {rate:.2f}"
    else:
        msg = f"FEDAN({label} - Updated) - {date_str}: {rate:.2f}"

    # ── Send WhatsApp ─────────────────────────────────────────
    sent = broadcast(wb, msg)

    if sent:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        # Write rate AND timestamp only after confirmed delivery
        ws.cell(row=row, column=r_cols[next_idx], value=rate)
        ws.cell(row=row, column=d_cols[next_idx], value=ts)
        wb.save(EXCEL_FILE)
        log.info(f"[{label}] Slot {next_idx + 1} → Rate: {rate:.2f} | Delivered: {ts}")
    else:
        log.warning(f"[{label}] Slot {next_idx + 1} → Rate: {rate:.2f} | Delivery FAILED — not saved to Excel, will retry next check")


# ═════════════════════════════════════════════════════════════
# Timing utilities
# ═════════════════════════════════════════════════════════════

def secs_until_hour(hour):
    """Seconds until next occurrence of HH:00:00."""
    now    = datetime.datetime.now()
    target = now.replace(hour=hour, minute=0, second=0, microsecond=0)
    if target <= now:
        target += datetime.timedelta(days=1)
    return (target - now).total_seconds()


# ═════════════════════════════════════════════════════════════
# Main loop
# ═════════════════════════════════════════════════════════════

def main():
    log.info("=" * 55)
    log.info("  FEDAN Rate Scraper  —  Started")
    log.info(f"  Excel file   : {EXCEL_FILE}")
    log.info(f"  Chrome profile  : {CHROME_PROFILE}")
    log.info(f"  Morning window: {MORNING_HOUR}:00 – {MORNING_HOUR}:59")
    log.info(f"  Evening window: {EVENING_HOUR}:00 – {EVENING_HOUR}:59")
    log.info("=" * 55)

    load_or_create_workbook()   # Ensure file + sheets exist on startup

    while True:
        now = datetime.datetime.now()

        # ── Skip Saturday ──────────────────────────────────
        if now.weekday() == 5:
            secs = min(3600, secs_until_hour(MORNING_HOUR))
            log.info(f"Saturday — sleeping {secs / 60:.0f} min")
            time.sleep(secs)
            continue

        # ── Active windows ─────────────────────────────────
        if now.hour == MORNING_HOUR:
            process_session("morning")
            time.sleep(CHECK_INTERVAL)

        elif now.hour == EVENING_HOUR:
            process_session("evening")
            time.sleep(CHECK_INTERVAL)

        # ── Idle — sleep until next window ─────────────────
        else:
            if now.hour < MORNING_HOUR:
                secs  = secs_until_hour(MORNING_HOUR)
                label = f"Morning 10:00 AM (in {secs / 60:.0f} min)"
            elif now.hour < EVENING_HOUR:
                secs  = secs_until_hour(EVENING_HOUR)
                label = f"Evening 2:00 PM (in {secs / 60:.0f} min)"
            else:
                secs  = secs_until_hour(MORNING_HOUR)
                label = f"Tomorrow 10:00 AM (in {secs / 3600:.1f} h)"

            log.info(f"Idle — next window: {label}")
            # Sleep in 15-min chunks so laptop wake from sleep doesn't miss a window
            time.sleep(min(900, secs))


if __name__ == "__main__":
    main()
