#!/usr/bin/env python3
"""
FEDAN USD Rate Scraper & WhatsApp Notifier
==========================================
Monitors: http://fedan.com.np/today-foreign-rate.aspx
Windows : 10:00–10:59 AM  and  2:00–2:59 PM (NST)
Interval: Every 1 minute within each window
Output  : C:\\fedan_rate\\fedan_rates.xlsx
Alert   : WhatsApp Web via Selenium (Chrome)
Skips   : Saturday
"""

import os
import sys
import time
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
CHECK_INTERVAL   = 300        # seconds between scrape checks
MORNING_HOUR     = 10        # 10:00 AM – 10:59 AM
EVENING_HOUR     = 14        # 2:00 PM  – 2:59 PM

RATES_SHEET      = "Rates"
RECIPIENTS_SHEET = "Recipients"

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
    "morning": {"rates": [3, 5, 7],   "delivered": [4, 6, 8],    "label": "Morning"},
    "evening": {"rates": [9, 11, 13], "delivered": [10, 12, 14], "label": "Evening"},
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
    wr.append(["Recipient 3", "9801879256"])
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
                        log.debug(f"[{session}] Parsed rate: {rate}")
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

def get_driver():
    """Return (or create) the persistent Chrome Selenium driver."""
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

    try:
        opts = Options()
        # Dedicated Chrome profile stored alongside the Excel file.
        # First run: Chrome opens, scan QR at web.whatsapp.com — stays logged in forever.
        opts.add_argument(f"--user-data-dir={CHROME_PROFILE}")
        opts.add_argument("--profile-directory=Default")
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--no-first-run")
        opts.add_experimental_option("detach", True)   # Keep browser open if script crashes

        svc     = Service(ChromeDriverManager().install())
        _driver = webdriver.Chrome(service=svc, options=opts)
        log.info("Chrome driver started successfully")
        return _driver

    except Exception as exc:
        log.error(f"Chrome init failed: {exc}")
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

        wait = WebDriverWait(drv, 60)  # increased from 30 to 60 seconds

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

        # Step 2: Try multiple send-button selectors (WhatsApp Web updates CSS periodically)
        locators = [
            (By.CSS_SELECTOR, 'button[aria-label="Send"]'),
            (By.XPATH,        '//button[@aria-label="Send"]'),
            (By.XPATH,        '//span[@data-icon="send"]/ancestor::button[1]'),
            (By.CSS_SELECTOR, '[data-tab="11"]'),
        ]

        for locator in locators:
            try:
                btn = wait.until(EC.element_to_be_clickable(locator))
                time.sleep(0.5)
                btn.click()
                time.sleep(3)
                log.info(f"    Sent to {phone}")
                return True
            except Exception:
                continue

        log.error(f"Send button not found for {phone}")
        return False

    except Exception as exc:
        log.error(f"WhatsApp error ({phone}): {exc}")
        return False


def broadcast(wb, message):
    """
    Send message to all recipients in the Recipients sheet.
    Returns True if at least one delivery succeeded.
    """
    numbers = load_recipients(wb)
    if not numbers:
        log.warning("No recipients configured — skipping WhatsApp")
        return False

    any_sent = False
    for phone in numbers:
        ok = send_whatsapp(phone, message)
        if ok:
            any_sent = True
        time.sleep(2)   # Small gap between recipients

    return any_sent


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
        log.debug(f"[{label}] Rate {rate} unchanged since last check")
        return

    state["last_scraped"] = rate
    log.info(f"[{label}] Scraped rate: {rate}")

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
        log.debug(f"[{label}] Rate matches last delivered value ({rate}) — no new entry")
        return

    date_str = today.strftime("%B %d, %Y")
    if next_idx == 0:
        msg = f"FEDAN USD Rate ({label}) - {date_str}: {rate}"
    else:
        msg = f"FEDAN USD Rate ({label} - Updated) - {date_str}: {rate}"

    # ── Send WhatsApp ─────────────────────────────────────────
    sent = broadcast(wb, msg)

    if sent:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        # Write rate AND timestamp only after confirmed delivery
        ws.cell(row=row, column=r_cols[next_idx], value=rate)
        ws.cell(row=row, column=d_cols[next_idx], value=ts)
        wb.save(EXCEL_FILE)
        log.info(f"[{label}] Slot {next_idx + 1} → Rate: {rate} | Delivered: {ts}")
    else:
        log.warning(f"[{label}] Slot {next_idx + 1} → Rate: {rate} | Delivery FAILED — not saved to Excel, will retry next check")


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
    log.info(f"  Chrome profile: {CHROME_PROFILE}")
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
