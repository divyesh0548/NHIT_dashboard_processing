"""
ICICI Bank Violation Report Bot

Flow implemented:
1) Read ICICI_Reference_file.xlsx (case-insensitive headers)
2) Login per row (Plaza_name + Username + Password + Done)
3) Navigate: Reports -> Reconciliation Reports -> Transactions Reconciliation And Settlements
4) Split configured date range into non-overlapping 3-day intervals
5) For each interval:
   - Set Start Date and End Date
   - Submit and wait (no fixed timeout) for export button
   - Export Excel
   - Detect newly downloaded .csv/.xlsx in Downloads (last 10 minutes)
   - Move file to a plaza-named folder next to this script with interval-based filename
6) Mark row as done only after all intervals are completed.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from datetime import datetime, timedelta
import shutil
import time

import numpy as np
import pandas as pd
from PIL import Image
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

from ocr_module import extract_text, special_character_remover
from db_config import get_process_date_range_strings


LOGIN_URL = "https://acqweb.icici.bank.in/cebacons"
REFERENCE_EXCEL = "ICICI_Reference_file.xlsx"
MAX_LOGIN_ATTEMPTS = 5
DOWNLOADS_FOLDER = r"C:\Users\Divyesh Parmar\Downloads"
# Date range comes from PROCESS_FROM_DATE / PROCESS_TO_DATE in project-root .env
INTERVAL_SOURCE_MODE = "date_range"  # "date_range" or "excel"
INTERVAL_LIST_EXCEL = "ICICI_interval_list.xlsx"
SKIPPED_INTERVALS_EXCEL = "ICICI_skipped_intervals.xlsx"
POLL_SECONDS = 2
INTERVAL_DAYS = 3  # Max inclusive days per download interval
MAX_ELEMENT_WAIT_SECONDS = 180
MAX_REPORT_WAIT_SECONDS = 240
SCRIPT_DIR = Path(__file__).resolve().parent


class RestartBrowserException(Exception):
    """Raised when the browser session must be restarted and the interval retried."""


def setup_driver(headless: bool = False):
    """Create Chrome driver with practical defaults."""
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    prefs = {
        "download.default_directory": str(Path(DOWNLOADS_FOLDER).resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=chrome_options)


def normalize_header_name(value) -> str:
    """Normalize header text for case-insensitive matching."""
    if value is None:
        return ""
    return str(value).strip().lower()


def resolve_required_columns(df: pd.DataFrame) -> dict:
    """
    Resolve required columns by lowercase header comparison.
    Required columns:
      - Plaza_name
      - Username
      - Password
      - Done
    """
    wanted = {
        "plaza_name": None,
        "username": None,
        "password": None,
        "done": None,
    }

    for col in df.columns:
        norm = normalize_header_name(col)
        if norm in wanted and wanted[norm] is None:
            wanted[norm] = col

    missing = [k for k, v in wanted.items() if v is None]
    if missing:
        raise ValueError(
            f"Missing required columns in {REFERENCE_EXCEL}: {missing}. "
            "Expected headers: Plaza_name, Username, Password, Done (case-insensitive)."
        )
    return wanted


def resolve_interval_columns(df: pd.DataFrame) -> dict:
    """
    Resolve interval-list Excel columns by lowercase header comparison.
    Required columns:
      - Plaza_name
      - Start_Date
      - End_Date
    """
    wanted = {
        "plaza_name": None,
        "start_date": None,
        "end_date": None,
    }

    for col in df.columns:
        norm = normalize_header_name(col)
        if norm in wanted and wanted[norm] is None:
            wanted[norm] = col

    missing = [k for k, v in wanted.items() if v is None]
    if missing:
        raise ValueError(
            f"Missing required columns in interval Excel: {missing}. "
            "Expected headers: Plaza_name, Start_Date, End_Date (case-insensitive)."
        )
    return wanted


def mark_done_in_excel(excel_file_path: str, row_index: int, done_header_name):
    """Mark Done='done' for the given DataFrame row index."""
    excel_row = row_index + 2  # +2: header row is 1 and df index starts at 0
    wb = load_workbook(excel_file_path)
    sheet = wb.active

    done_col = None
    for col_idx in range(1, sheet.max_column + 1):
        header = sheet.cell(row=1, column=col_idx).value
        if normalize_header_name(header) == normalize_header_name(done_header_name):
            done_col = col_idx
            break

    if done_col is None:
        wb.close()
        raise ValueError("Could not locate 'Done' column in Excel while marking status.")

    sheet.cell(row=excel_row, column=done_col, value="done")
    wb.save(excel_file_path)
    wb.close()
    print(f"Marked row {excel_row} as done in {excel_file_path}")


def get_captcha_image_array(driver) -> np.ndarray:
    """Capture captcha image element and return as numpy array for OCR."""
    captcha_element = wait_for_visible_element(driver, By.ID, "imgCaptcha")
    png_bytes = captcha_element.screenshot_as_png
    image = Image.open(BytesIO(png_bytes)).convert("RGB")
    return np.array(image)


def solve_captcha_text(driver) -> str:
    """Extract captcha text using OCR and sanitize to alphanumeric."""
    captcha_array = get_captcha_image_array(driver)
    extracted = extract_text(captcha_array)
    raw = " ".join(extracted)
    cleaned = special_character_remover(raw)
    # Captcha field maxlength is 5, trim to avoid overflow.
    return cleaned[:5]


def is_login_successful(driver) -> bool:
    """
    Determine whether login succeeded.
    Heuristic:
    - If username/password fields are still visible, we are likely still on login page.
    - Otherwise consider login successful.
    """
    try:
        username_inputs = driver.find_elements(By.ID, "txtUserName")
        password_inputs = driver.find_elements(By.ID, "txtPassword")
        if username_inputs and username_inputs[0].is_displayed():
            return False
        if password_inputs and password_inputs[0].is_displayed():
            return False
    except Exception:
        pass
    return True


def print_login_errors_if_any(driver):
    """Best-effort extraction of visible error text after login attempt."""
    selectors = [
        ".error",
        ".alert",
        ".validation-summary-errors",
        "[class*='error']",
        "[id*='error']",
    ]
    for selector in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            for elem in elements:
                if elem.is_displayed():
                    text = (elem.text or "").strip()
                    if text:
                        print(f"Error message: {text}")
        except Exception:
            continue


def handle_duplicate_login_popup(driver, timeout_seconds: int = 5) -> bool:
    """
    If the duplicate active-session popup appears after login, click Yes.
    Returns True when the popup was found and handled.
    """
    try:
        wait = WebDriverWait(driver, timeout_seconds)
        popup = wait.until(
            EC.presence_of_element_located((By.ID, "pnlDuplicateLogin"))
        )
        if not popup.is_displayed():
            return False

        yes_button = wait.until(EC.element_to_be_clickable((By.ID, "btnYes")))
        yes_button.click()
        print("Duplicate session popup detected. Clicked 'Yes' to continue login.")

        try:
            WebDriverWait(driver, timeout_seconds).until(EC.invisibility_of_element(popup))
        except Exception:
            pass

        return True
    except Exception:
        return False


def handle_post_login_notification_popup(driver, timeout_seconds: int = 6) -> bool:
    """
    If the post-login Notification popup appears, click its close button and continue.
    Popup markup: #Body_divComment with close button #Body_btnClose.
    Returns True when the popup was found and dismissed.
    """
    try:
        wait = WebDriverWait(driver, timeout_seconds)
        popup = wait.until(
            EC.presence_of_element_located((By.ID, "Body_divComment"))
        )
        if not popup.is_displayed():
            return False

        heading = ""
        try:
            heading_el = driver.find_element(By.ID, "Body_lblConfirmHead")
            heading = (heading_el.text or "").strip()
        except Exception:
            pass

        close_buttons = driver.find_elements(By.ID, "Body_btnClose")
        if not close_buttons:
            close_buttons = popup.find_elements(
                By.CSS_SELECTOR, "input.popupbtnClose, input[id*='btnClose']"
            )
        if not close_buttons:
            return False

        close_button = close_buttons[0]
        try:
            close_button.click()
        except Exception:
            driver.execute_script("arguments[0].click();", close_button)

        label = f" ({heading})" if heading else ""
        print(f"Post-login notification popup detected{label}. Dismissed it.")

        try:
            WebDriverWait(driver, timeout_seconds).until(
                EC.invisibility_of_element(popup)
            )
        except Exception:
            try:
                driver.execute_script(
                    "arguments[0].style.display = 'none';", popup
                )
            except Exception:
                pass

        time.sleep(1)
        return True
    except Exception:
        return False


def wait_for_visible_element(
    driver,
    by,
    value,
    timeout_seconds: int = MAX_ELEMENT_WAIT_SECONDS,
    poll_seconds: int = POLL_SECONDS, 
   ):
    """Poll until an element is present and displayed, else request browser restart."""
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        try:
            elements = driver.find_elements(by, value)
            for elem in elements:
                if elem.is_displayed():
                    return elem
        except Exception:
            pass
        time.sleep(poll_seconds)
    raise RestartBrowserException(
        f"Element not visible within {timeout_seconds} seconds: by={by}, value={value}"
    )


def parse_date_yyyy_mm_dd(date_str: str) -> datetime:
    return datetime.strptime(date_str.strip(), "%Y-%m-%d")


def format_dd_mm_yyyy(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y")


def split_date_range_to_intervals(start_date: str, end_date: str, interval_days: int = INTERVAL_DAYS):
    """
    Split inclusive date range into non-overlapping intervals of max interval_days.
    Next interval starts at previous_end + 1 day.
    """
    if interval_days < 1:
        raise ValueError("interval_days must be at least 1")

    start_dt = parse_date_yyyy_mm_dd(start_date)
    end_dt = parse_date_yyyy_mm_dd(end_date)
    if end_dt < start_dt:
        raise ValueError("DATE_RANGE_END must be >= DATE_RANGE_START")

    intervals = []
    current = start_dt
    while current <= end_dt:
        interval_end = min(current + timedelta(days=interval_days - 1), end_dt)
        intervals.append((current, interval_end))
        current = interval_end + timedelta(days=1)
    return intervals


def load_intervals_from_excel(interval_excel_path: str) -> dict[str, list[tuple[datetime, datetime]]]:
    """Load plaza-specific intervals from Excel and group them by normalized plaza name."""
    interval_path = Path(interval_excel_path)
    if not interval_path.exists():
        raise FileNotFoundError(f"Interval Excel file not found: {interval_excel_path}")

    df = pd.read_excel(interval_path, engine="openpyxl")
    col_map = resolve_interval_columns(df)

    plaza_col = col_map["plaza_name"]
    start_col = col_map["start_date"]
    end_col = col_map["end_date"]

    grouped_intervals: dict[str, list[tuple[datetime, datetime]]] = {}
    for row_num, (_, row) in enumerate(df.iterrows(), start=2):
        plaza_name = str(row[plaza_col]).strip() if pd.notna(row[plaza_col]) else ""
        start_raw = row[start_col]
        end_raw = row[end_col]

        if not plaza_name or pd.isna(start_raw) or pd.isna(end_raw):
            print(f"Skipping interval Excel row {row_num} due to missing Plaza_name/Start_Date/End_Date.")
            continue

        start_dt = pd.to_datetime(start_raw).to_pydatetime()
        end_dt = pd.to_datetime(end_raw).to_pydatetime()
        start_dt = datetime(start_dt.year, start_dt.month, start_dt.day)
        end_dt = datetime(end_dt.year, end_dt.month, end_dt.day)
        if end_dt < start_dt:
            raise ValueError(
                f"Invalid interval Excel row {row_num}: End_Date earlier than Start_Date for plaza '{plaza_name}'."
            )

        grouped_intervals.setdefault(plaza_name.strip().lower(), []).append((start_dt, end_dt))

    for plaza_key in grouped_intervals:
        grouped_intervals[plaza_key].sort(key=lambda item: (item[0], item[1]))

    return grouped_intervals


def build_intervals_by_plaza(
    pending_rows: pd.DataFrame,
    plaza_col,
    interval_source_mode: str,
    date_range_start: str,
    date_range_end: str,
    interval_excel_path: str,
     ) -> dict[str, list[tuple[datetime, datetime]]]:
    """Resolve intervals to process for each plaza based on configured mode."""
    mode = interval_source_mode.strip().lower()
    if mode == "date_range":
        shared_intervals = split_date_range_to_intervals(
            date_range_start,
            date_range_end,
            interval_days=INTERVAL_DAYS,
        )
        intervals_by_plaza: dict[str, list[tuple[datetime, datetime]]] = {}
        for _, row in pending_rows.iterrows():
            plaza_name = str(row[plaza_col]).strip() if pd.notna(row[plaza_col]) else ""
            if plaza_name:
                intervals_by_plaza[plaza_name.strip().lower()] = list(shared_intervals)
        return intervals_by_plaza

    if mode == "excel":
        return load_intervals_from_excel(interval_excel_path)

    raise ValueError("INTERVAL_SOURCE_MODE must be either 'date_range' or 'excel'.")


def write_skipped_intervals_excel(skipped_intervals: list[dict], output_path: Path):
    """Write skipped intervals into an Excel file for later reruns."""
    columns = ["Plaza_name", "Start_Date", "End_Date", "Reason"]
    if skipped_intervals:
        df = pd.DataFrame(skipped_intervals, columns=columns)
    else:
        df = pd.DataFrame(columns=columns)
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"Skipped intervals saved to: {output_path}")


def get_plaza_output_dir(plaza_name: str) -> Path:
    """Create (if needed) and return plaza folder in the same directory as this script."""
    plaza_dir = SCRIPT_DIR / sanitize_filename_part(plaza_name)
    plaza_dir.mkdir(parents=True, exist_ok=True)
    return plaza_dir


def sanitize_filename_part(value: str) -> str:
    safe = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in value.strip())
    return safe.strip("_") or "unknown"


def navigate_to_transactions_reconciliation_page(driver):
    """
    Navigate via menu hierarchy:
    Reports -> Reconciliation Reports -> Transactions Reconciliation And Settlements
    """
    print("Navigating menu: Reports -> Reconciliation Reports -> Transactions Reconciliation And Settlements")
    handle_post_login_notification_popup(driver, timeout_seconds=2)
    wait_for_visible_element(driver, By.ID, "myslidemenu")

    reports_link = wait_for_visible_element(
        driver,
        By.XPATH,
        "//div[@id='myslidemenu']//a[normalize-space()='Reports']",
    )
    ActionChains(driver).move_to_element(reports_link).perform()
    time.sleep(1)

    reconciliation_reports_link = wait_for_visible_element(
        driver,
        By.XPATH,
        "//div[@id='myslidemenu']//a[contains(normalize-space(.), 'Reconciliation Reports')]",
    )
    ActionChains(driver).move_to_element(reconciliation_reports_link).perform()
    time.sleep(1)

    target_link = wait_for_visible_element(
        driver,
        By.XPATH,
        "//div[@id='myslidemenu']//a[contains(normalize-space(.), 'Transactions Reconciliation And Settlements')]",
    )
    target_link.click()

    print(f"Waiting for report page heading (max {MAX_ELEMENT_WAIT_SECONDS} seconds)...")
    wait_for_visible_element(driver, By.ID, "Heading_lblHeading")
    print("Transactions reconciliation page loaded.")


def set_date_inputs_and_submit(driver, start_dt: datetime, end_dt: datetime):
    """Set start/end dates and submit report form."""
    start_value = format_dd_mm_yyyy(start_dt)
    end_value = format_dd_mm_yyyy(end_dt)
    print(f"Submitting date interval: {start_value} to {end_value}")

    start_input = wait_for_visible_element(driver, By.ID, "Body_txtStartEffectiveDate")
    end_input = wait_for_visible_element(driver, By.ID, "Body_txtEndeffectiveDate")

    # Use JS set to avoid datepicker/input restrictions.
    driver.execute_script("arguments[0].value = arguments[1];", start_input, start_value)
    driver.execute_script("arguments[0].value = arguments[1];", end_input, end_value)
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));"
        "arguments[1].dispatchEvent(new Event('change', { bubbles: true }));",
        start_input,
        end_input,
    )

    submit_btn = wait_for_visible_element(driver, By.ID, "Body_btnSubmit")
    submit_btn.click()


def wait_for_export_button(driver):
    """Wait until export is ready or no-records banner appears, else restart browser."""
    print(f"Waiting for records and export button (max {MAX_REPORT_WAIT_SECONDS} seconds)...")
    export_button_id = "ctl00_Body_rvTransactionReconciliationSettlements_ctl05_ctl04_ctl00_ButtonLink"
    no_records_id = "Body_Nodata_lblNoRecords"
    deadline = time.time() + MAX_REPORT_WAIT_SECONDS

    while time.time() < deadline:
        try:
            export_buttons = driver.find_elements(By.ID, export_button_id)
            for elem in export_buttons:
                if elem.is_displayed():
                    return "export_ready"
        except Exception:
            pass

        try:
            no_record_labels = driver.find_elements(By.ID, no_records_id)
            for elem in no_record_labels:
                if elem.is_displayed():
                    text = (elem.text or "").strip().lower()
                    if "no records to display" in text:
                        print("No records banner detected for this interval. Skipping export.")
                        return "no_records"
        except Exception:
            pass

        time.sleep(POLL_SECONDS)

    raise RestartBrowserException(
        f"Report results did not load within {MAX_REPORT_WAIT_SECONDS} seconds."
    )


def click_export_excel(driver):
    """Open export menu from ButtonLink, verify menu opens, then click Excel."""
    # Ensure report/export widget is present.
    wait_result = wait_for_export_button(driver)
    if wait_result == "no_records":
        return False

    button_link_id = "ctl00_Body_rvTransactionReconciliationSettlements_ctl05_ctl04_ctl00_ButtonLink"
    menu_id = "ctl00_Body_rvTransactionReconciliationSettlements_ctl05_ctl04_ctl00_Menu"

    button_link = wait_for_visible_element(driver, By.ID, button_link_id)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", button_link)
    except Exception:
        pass

    def is_menu_open() -> bool:
        try:
            menu = driver.find_element(By.ID, menu_id)
            style = (menu.get_attribute("style") or "").lower()
            if "display: none" in style:
                return False
            if "visibility: hidden" in style:
                return False
            return menu.is_displayed()
        except Exception:
            return False

    # Try multiple click methods on the exact ButtonLink.
    for _ in range(8):
        if is_menu_open():
            break
        try:
            button_link.click()
        except Exception:
            pass
        if is_menu_open():
            break
        try:
            ActionChains(driver).move_to_element(button_link).click(button_link).perform()
        except Exception:
            pass
        if is_menu_open():
            break
        try:
            driver.execute_script("arguments[0].click();", button_link)
        except Exception:
            pass
        if is_menu_open():
            break
        try:
            # Dispatch full mouse events to emulate manual click.
            driver.execute_script(
                "var e=arguments[0];"
                "['mouseover','mousedown','mouseup','click'].forEach(function(t){"
                "e.dispatchEvent(new MouseEvent(t,{view:window,bubbles:true,cancelable:true}));"
                "});",
                button_link,
            )
        except Exception:
            pass
        time.sleep(0.5)

    # Block until menu is actually open (no timeout).
    menu_deadline = time.time() + MAX_ELEMENT_WAIT_SECONDS
    while not is_menu_open():
        if time.time() >= menu_deadline:
            raise RestartBrowserException(
                f"Export menu did not open within {MAX_ELEMENT_WAIT_SECONDS} seconds."
            )
        time.sleep(POLL_SECONDS)

    # Click Excel option with robust fallbacks.
    excel_option = None
    excel_xpaths = [
        "//a[@title='Excel' and contains(@onclick, 'EXCELOPENXML')]",
        "//div[@id='ctl00_Body_rvTransactionReconciliationSettlements_ctl05_ctl04_ctl00_Menu']//a[@title='Excel']",
        "//a[contains(@onclick, \"exportReport('EXCELOPENXML')\")]",
    ]
    for xp in excel_xpaths:
        elems = driver.find_elements(By.XPATH, xp)
        for elem in elems:
            if elem.is_displayed():
                excel_option = elem
                break
        if excel_option is not None:
            break

    if excel_option is None:
        excel_option = wait_for_visible_element(
            driver,
            By.XPATH,
            "//a[@title='Excel']",
        )

    try:
        excel_option.click()
    except Exception:
        driver.execute_script("arguments[0].click();", excel_option)

    print("Excel export triggered.")
    return True


def detect_recent_download_file(downloads_folder: str, max_age_minutes: int = 10):
    """Detect latest .csv/.xlsx file modified within max_age_minutes."""
    folder = Path(downloads_folder)
    if not folder.exists():
        return None

    cutoff = time.time() - (max_age_minutes * 60)
    candidates = []
    for pattern in ("*.csv", "*.xlsx"):
        for path in folder.glob(pattern):
            try:
                mtime = path.stat().st_mtime
                if mtime >= cutoff:
                    candidates.append(path)
            except Exception:
                continue

    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def wait_and_move_downloaded_file(
    plaza_name: str,
    start_dt: datetime,
    end_dt: datetime,
    output_dir: Path,
    downloads_folder: str = DOWNLOADS_FOLDER,
    ):
    """Wait for a recent download and move into the plaza output folder."""
    print(f"Waiting for downloaded .csv/.xlsx file (max {MAX_ELEMENT_WAIT_SECONDS} seconds)...")
    downloaded = None
    deadline = time.time() + MAX_ELEMENT_WAIT_SECONDS
    while downloaded is None and time.time() < deadline:
        downloaded = detect_recent_download_file(downloads_folder, max_age_minutes=10)
        if downloaded is None:
            time.sleep(POLL_SECONDS)

    if downloaded is None:
        raise RestartBrowserException(
            f"Downloaded file not found within {MAX_ELEMENT_WAIT_SECONDS} seconds."
        )

    safe_plaza = sanitize_filename_part(plaza_name)
    start_s = start_dt.strftime("%Y-%m-%d")
    end_s = end_dt.strftime("%Y-%m-%d")
    new_name = f"{safe_plaza}_{start_s}_to_{end_s}{downloaded.suffix.lower()}"
    target_path = output_dir / new_name

    if target_path.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_path = output_dir / f"{safe_plaza}_{start_s}_to_{end_s}_{stamp}{downloaded.suffix.lower()}"

    shutil.move(str(downloaded), str(target_path))
    print(f"Moved download: {target_path}")
    return str(target_path)


def login_with_retries(driver, username: str, password: str, max_attempts: int = MAX_LOGIN_ATTEMPTS) -> bool:
    """Attempt sign-in multiple times by refreshing captcha only."""
    driver.get(LOGIN_URL)

    username_field = wait_for_visible_element(driver, By.ID, "txtUserName")
    password_field = wait_for_visible_element(driver, By.ID, "txtPassword")
    captcha_field = wait_for_visible_element(driver, By.ID, "txtCaptcha")
    sign_in_button = wait_for_visible_element(driver, By.ID, "btnSignIn")

    # Enter username and password once, as requested.
    username_field.clear()
    username_field.send_keys(username)
    password_field.clear()
    password_field.send_keys(password)

    for attempt in range(1, max_attempts + 1):
        print(f"\nLogin attempt {attempt}/{max_attempts}")
        try:
            captcha_text = solve_captcha_text(driver)
            if not captcha_text:
                print("Captcha OCR returned empty text; retrying.")
                continue

            print(f"OCR captcha text: {captcha_text}")
            captcha_field = wait_for_visible_element(driver, By.ID, "txtCaptcha")
            captcha_field.clear()
            captcha_field.send_keys(captcha_text)

            sign_in_button = wait_for_visible_element(driver, By.ID, "btnSignIn")
            sign_in_button.click()
            handle_duplicate_login_popup(driver)

            # Wait for either login form persistence or post-login transition.
            time.sleep(3)
            if is_login_successful(driver):
                print("Sign-in successful.")
                handle_post_login_notification_popup(driver)
                return True

            print("Sign-in failed; still on login page.")
            print_login_errors_if_any(driver)

        except Exception as exc:
            print(f"Attempt {attempt} failed with exception: {exc}")

    return False


def process_reference_file(
    excel_file_path: str = REFERENCE_EXCEL,
    headless: bool = False,
    date_range_start: str | None = None,
    date_range_end: str | None = None,
    interval_source_mode: str = INTERVAL_SOURCE_MODE,
    interval_excel_path: str = INTERVAL_LIST_EXCEL,
    skipped_intervals_excel_path: str = SKIPPED_INTERVALS_EXCEL,
  ) -> bool:
    """Read rows from reference Excel and perform row-wise login attempts."""
    if date_range_start is None or date_range_end is None:
        date_range_start, date_range_end = get_process_date_range_strings()

    reference_path = Path(excel_file_path)
    if not reference_path.exists():
        print(f"Reference file not found: {excel_file_path}")
        return False

    df = pd.read_excel(reference_path, engine="openpyxl")
    try:
        col_map = resolve_required_columns(df)
    except ValueError as exc:
        print(str(exc))
        return False

    plaza_col = col_map["plaza_name"]
    username_col = col_map["username"]
    password_col = col_map["password"]
    done_col = col_map["done"]

    # Pending rows: Done empty or not equal to "done" (case-insensitive).
    pending_rows = df[
        (df[done_col].isna())
        | (df[done_col].astype(str).str.strip().str.lower() != "done")
    ]

    print(f"Total rows: {len(df)}")
    print(f"Pending rows: {len(pending_rows)}")

    if pending_rows.empty:
        print("No pending rows to process.")
        return True

    try:
        intervals_by_plaza = build_intervals_by_plaza(
            pending_rows,
            plaza_col,
            interval_source_mode,
            date_range_start,
            date_range_end,
            interval_excel_path,
        )
    except Exception as exc:
        print(f"Failed to resolve intervals: {exc}")
        return False

    if interval_source_mode.strip().lower() == "date_range":
        sample_intervals = next(iter(intervals_by_plaza.values()), [])
        print(
            f"Using {len(sample_intervals)} interval(s) of {INTERVAL_DAYS} day(s) "
            f"from {date_range_start} to {date_range_end}"
        )
        for i, (s, e) in enumerate(sample_intervals, start=1):
            print(f"  Interval {i}: {s.strftime('%Y-%m-%d')} to {e.strftime('%Y-%m-%d')}")
    else:
        total_interval_count = sum(len(items) for items in intervals_by_plaza.values())
        print(
            f"Using interval Excel source: {interval_excel_path} "
            f"with {total_interval_count} interval(s) across {len(intervals_by_plaza)} plaza(s)"
        )

    success_count = 0
    fail_count = 0
    skipped_intervals: list[dict] = []

    for idx, row in pending_rows.iterrows():
        plaza_name = str(row[plaza_col]).strip() if pd.notna(row[plaza_col]) else ""
        username = str(row[username_col]).strip() if pd.notna(row[username_col]) else ""
        password = str(row[password_col]).strip() if pd.notna(row[password_col]) else ""
        plaza_intervals = intervals_by_plaza.get(plaza_name.lower(), [])

        print("\n" + "=" * 80)
        print(f"Processing row {idx + 2}")
        print(f"Plaza: {plaza_name}")
        print(f"Username: {username}")
        print("=" * 80)

        if not plaza_name or not username or not password:
            print("Skipping row due to missing Plaza_name/Username/Password.")
            fail_count += 1
            continue

        if not plaza_intervals:
            print(f"No intervals configured for plaza: {plaza_name}. Skipping this row.")
            continue

        driver = None
        interval_index = 0
        row_completed = False
        plaza_output_dir = get_plaza_output_dir(plaza_name)
        print(f"Output folder: {plaza_output_dir}")

        try:
            while interval_index < len(plaza_intervals):
                try:
                    driver = setup_driver(headless=headless)
                    login_ok = login_with_retries(driver, username=username, password=password, max_attempts=MAX_LOGIN_ATTEMPTS)
                    if not login_ok:
                        print(f"Login failed after {MAX_LOGIN_ATTEMPTS} attempts for plaza: {plaza_name}")
                        fail_count += 1
                        break

                    navigate_to_transactions_reconciliation_page(driver)

                    while interval_index < len(plaza_intervals):
                        interval_start, interval_end = plaza_intervals[interval_index]
                        set_date_inputs_and_submit(driver, interval_start, interval_end)
                        export_triggered = click_export_excel(driver)
                        if not export_triggered:
                            skipped_intervals.append(
                                {
                                    "Plaza_name": plaza_name,
                                    "Start_Date": interval_start.strftime("%Y-%m-%d"),
                                    "End_Date": interval_end.strftime("%Y-%m-%d"),
                                    "Reason": "No records to display",
                                }
                            )
                            interval_index += 1
                            time.sleep(2)
                            continue

                        wait_and_move_downloaded_file(
                            plaza_name,
                            interval_start,
                            interval_end,
                            output_dir=plaza_output_dir,
                            downloads_folder=DOWNLOADS_FOLDER,
                        )
                        interval_index += 1
                        time.sleep(2)

                    if interval_index >= len(plaza_intervals):
                        row_completed = True
                        break

                except RestartBrowserException as exc:
                    current_start, current_end = plaza_intervals[interval_index]
                    print(
                        "Restarting browser for same interval "
                        f"{current_start.strftime('%Y-%m-%d')} to {current_end.strftime('%Y-%m-%d')}: {exc}"
                    )
                finally:
                    if driver:
                        try:
                            driver.quit()
                        except Exception:
                            pass
                        driver = None

            if row_completed:
                mark_done_in_excel(excel_file_path, idx, done_col)
                success_count += 1
            elif interval_index >= len(plaza_intervals):
                fail_count += 1

        except Exception as exc:
            print(f"Unexpected error while processing row {idx + 2}: {exc}")
            fail_count += 1

    print("\n" + "=" * 80)
    print("ICICI LOGIN SUMMARY")
    print("=" * 80)
    print(f"Successful: {success_count}")
    print(f"Failed: {fail_count}")
    print("=" * 80)
    write_skipped_intervals_excel(skipped_intervals, SCRIPT_DIR / skipped_intervals_excel_path)
    return fail_count == 0


def main():
    try:
        date_range_start, date_range_end = get_process_date_range_strings()
    except (RuntimeError, ValueError) as exc:
        print(f"\nError: {exc}")
        return

    print(f"Using process date range from .env: {date_range_start} to {date_range_end}")
    process_reference_file(
        excel_file_path=REFERENCE_EXCEL,
        headless=False,
        date_range_start=date_range_start,
        date_range_end=date_range_end,
        interval_source_mode=INTERVAL_SOURCE_MODE,
        interval_excel_path=INTERVAL_LIST_EXCEL,
        skipped_intervals_excel_path=SKIPPED_INTERVALS_EXCEL,
    )


if __name__ == "__main__":
    main()
