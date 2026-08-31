from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import requests
import numpy as np
from PIL import Image
import io
from ocr_module import extract_text, special_character_remover
import os
import subprocess
from pathlib import Path
from datetime import datetime
import shutil
import pandas as pd
import tempfile


IDFC_PORTAL_ORIGIN = "https://etoll.idfcbank.com"
IDFC_LOGIN_URL = f"{IDFC_PORTAL_ORIGIN}/dimtspay_concess/login"
IDFC_SESSION_RELEASE_WAIT_SECONDS = 45
_IDFC_LOGOUT_PATHS = (
    "/dimtspay_concess/logout",
    "/dimtspay_concess/j_spring_security_logout",
    "/dimtspay_concess/signout",
)
_IDFC_SESSION_LIMIT_MARKERS = (
    "maximum sessions",
    "max sessions",
    "session limit",
    "session of 1",
    "sessions of 1",
    "principal exceeded",
    "already logged in",
    "concurrent session",
)


def setup_driver(headless=False, fresh_profile=True):
    """
    Setup and return Chrome WebDriver.

    fresh_profile: use a new temp user-data-dir and incognito so no prior
    portal cookies/storage leak between plaza runs.
    """
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    profile_dir = None
    if fresh_profile:
        # Isolated temp profile (no shared Chrome cookies/cache with other runs).
        profile_dir = tempfile.mkdtemp(prefix="idfc_chrome_profile_")
        chrome_options.add_argument(f"--user-data-dir={profile_dir}")

    chrome_options.add_argument("--disable-application-cache")
    chrome_options.add_experimental_option(
        "prefs",
        {
            "profile.default_content_setting_values.cookies": 1,
            "profile.cookie_controls_mode": 0,
        },
    )

    driver = webdriver.Chrome(options=chrome_options)
    if profile_dir:
        driver._idfc_chrome_profile_dir = profile_dir
    return driver


def is_idfc_session_limit_message(text: str) -> bool:
    """Return True when page/error text indicates the one-session-per-user limit."""
    normalized = (text or "").strip().lower()
    return any(marker in normalized for marker in _IDFC_SESSION_LIMIT_MARKERS)


def get_idfc_visible_page_text(driver) -> str:
    """Collect visible text from the current page for error detection."""
    chunks: list[str] = []
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        chunks.append(body.text or "")
    except Exception:
        pass
    for selector in (".error", ".alert", "[class*='error']", "[class*='alert']"):
        try:
            for elem in driver.find_elements(By.CSS_SELECTOR, selector):
                if elem.is_displayed():
                    chunks.append(elem.text or "")
        except Exception:
            continue
    return "\n".join(chunks)


def wait_for_idfc_session_slot(
    seconds: int = IDFC_SESSION_RELEASE_WAIT_SECONDS,
    reason: str = "",
    ):
    """Wait so the portal can drop the previous server-side session."""
    label = f" ({reason})" if reason else ""
    print(
        f"Waiting {seconds}s for IDFC portal to release the previous session{label}..."
    )
    time.sleep(seconds)


def _clear_idfc_browser_storage(driver):
    """Clear cookies, cache, and web storage for the portal origin."""
    try:
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.clearBrowserCookies", {})
        driver.execute_cdp_cmd("Network.clearBrowserCache", {})
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd(
            "Storage.clearDataForOrigin",
            {"origin": IDFC_PORTAL_ORIGIN, "storageTypes": "all"},
        )
    except Exception:
        pass
    try:
        driver.execute_script(
            """
            try { localStorage.clear(); } catch (e) {}
            try { sessionStorage.clear(); } catch (e) {}
            """
        )
    except Exception:
        pass
    try:
        driver.delete_all_cookies()
    except Exception:
        pass


def release_idfc_browser_session(driver):
    """
    Ask the portal to end the server-side session and wipe local state.
    IDFC allows only one concurrent session per user; closing Chrome alone
    often leaves that session active, so the next login fails.
    """
    if not driver:
        return
    print("Releasing IDFC portal session (logout + clearing storage/cookies)...")
    try:
        driver.set_page_load_timeout(25)
    except Exception:
        pass
    try:
        driver.get(IDFC_LOGIN_URL)
        time.sleep(1)
    except Exception:
        pass
    for path in _IDFC_LOGOUT_PATHS:
        try:
            driver.get(IDFC_PORTAL_ORIGIN + path)
            time.sleep(1.5)
        except Exception:
            pass
    try:
        for link_text in ("Logout", "Log out", "Sign out", "Sign Out"):
            links = driver.find_elements(
                By.XPATH,
                f"//a[contains(normalize-space(.), '{link_text}')]",
            )
            for link in links:
                if link.is_displayed():
                    link.click()
                    time.sleep(1.5)
                    break
    except Exception:
        pass
    _clear_idfc_browser_storage(driver)


def prepare_idfc_login_page(driver):
    """Clear stale portal state and open a fresh login page."""
    release_idfc_browser_session(driver)
    try:
        driver.get(IDFC_LOGIN_URL)
        time.sleep(1)
        _clear_idfc_browser_storage(driver)
        driver.get(IDFC_LOGIN_URL)
        time.sleep(1)
    except Exception:
        pass


def probe_release_stale_idfc_session(headless: bool = False):
    """
    Open a short-lived browser to hit logout/clear endpoints before a real run.
    Helps when a prior run ended without shutdown_idfc_driver().
    """
    print("Probing IDFC portal to release any stale server-side session...")
    driver = None
    try:
        driver = setup_driver(headless=headless, fresh_profile=True)
        release_idfc_browser_session(driver)
    except Exception as exc:
        print(f"Session probe warning: {exc}")
    finally:
        quit_driver_and_cleanup_profile(driver)
    wait_for_idfc_session_slot(reason="after session probe")


def quit_driver_and_cleanup_profile(driver):
    """Quit Chrome and remove the temporary profile directory if we created one."""
    if not driver:
        return
    profile_dir = getattr(driver, "_idfc_chrome_profile_dir", None)
    try:
        driver.quit()
    except Exception:
        pass
    if profile_dir and os.path.isdir(profile_dir):
        shutil.rmtree(profile_dir, ignore_errors=True)


def shutdown_idfc_driver(driver):
    """Release portal session, quit browser, delete temp profile."""
    if not driver:
        return
    try:
        release_idfc_browser_session(driver)
    except Exception:
        pass
    quit_driver_and_cleanup_profile(driver)


def get_captcha_image(driver, captcha_element):
    """
    Get captcha image from the element and convert to numpy array.
    Returns numpy array of the image.
    """
    try:
        # Method 1: Try to get image from src URL with session cookies
        captcha_src = captcha_element.get_attribute("src")
        
        if captcha_src:
            # Construct full URL if relative
            if captcha_src.startswith("/"):
                # Relative URL - construct full URL
                base_url = "https://etoll.idfcbank.com"
                image_url = base_url + captcha_src
            elif captcha_src.startswith("http"):
                image_url = captcha_src
            else:
                # Fallback to screenshot method
                return get_captcha_from_screenshot(driver, captcha_element)
            
            # Create a session and add cookies
            session = requests.Session()
            selenium_cookies = driver.get_cookies()
            for cookie in selenium_cookies:
                session.cookies.set(cookie['name'], cookie['value'])
            
            # Download the image
            headers = {
                'User-Agent': driver.execute_script("return navigator.userAgent;"),
                'Referer': driver.current_url
            }
            response = session.get(image_url, headers=headers, timeout=10)
            
            if response.status_code == 200:
                # Convert to PIL Image then to numpy array
                img = Image.open(io.BytesIO(response.content))
                # Convert to RGB if needed
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img_array = np.array(img)
                return img_array
            else:
                print(f"Failed to download image. Status code: {response.status_code}")
                # Fallback to screenshot method
                return get_captcha_from_screenshot(driver, captcha_element)
        else:
            # No src attribute, use screenshot method
            return get_captcha_from_screenshot(driver, captcha_element)
            
    except Exception as e:
        print(f"Error getting captcha from URL: {e}")
        # Fallback to screenshot method
        return get_captcha_from_screenshot(driver, captcha_element)


def get_captcha_from_screenshot(driver, captcha_element):
    """Take screenshot of captcha element and convert to numpy array"""
    try:
        # Take screenshot of the element
        import os
        temp_file = "temp_captcha.png"
        captcha_element.screenshot(temp_file)
        
        # Read the image and convert to numpy array
        img = Image.open(temp_file)
        # Convert to RGB if needed (screenshots might be RGBA)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        img_array = np.array(img)
        
        # Clean up temp file
        if os.path.exists(temp_file):
            os.remove(temp_file)
        
        return img_array
    except Exception as e:
        print(f"Error taking screenshot: {e}")
        raise


def handle_popup(driver, wait):
    """
    Check for popup and click 'No' button if it appears.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    """
    try:
        print("Checking for popup...")
        
        # Wait for popup to appear (with short timeout, as it may or may not appear)
        try:
            # Wait up to 5 seconds for popup to appear
            popup_wait = WebDriverWait(driver, 5)
            popup = popup_wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".jconfirm-box"))
            )
            
            # Check if popup is visible
            if popup.is_displayed():
                print("Popup detected. Looking for 'No' button...")
                
                # Find the 'No' button within the popup
                # The button has class "btn btn-default" and text "No"
                no_buttons = popup.find_elements(By.XPATH, ".//button[contains(@class, 'btn-default') and contains(text(), 'No')]")
                
                if not no_buttons:
                    # Alternative: try finding by class only
                    no_buttons = popup.find_elements(By.CSS_SELECTOR, "button.btn.btn-default")
                
                if no_buttons:
                    print("Clicking 'No' button to close popup...")
                    no_button = no_buttons[0]
                    # Wait for button to be clickable
                    popup_wait.until(EC.element_to_be_clickable(no_button))
                    no_button.click()
                    time.sleep(1)
                    print("Popup closed successfully.")
                    return True
                else:
                    print("Warning: Popup found but 'No' button not found.")
        except:
            # Popup didn't appear within timeout - this is fine
            print("No popup detected (or popup didn't appear within timeout).")
        
        return False
    except Exception as e:
        print(f"Error handling popup: {e}")
        # Don't fail the whole process if popup handling fails
        return False


def click_concessionaire_button(driver, wait):
    """
    Click on the Concessionaire button/link.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    """
    try:
        print("Looking for Concessionaire button...")
        
        # Wait for the Concessionaire button to be present
        concessionaire_button = wait.until(
            EC.presence_of_element_located((By.ID, "concessionairebtn"))
        )
        
        # Wait for it to be clickable
        wait.until(EC.element_to_be_clickable((By.ID, "concessionairebtn")))
        
        print("Clicking Concessionaire button...")
        concessionaire_button.click()
        time.sleep(2)  # Wait for navigation/action to complete
        
        print("Concessionaire button clicked successfully.")
        return True
        
    except Exception as e:
        print(f"Error clicking Concessionaire button: {e}")
        import traceback
        traceback.print_exc()
        return False


def click_life_cycle_report(driver, wait):
    """
    Click on the 'Life Cycle Report(New)' link.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    """
    try:
        print("Looking for Life Cycle Report(New) link...")
        
        # Wait for the link to be present - using partial text match
        life_cycle_link = wait.until(
            EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Life Cycle Report(New)"))
        )
        
        # Wait for it to be clickable
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Life Cycle Report(New)")))
        
        print("Clicking Life Cycle Report(New) link...")
        life_cycle_link.click()
        time.sleep(2)  # Wait for page to load
        
        print("Life Cycle Report(New) link clicked successfully.")
        return True
        
    except Exception as e:
        print(f"Error clicking Life Cycle Report link: {e}")
        import traceback
        traceback.print_exc()
        return False


def select_date_type(driver, wait, date_type="Reader Read Date"):
    """
    Click on the date type dropdown and select the specified option.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
        date_type: The date type to select (default: "Reader Read Date")
    """
    try:
        print(f"Looking for date type dropdown...")
        
        # Wait for the dropdown to be present - using class name
        dropdown = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.chzn-single.chzn-default"))
        )
        
        # Wait for it to be clickable
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.chzn-single.chzn-default")))
        
        print("Clicking date type dropdown...")
        dropdown.click()
        time.sleep(1)  # Wait for dropdown to open
        
        # Wait for dropdown options to appear
        print(f"Waiting for dropdown options to appear...")
        options_wait = WebDriverWait(driver, 5)
        
        # Try multiple ways to find the option
        try:
            # Method 1: XPath with exact text match
            date_type_option = options_wait.until(
                EC.element_to_be_clickable((By.XPATH, f"//li[@class='active-result' and normalize-space(text())='{date_type}']"))
            )
        except:
            # Method 2: XPath with contains
            try:
                date_type_option = options_wait.until(
                    EC.element_to_be_clickable((By.XPATH, f"//li[@class='active-result' and contains(text(), '{date_type}')]"))
                )
            except:
                # Method 3: Find all options and match by text
                options = driver.find_elements(By.CSS_SELECTOR, "ul.chzn-results li.active-result")
                date_type_option = None
                for option in options:
                    if date_type in option.text:
                        date_type_option = option
                        break
                if not date_type_option:
                    raise Exception(f"Could not find option with text '{date_type}'")
        
        print(f"Selecting '{date_type}' from dropdown...")
        date_type_option.click()
        time.sleep(1)  # Wait for selection to complete
        
        print(f"Date type '{date_type}' selected successfully.")
        return True
        
    except Exception as e:
        print(f"Error selecting date type: {e}")
        import traceback
        traceback.print_exc()
        return False


def open_date_range_picker(driver, wait):
    """
    Click on the date range input field to open the calendar pickers.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print("Looking for date range input field...")
        
        # Wait for the date range input to be present
        date_range_input = wait.until(
            EC.presence_of_element_located((By.ID, "dateRange"))
        )
        
        # Wait for it to be clickable
        wait.until(EC.element_to_be_clickable((By.ID, "dateRange")))
        
        print("Clicking date range input field to open calendars...")
        date_range_input.click()
        time.sleep(1)  # Wait for calendars to appear
        
        print("Date range picker opened successfully.")
        return True
        
    except Exception as e:
        print(f"Error opening date range picker: {e}")
        import traceback
        traceback.print_exc()
        return False


def get_from_date(driver, wait):
    """
    Read the 'from' date from the date range picker.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    
    Returns:
        datetime object or None if failed
    """
    try:
        from datetime import datetime
        
        print("Reading 'from' date from date range picker...")
        
        # Wait for the calendar to be visible
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".calendar.left")))
        
        # Find the "from" input field - it's typically the first input in the date range picker
        # The date range picker usually has two inputs: from and to (both may be disabled)
        # We need to find the input fields within the date range picker
        try:
            # Method 1: Try to find disabled input fields that might contain "from" date
            # Look for input fields within the date range picker container
            try:
                # Try to find input fields by looking for disabled inputs or inputs with specific classes
                # The date range picker might have separate inputs for from and to
                all_inputs = driver.find_elements(By.CSS_SELECTOR, ".daterangepicker input[readonly], .daterangepicker input[disabled]")
                
                if all_inputs and len(all_inputs) >= 1:
                    # First input is usually "from" date
                    from_input = all_inputs[0]
                    from_date_str = from_input.get_attribute("value")
                    
                    if from_date_str:
                        print(f"Found 'from' date in input field: {from_date_str}")
                        # Parse the date - format is dd/mm/yyyy: "12/01/2026 12:00 AM"
                        # This means day=12, month=01, year=2026 (January 12, 2026)
                        try:
                            from_date = datetime.strptime(from_date_str, "%d/%m/%Y %I:%M %p")
                            print(f"Parsed 'from' date (dd/mm/yyyy): {from_date.strftime('%Y-%m-%d')} (Month: {from_date.month}, Day: {from_date.day})")
                            return from_date
                        except ValueError:
                            # Try alternative format mm/dd/yyyy as fallback
                            try:
                                from_date = datetime.strptime(from_date_str, "%m/%d/%Y %I:%M %p")
                                print(f"Parsed 'from' date (mm/dd/yyyy fallback): {from_date.strftime('%Y-%m-%d')} (Month: {from_date.month}, Day: {from_date.day})")
                                return from_date
                            except ValueError:
                                print(f"Could not parse date format: {from_date_str}")
            except:
                pass
            
            # Method 2: Try the dateRange input value
            date_range_input = driver.find_element(By.ID, "dateRange")
            date_value = date_range_input.get_attribute("value")
            
            if date_value:
                # Parse the date value - format might be "12/01/2026 12:00 AM - 12/31/2026 11:59 PM"
                # or just the from date
                print(f"Date range value: {date_value}")
                
                # Try to extract just the from date part
                if " - " in date_value:
                    from_date_str = date_value.split(" - ")[0].strip()
                else:
                    from_date_str = date_value.strip()
                
                # Parse the date - format is dd/mm/yyyy: "12/01/2026 12:00 AM"
                # This means day=12, month=01, year=2026 (January 12, 2026)
                try:
                    from_date = datetime.strptime(from_date_str, "%d/%m/%Y %I:%M %p")
                    print(f"Parsed 'from' date (dd/mm/yyyy): {from_date.strftime('%Y-%m-%d')} (Month: {from_date.month}, Day: {from_date.day})")
                    return from_date
                except ValueError:
                    # Try alternative format mm/dd/yyyy as fallback
                    try:
                        from_date = datetime.strptime(from_date_str, "%m/%d/%Y %I:%M %p")
                        print(f"Parsed 'from' date (mm/dd/yyyy fallback): {from_date.strftime('%Y-%m-%d')} (Month: {from_date.month}, Day: {from_date.day})")
                        return from_date
                    except ValueError:
                        print(f"Could not parse date format: {from_date_str}")
                        return None
            else:
                print("Date range input is empty")
                return None
                
        except Exception as e:
            print(f"Error reading from date: {e}")
            return None
            
    except Exception as e:
        print(f"Error getting from date: {e}")
        import traceback
        traceback.print_exc()
        return None


def calculate_months_to_move(from_date, target_date):
    """
    Calculate how many months to move in the calendar.
    
    Args:
        from_date: datetime object of the current 'from' date
        target_date: datetime object of the target date
    
    Returns:
        int: Number of months to move (negative for backward, positive for forward)
    """
    if not from_date or not target_date:
        return 0
    
    # Calculate the difference in months
    months_diff = (target_date.year - from_date.year) * 12 + (target_date.month - from_date.month)
    
    print(f"Calculated months difference: {months_diff} (from {from_date.strftime('%m/%Y')} to {target_date.strftime('%m/%Y')})")
    
    return months_diff


def select_date_in_calendar(driver, wait, target_date, calendar_side="left"):
    """
    Select a specific date in the calendar.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
        target_date: datetime object of the date to select
        calendar_side: Which calendar to use ("left" or "right")
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        from datetime import datetime
        
        day = target_date.day
        month = target_date.month
        year = target_date.year
        
        print(f"Selecting date: {day}/{month}/{year} in {calendar_side} calendar...")
        
        # Wait for the calendar to be present
        calendar_selector = f".calendar.{calendar_side}"
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, calendar_selector)))
        
        # Verify we're on the correct month/year
        month_year_element = driver.find_element(By.CSS_SELECTOR, f"{calendar_selector} .month")
        current_month_year = month_year_element.text.strip()
        
        # Parse current month/year (format: "Jan 2026")
        try:
            current_date = datetime.strptime(current_month_year, "%b %Y")
            if current_date.month != month or current_date.year != year:
                print(f"Warning: Calendar shows {current_month_year}, but target is {month}/{year}")
                print("Calendar might not be on the correct month. Please navigate first.")
        except:
            print(f"Could not parse current month/year: {current_month_year}")
        
        # Find and click the date
        # The date cells have class "available" and contain the day number
        # We need to find the cell with the correct day that's not "off" (disabled)
        date_cell = None
        
        # Try to find the date cell
        date_cells = driver.find_elements(By.CSS_SELECTOR, f"{calendar_selector} td.available")
        
        for cell in date_cells:
            cell_text = cell.text.strip()
            # Check if this cell has the day we want and is not disabled
            if cell_text == str(day) and "off" not in cell.get_attribute("class"):
                date_cell = cell
                break
        
        if date_cell:
            print(f"Clicking date cell for day {day}...")
            date_cell.click()
            time.sleep(0.5)
            print(f"Date {day}/{month}/{year} selected successfully.")
            return True
        else:
            print(f"Could not find date cell for day {day} in {calendar_side} calendar.")
            return False
        
    except Exception as e:
        print(f"Error selecting date in calendar: {e}")
        import traceback
        traceback.print_exc()
        return False


def get_to_date(driver, wait):
    """
    Read the 'to' date from the date range picker.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    
    Returns:
        datetime object or None if failed
    """
    try:
        from datetime import datetime
        
        print("Reading 'to' date from date range picker...")
        
        # Wait for the calendar to be visible
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".calendar.right")))
        
        # Find the "to" input field by name
        try:
            # Method 1: Try to find the input field by name
            to_input = driver.find_element(By.NAME, "daterangepicker_end")
            to_date_str = to_input.get_attribute("value")
            
            if to_date_str:
                print(f"Found 'to' date in input field: {to_date_str}")
                # Parse the date - format is dd/mm/yyyy: "12/01/2026 12:00 AM"
                try:
                    to_date = datetime.strptime(to_date_str, "%d/%m/%Y %I:%M %p")
                    print(f"Parsed 'to' date (dd/mm/yyyy): {to_date.strftime('%Y-%m-%d')} (Month: {to_date.month}, Day: {to_date.day})")
                    return to_date
                except ValueError:
                    # Try alternative format mm/dd/yyyy as fallback
                    try:
                        to_date = datetime.strptime(to_date_str, "%m/%d/%Y %I:%M %p")
                        print(f"Parsed 'to' date (mm/dd/yyyy fallback): {to_date.strftime('%Y-%m-%d')} (Month: {to_date.month}, Day: {to_date.day})")
                        return to_date
                    except ValueError:
                        print(f"Could not parse date format: {to_date_str}")
                        return None
            else:
                print("'To' date input field is empty")
                return None
                
        except Exception as e:
            print(f"Error reading to date: {e}")
            return None
            
    except Exception as e:
        print(f"Error getting to date: {e}")
        import traceback
        traceback.print_exc()
        return None


def navigate_calendar_months(driver, wait, months_to_move, calendar_side="left"):
    """
    Navigate the calendar by moving months backward or forward.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
        months_to_move: Number of months to move (negative for backward, positive for forward)
        calendar_side: Which calendar to navigate ("left" or "right")
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if months_to_move == 0:
            print("No months to move.")
            return True
        
        direction = "backward" if months_to_move < 0 else "forward"
        abs_months = abs(months_to_move)
        
        print(f"Moving {calendar_side} calendar {direction} by {abs_months} month(s)...")
        
        # Wait for the calendar to be present
        calendar_selector = f".calendar.{calendar_side}"
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, calendar_selector)))
        
        # Determine which arrow to click
        if months_to_move < 0:
            # Move backward - click left arrow
            arrow_selector = f"{calendar_selector} .prev.available"
        else:
            # Move forward - click right arrow
            arrow_selector = f"{calendar_selector} .next.available"
        
        # Check if the arrow exists (especially for right calendar, left arrow might not be available)
        try:
            arrow = driver.find_element(By.CSS_SELECTOR, arrow_selector)
        except:
            if months_to_move < 0 and calendar_side == "right":
                print(f"Left arrow not available in right calendar. Skipping navigation.")
                return True  # Return True since we can't navigate, but that's okay
            else:
                print(f"Arrow not found: {arrow_selector}")
                return False
        
        # Click the arrow the required number of times
        for i in range(abs_months):
            try:
                arrow = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, arrow_selector)))
                arrow.click()
                time.sleep(0.5)  # Wait for calendar to update
                print(f"  Moved {direction} - month {i+1}/{abs_months}")
            except Exception as e:
                if months_to_move < 0 and calendar_side == "right":
                    print(f"Left arrow not available or not clickable. Skipping remaining navigation.")
                    return True  # Return True since we can't navigate backward in right calendar
                print(f"Error clicking arrow: {e}")
                return False
        
        print(f"Calendar navigation complete.")
        return True
        
    except Exception as e:
        print(f"Error navigating calendar: {e}")
        import traceback
        traceback.print_exc()
        return False


def select_from_date(driver, wait, target_date):
    """
    Complete function to select the 'from' date in the left calendar.
    Reads current 'from' date, calculates months to move, navigates, and selects the date.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
        target_date: datetime object of the target 'from' date
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        from datetime import datetime
        
        # Read current 'from' date
        current_from_date = get_from_date(driver, wait)
        
        if not current_from_date:
            print("Could not read current 'from' date. Proceeding with target date selection...")
            # Still try to select the date even if we can't read the current one
        else:
            # Calculate months to move
            months_to_move = calculate_months_to_move(current_from_date, target_date)
            
            # Navigate the calendar
            if months_to_move != 0:
                navigate_calendar_months(driver, wait, months_to_move, calendar_side="left")
        
        # Select the target date
        return select_date_in_calendar(driver, wait, target_date, calendar_side="left")
        
    except Exception as e:
        print(f"Error in select_from_date: {e}")
        import traceback
        traceback.print_exc()
        return False


def select_to_date(driver, wait, target_date):
    """
    Complete function to select the 'to' date in the right calendar.
    Reads current 'to' date, calculates months to move, navigates, and selects the date.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
        target_date: datetime object of the target 'to' date
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        from datetime import datetime
        
        # Read current 'to' date
        current_to_date = get_to_date(driver, wait)
        
        if not current_to_date:
            print("Could not read current 'to' date. Proceeding with target date selection...")
            # Still try to select the date even if we can't read the current one
        else:
            # Calculate months to move
            months_to_move = calculate_months_to_move(current_to_date, target_date)
            
            # Navigate the calendar (only if we need to move backward and arrow is available)
            if months_to_move != 0:
                navigate_calendar_months(driver, wait, months_to_move, calendar_side="right")
        
        # Select the target date
        return select_date_in_calendar(driver, wait, target_date, calendar_side="right")
        
    except Exception as e:
        print(f"Error in select_to_date: {e}")
        import traceback
        traceback.print_exc()
        return False


def click_apply_button(driver, wait):
    """
    Click on the Apply button after selecting date range.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print("Looking for Apply button...")
        
        # Wait for the Apply button to be present
        apply_button = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.applyBtn.btn.btn-small.btn-success"))
        )
        
        # Wait for it to be clickable
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.applyBtn.btn.btn-small.btn-success")))
        
        print("Clicking Apply button...")
        apply_button.click()
        time.sleep(2)  # Wait for the action to complete
        
        print("Apply button clicked successfully.")
        return True
        
    except Exception as e:
        print(f"Error clicking Apply button: {e}")
        import traceback
        traceback.print_exc()
        return False


def normalize_plaza_name(plaza_name):
    """
    Normalize plaza name by removing common words like "Toll" and "Plaza".
    Used for matching plaza names in dropdown.
    
    Args:
        plaza_name: The plaza name to normalize
    
    Returns:
        str: Normalized plaza name (lowercase, without "toll" and "plaza")
    """
    if not plaza_name:
        return ""
    
    # Convert to lowercase and split into words
    words = str(plaza_name).lower().strip().split()
    
    # Remove common words to ignore
    ignore_words = ['toll', 'plaza']
    
    # Filter out ignore words
    normalized_words = [word for word in words if word not in ignore_words]
    
    # Join back and return
    return ' '.join(normalized_words).strip()


def select_plaza_from_dropdown(driver, wait, plaza_name="Raibha"):
    """
    Click on the plaza dropdown and select the specified plaza.
    Ignores words "Toll" and "Plaza" when matching.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
        plaza_name: The plaza name to select (default: "Raibha")
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print(f"Looking for plaza dropdown...")
        
        # Normalize the search plaza name (remove "Toll" and "Plaza")
        normalized_search_name = normalize_plaza_name(plaza_name)
        print(f"Searching for plaza: '{plaza_name}' (normalized: '{normalized_search_name}')")
        
        # Wait for the dropdown to be present - using the chzn-container ID
        dropdown = wait.until(
            EC.presence_of_element_located((By.ID, "npciPlazaId_chzn"))
        )
        
        # Find the clickable element (the chzn-single link)
        dropdown_clickable = dropdown.find_element(By.CSS_SELECTOR, "a.chzn-single.chzn-default")
        
        # Wait for it to be clickable
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#npciPlazaId_chzn a.chzn-single.chzn-default")))
        
        print("Clicking plaza dropdown...")
        dropdown_clickable.click()
        time.sleep(1)  # Wait for dropdown to open
        
        # Wait for dropdown options to appear
        print(f"Waiting for dropdown options to appear...")
        options_wait = WebDriverWait(driver, 5)
        
        # Try multiple ways to find the plaza option
        plaza_option = None
        
        try:
            # Get all dropdown options
            options = options_wait.until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "ul.chzn-results li.active-result"))
            )
            
            # Search through options for matching plaza name (ignoring "Toll" and "Plaza")
            for option in options:
                option_text = option.text.strip()
                normalized_option_text = normalize_plaza_name(option_text)
                
                # Check if normalized plaza_name is in the normalized option text
                if normalized_search_name and normalized_search_name in normalized_option_text:
                    plaza_option = option
                    print(f"Found plaza option: '{option_text}' (normalized: '{normalized_option_text}')")
                    break
            
            if not plaza_option:
                # Try exact match on normalized names
                for option in options:
                    option_text = option.text.strip()
                    normalized_option_text = normalize_plaza_name(option_text)
                    if normalized_search_name == normalized_option_text:
                        plaza_option = option
                        print(f"Found plaza option (exact normalized match): '{option_text}'")
                        break
            
            if not plaza_option:
                # List all available options for debugging
                print(f"Available plaza options:")
                for i, option in enumerate(options, 1):
                    option_text = option.text.strip()
                    normalized = normalize_plaza_name(option_text)
                    print(f"  {i}. '{option_text}' (normalized: '{normalized}')")
                raise Exception(f"Could not find plaza option matching '{plaza_name}' (normalized: '{normalized_search_name}')")
            
        except Exception as e:
            print(f"Error finding plaza option: {e}")
            return False
        
        print(f"Selecting '{plaza_option.text.strip()}' from dropdown...")
        # Wait for option to be clickable
        options_wait.until(EC.element_to_be_clickable(plaza_option))
        plaza_option.click()
        time.sleep(1)  # Wait for selection to complete
        
        print(f"Plaza '{plaza_option.text.strip()}' selected successfully.")
        return True
        
    except Exception as e:
        print(f"Error selecting plaza from dropdown: {e}")
        import traceback
        traceback.print_exc()
        return False


def click_export_excel_button(driver, wait):
    """
    Click on the "Export to Excel" button.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print("Looking for Export to Excel button...")
        
        # Wait for the Export button to be present
        export_button = wait.until(
            EC.presence_of_element_located((By.ID, "getExclButton"))
        )
        
        # Wait for it to be clickable
        wait.until(EC.element_to_be_clickable((By.ID, "getExclButton")))
        
        print("Clicking Export to Excel button...")
        export_button.click()
        time.sleep(2)  # Wait for popup to appear
        
        print("Export to Excel button clicked successfully.")
        return True
        
    except Exception as e:
        print(f"Error clicking Export to Excel button: {e}")
        import traceback
        traceback.print_exc()
        return False


def handle_download_popup(driver, wait):
    """
    Handle the popup that appears after clicking Export button and click 'Yes'.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print("Waiting for download popup to appear...")
        
        # Wait for the popup to appear
        popup_wait = WebDriverWait(driver, 10)
        popup = popup_wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".jconfirm-box"))
        )
        
        # Check if popup is visible
        if popup.is_displayed():
            print("Download popup detected. Looking for 'Yes' button...")
            
            # Find the 'Yes' button within the popup
            yes_button = popup.find_element(By.CSS_SELECTOR, "button.btn.btn-purple")
            
            if yes_button:
                print("Clicking 'Yes' button...")
                popup_wait.until(EC.element_to_be_clickable(yes_button))
                yes_button.click()
                time.sleep(3)  # Wait for navigation/action to complete
                print("Popup handled successfully. Navigated to Download Report Dashboard.")
                return True
            else:
                print("Warning: Popup found but 'Yes' button not found.")
                return False
        else:
            print("Warning: Popup element found but not visible.")
            return False
            
    except Exception as e:
        print(f"Error handling download popup: {e}")
        import traceback
        traceback.print_exc()
        return False


def click_top_download_button(driver, wait):
    """
    Wait for tbody to appear and click the download button of the top-most row.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print("Waiting for download table to appear...")
        
        # Wait for tbody to be present
        tbody = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tbody[role='alert']"))
        )
        
        # Wait a bit for rows to load
        time.sleep(2)
        
        # Find all download buttons in the tbody
        download_buttons = tbody.find_elements(By.CSS_SELECTOR, "a.btn.btn-primary.btn-mini[onclick*='downloadCsv']")
        
        if not download_buttons:
            print("No download buttons found in the table.")
            return False
        
        # Click the first (top-most) download button
        top_button = download_buttons[0]
        print(f"Clicking top download button (found {len(download_buttons)} total buttons)...")
        
        # Get the onclick attribute to see what will be downloaded
        onclick_attr = top_button.get_attribute("onclick")
        print(f"Download button onclick: {onclick_attr}")
        
        wait.until(EC.element_to_be_clickable(top_button))
        top_button.click()
        
        print("Download button clicked. File download initiated...")
        time.sleep(3)  # Wait for download to start
        
        return True
        
    except Exception as e:
        print(f"Error clicking download button: {e}")
        import traceback
        traceback.print_exc()
        return False


def detect_downloaded_zip_file(downloads_folder, max_age_minutes=10):
    try:
        downloads_path = Path(downloads_folder)
        if not downloads_path.exists():
            print(f"Downloads folder not found: {downloads_folder}")
            return None
        
        # Debug: Print current time
        current_time = time.time()
        current_datetime = datetime.fromtimestamp(current_time)
        print("\n" + "="*60)
        print("DEBUG INFORMATION")
        print("="*60)
        print(f"Current time (Python): {current_datetime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Current time (timestamp): {current_time}")
        print("="*60)
        
        # Debug: Print some files in Downloads folder
        print(f"\nScanning Downloads folder: {downloads_folder}")
        all_files = list(downloads_path.iterdir())
        print(f"Total files/folders found: {len(all_files)}")
        
        # Show first 10 files (any type) for debugging
        print("\nSample files in Downloads folder (first 10):")
        for i, file_path in enumerate(sorted(all_files, key=lambda x: x.stat().st_mtime, reverse=True)[:10], 1):
            try:
                file_mtime = file_path.stat().st_mtime
                file_datetime = datetime.fromtimestamp(file_mtime)
                age_seconds = current_time - file_mtime
                file_type = "DIR" if file_path.is_dir() else file_path.suffix.upper() or "FILE"
                size_mb = file_path.stat().st_size / 1024 / 1024 if file_path.is_file() else 0
                
                print(f"  {i}. {file_path.name}")
                print(f"     Type: {file_type}")
                print(f"     Modified: {file_datetime.strftime('%Y-%m-%d %H:%M:%S')} ({age_seconds/60:.2f} minutes ago)")
                if file_path.is_file():
                    print(f"     Size: {size_mb:.2f} MB")
            except Exception as e:
                print(f"  {i}. {file_path.name} (Error reading: {e})")
        
        print("\n" + "="*60)
        print(f"Searching for ZIP files modified within last {max_age_minutes} minutes...")
        print("="*60)
        
        max_age_seconds = max_age_minutes * 60
        
        # Find all ZIP files modified within the last 10 minutes
        recent_files = []
        
        for zip_file in downloads_path.glob("*.zip"):
            try:
                # Get file modification time
                file_mtime = zip_file.stat().st_mtime
                
                # Calculate age in seconds
                age_seconds = current_time - file_mtime
                
                # Check if file was modified within the last 10 minutes
                if age_seconds <= max_age_seconds:
                    recent_files.append((zip_file, file_mtime, age_seconds))
                    print(f"  Found: {zip_file.name}")
                    print(f"    Modified: {age_seconds:.1f} seconds ago ({age_seconds/60:.2f} minutes)")
                    print(f"    Size: {zip_file.stat().st_size / 1024 / 1024:.2f} MB")
            except (OSError, PermissionError) as e:
                # File might be locked, skip it
                print(f"  Skipping {zip_file.name} (locked or inaccessible): {e}")
                continue
        
        if not recent_files:
            print(f"No ZIP files found modified within the last {max_age_minutes} minutes.")
            return None
        
        # Sort by modification time (most recent first)
        recent_files.sort(key=lambda x: x[1], reverse=True)
        
        # Return the most recently modified file
        most_recent_file, mtime, age = recent_files[0]
        print(f"\nSelected most recent ZIP file: {most_recent_file.name}")
        print(f"  Modified: {age:.1f} seconds ago ({age/60:.2f} minutes)")
        print(f"  Size: {most_recent_file.stat().st_size / 1024 / 1024:.2f} MB")
        
        return str(most_recent_file)
        
    except Exception as e:
        print(f"Error detecting ZIP file: {e}")
        import traceback
        traceback.print_exc()
        return None


def extract_zip_file(zip_file_path, extract_to_folder=None):
    """
    Extract a ZIP file to a folder.
    
    Args:
        zip_file_path: Path to the ZIP file
        extract_to_folder: Folder to extract to (default: same folder as ZIP file)
    
    Returns:
        str: Path to the extracted folder, or None if failed
    """
    try:
        import zipfile
        
        zip_path = Path(zip_file_path)
        if not zip_path.exists():
            print(f"ZIP file not found: {zip_file_path}")
            return None
        
        if extract_to_folder is None:
            # Extract to a folder with the same name as the ZIP file (without extension)
            extract_to_folder = zip_path.parent / zip_path.stem
        else:
            extract_to_folder = Path(extract_to_folder)
        
        extract_to_folder.mkdir(parents=True, exist_ok=True)
        
        print(f"Extracting ZIP file: {zip_path.name}")
        print(f"Extracting to: {extract_to_folder}")
        
        # Extract using Python's built-in zipfile module
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to_folder)
        
        print("ZIP file extracted successfully.")
        return str(extract_to_folder)
        
    except Exception as e:
        print(f"Error extracting ZIP file: {e}")
        import traceback
        traceback.print_exc()
        return None


def extract_rar_file(rar_file_path, extract_to_folder=None):
    """
    Extract a RAR file to a folder.
    
    Args:
        rar_file_path: Path to the RAR file
        extract_to_folder: Folder to extract to (default: same folder as RAR file)
    
    Returns:
        str: Path to the extracted folder, or None if failed
    """
    try:
        rar_path = Path(rar_file_path)
        if not rar_path.exists():
            print(f"RAR file not found: {rar_file_path}")
            return None
        
        if extract_to_folder is None:
            # Extract to a folder with the same name as the RAR file (without extension)
            extract_to_folder = rar_path.parent / rar_path.stem
        else:
            extract_to_folder = Path(extract_to_folder)
        
        extract_to_folder.mkdir(parents=True, exist_ok=True)
        
        print(f"Extracting RAR file: {rar_path.name}")
        print(f"Extracting to: {extract_to_folder}")
        
        # Try using unrar command (Windows)
        try:
            # Check if unrar is available
            result = subprocess.run(["unrar", "x", str(rar_path), str(extract_to_folder)], 
                                  capture_output=True, text=True, timeout=60)
            if result.returncode == 0:
                print("RAR file extracted successfully using unrar command.")
                return str(extract_to_folder)
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
        
        # Try using Python rarfile library if available
        try:
            import rarfile
            with rarfile.RarFile(rar_file_path) as rf:
                rf.extractall(extract_to_folder)
            print("RAR file extracted successfully using rarfile library.")
            return str(extract_to_folder)
        except ImportError:
            print("rarfile library not available. Trying alternative method...")
        except Exception as e:
            print(f"Error extracting with rarfile: {e}")
        
        # Fallback: Try 7zip if available
        try:
            # Common 7zip paths on Windows
            seven_zip_paths = [
                r"C:\Program Files\7-Zip\7z.exe",
                r"C:\Program Files (x86)\7-Zip\7z.exe",
            ]
            
            for seven_zip in seven_zip_paths:
                if os.path.exists(seven_zip):
                    result = subprocess.run([seven_zip, "x", str(rar_path), f"-o{extract_to_folder}", "-y"],
                                          capture_output=True, timeout=60)
                    if result.returncode == 0:
                        print("RAR file extracted successfully using 7-Zip.")
                        return str(extract_to_folder)
                    break
        except Exception as e:
            print(f"Error extracting with 7-Zip: {e}")
        
        print("Could not extract RAR file. Please install unrar, rarfile library, or 7-Zip.")
        return None
        
    except Exception as e:
        print(f"Error extracting RAR file: {e}")
        import traceback
        traceback.print_exc()
        return None


def cleanup_files(zip_file_path=None, extracted_folder_path=None, output_file_path=None):
    """
    Clean up temporary files and folders created during processing.
    
    Args:
        zip_file_path: Path to the downloaded ZIP file to delete
        extracted_folder_path: Path to the extracted folder to delete
        output_file_path: Path to the combined output Excel file to delete
    """
    print("\n" + "="*60)
    print("Cleaning up temporary files...")
    print("="*60)
    
    deleted_count = 0
    
    # Delete ZIP file
    if zip_file_path:
        try:
            zip_path = Path(zip_file_path)
            if zip_path.exists():
                zip_path.unlink()
                print(f"✓ Deleted ZIP file: {zip_path.name}")
                deleted_count += 1
            else:
                print(f"⊘ ZIP file not found: {zip_path.name}")
        except Exception as e:
            print(f"✗ Error deleting ZIP file: {e}")
    
    # Delete extracted folder
    if extracted_folder_path:
        try:
            extracted_path = Path(extracted_folder_path)
            if extracted_path.exists() and extracted_path.is_dir():
                shutil.rmtree(extracted_path)
                print(f"✓ Deleted extracted folder: {extracted_path.name}")
                deleted_count += 1
            else:
                print(f"⊘ Extracted folder not found: {extracted_folder_path}")
        except Exception as e:
            print(f"✗ Error deleting extracted folder: {e}")
    
    # Delete combined output file (if in current directory)
    if output_file_path:
        try:
            output_path = Path(output_file_path)
            if output_path.exists():
                output_path.unlink()
                print(f"✓ Deleted output file: {output_path.name}")
                deleted_count += 1
            else:
                # Also check in script directory
                script_dir = Path(__file__).parent
                output_path_alt = script_dir / output_file_path
                if output_path_alt.exists():
                    output_path_alt.unlink()
                    print(f"✓ Deleted output file: {output_path_alt.name}")
                    deleted_count += 1
        except Exception as e:
            print(f"✗ Error deleting output file: {e}")
    
    # Also check for combined_output.xlsx in script directory
    try:
        script_dir = Path(__file__).parent
        combined_output = script_dir / "combined_output.xlsx"
        if combined_output.exists():
            combined_output.unlink()
            print(f"✓ Deleted combined output file: {combined_output.name}")
            deleted_count += 1
    except Exception as e:
        print(f"✗ Error deleting combined output file: {e}")
    
    print(f"\nCleanup completed. Deleted {deleted_count} file(s)/folder(s).")
    print("="*60)


def run_update_db_script(extracted_folder, plaza_name, from_date, to_date):
    """
    Process the extracted Excel files and update the database.
    This function replicates the logic from update_db.py.
    
    Args:
        extracted_folder: Path to the extracted folder containing Excel files
        plaza_name: Name of the plaza (for entity_name in database)
        from_date: Start date (datetime object or string in 'YYYY-MM-DD' format)
        to_date: End date (datetime object or string in 'YYYY-MM-DD' format)
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Import the functions from extraction module
        import sys
        script_dir = Path(__file__).parent
        if str(script_dir) not in sys.path:
            sys.path.insert(0, str(script_dir))
        
        from extraction import process_folder_excel_files, read_processed_excel_and_update_db, delete_violation_amt
        
        output_file = "combined_output.xlsx"
        plaza_name_lower = plaza_name.lower()
        
        print(f"Processing Excel files from folder: {extracted_folder}")
        print(f"Plaza name: {plaza_name_lower}")
        
        print("\n" + "=" * 60)
        print("Step 1: Processing all Excel files in folder")
        print("=" * 60)
        
        # Step 1: Process all Excel files in the folder and generate combined output
        combined_totals = process_folder_excel_files(
            folder_path=extracted_folder,
            output_file_path=output_file
        )
        
        if combined_totals is None:
            print("\nError: Failed to process Excel files. Exiting...")
            return False


        
        print("\n" + "=" * 60)
        print(f"Deleting violation amt from database for plaza: {plaza_name_lower} for date range: {from_date} to {to_date}")
        print("=" * 60)
        delete_violation_amt(plaza_name_lower, from_date, to_date)

        print("Step 2: Updating database with combined results")
        print("=" * 60)
        
        # Step 2: Read the combined output file and update database
        stats = read_processed_excel_and_update_db(
            input_file_path=output_file,
            plaza_name=plaza_name_lower
        )
        
        # Print final summary
        print("\n" + "=" * 60)
        print("Final Summary:")
        print("=" * 60)
        if stats['updated'] > 0:
            print(f"  ✓ Updated: {stats['updated']}")
            print(f"  ⊘ Skipped (already exists): {stats.get('already_exists', 0)}")
            print(f"  ⊘ Skipped (no valid rows): {stats.get('no_valid_rows', 0)}")
            print(f"  ⊘ Skipped (other): {stats.get('skipped', 0)}")
        else:
            print("  No updates were made.")
            print(f"  Skipped (already exists): {stats.get('already_exists', 0)}")
            print(f"  Skipped (no valid rows): {stats.get('no_valid_rows', 0)}")
            print(f"  Skipped (other): {stats.get('skipped', 0)}")
        print("=" * 60)
        
        return True
        
    except Exception as e:
        print(f"Error running update_db script: {e}")
        import traceback
        traceback.print_exc()
        return False


def attempt_login(driver, wait, username, password):
    """
    Attempt a single login with username, password, and captcha.
    
    Args:
        driver: WebDriver instance
        wait: WebDriverWait instance
        username: Username for login
        password: Password for login
    
    Returns:
        bool: True if login successful, False otherwise
    """
    try:
        # Wait for username field to be present
        username_field = wait.until(EC.presence_of_element_located((By.ID, "j_username")))
        
        # Enter username
        print(f"Entering username: {username}")
        username_field.clear()
        username_field.send_keys(username)
        time.sleep(0.5)
        
        # Enter password
        print("Entering password...")
        password_field = driver.find_element(By.ID, "j_password")
        password_field.clear()
        password_field.send_keys(password)
        time.sleep(0.5)
        
        # Get captcha image
        print("Extracting captcha image...")
        captcha_element = driver.find_element(By.ID, "captchaId")
        captcha_image_array = get_captcha_image(driver, captcha_element)
        
        # Extract text from captcha using OCR
        print("Processing captcha with OCR...")
        extracted_text_list = extract_text(captcha_image_array)
        
        # Join the extracted text and clean it
        raw_captcha_text = " ".join(extracted_text_list)
        captcha_text = special_character_remover(raw_captcha_text)
        
        print(f"Extracted captcha text: {captcha_text}")
        
        if not captcha_text:
            print("Warning: Could not extract captcha text.")
            return False
        
        # Enter captcha text
        print("Entering captcha text...")
        captcha_field = driver.find_element(By.ID, "jcaptcha")
        captcha_field.clear()
        captcha_field.send_keys(captcha_text)
        time.sleep(0.5)
        
        # Click Sign In button
        print("Clicking Sign In button...")
        sign_in_button = driver.find_element(By.ID, "login")
        sign_in_button.click()
        
        # Wait a bit to see if login was successful
        time.sleep(3)

        page_text = get_idfc_visible_page_text(driver)
        if is_idfc_session_limit_message(page_text):
            print("Sign-in failed: maximum sessions limit detected on portal.")
            for line in page_text.splitlines():
                line = line.strip()
                if line and is_idfc_session_limit_message(line):
                    print(f"Error message: {line}")
            prepare_idfc_login_page(driver)
            return False

        # Check if we're still on login page (login failed) or redirected (login successful)
        current_url = driver.current_url
        if "login" not in current_url.lower():
            print("Login successful! Redirected to:", current_url)
            return True
        else:
            print("Login failed. Still on login page.")
            for line in page_text.splitlines():
                line = line.strip()
                if line:
                    print(f"Error message: {line}")
            if is_idfc_session_limit_message(page_text):
                prepare_idfc_login_page(driver)
            return False
            
    except Exception as e:
        print(f"Error during login attempt: {e}")
        return False


def login_to_idfc(username, password, from_date=None, to_date=None, plaza_name="Raibha", 
                  downloads_folder=r"C:\Users\Divyesh Parmar\Downloads", 
                  process_download=True, headless=False, max_attempts=1, max_restarts=15,
                  plaza_name_db: str | None = None):
    for restart_count in range(1, max_restarts + 1):
        driver = None
        try:
            if restart_count > 1:
                print("\n" + "="*80)
                print(f"RESTARTING BROWSER - Attempt {restart_count} of {max_restarts}")
                print("="*80)
                print("Previous login attempts failed. Restarting browser and retrying...")
                wait_for_idfc_session_slot(reason="browser restart")

            probe_release_stale_idfc_session(headless=headless)

            print("Setting up browser (fresh profile, no shared cookies/cache)...")
            driver = setup_driver(headless=headless, fresh_profile=True)

            print(f"Preparing login page: {IDFC_LOGIN_URL}")
            prepare_idfc_login_page(driver)

            wait = WebDriverWait(driver, 10)

            login_success = False
            for attempt in range(1, max_attempts + 1):
                print(f"\n{'='*50}")
                print(f"Login Attempt {attempt} of {max_attempts}")
                print(f"{'='*50}")

                success = attempt_login(driver, wait, username, password)
                
                if success:
                    login_success = True
                    break
                else:
                    if is_idfc_session_limit_message(get_idfc_visible_page_text(driver)):
                        prepare_idfc_login_page(driver)
                        wait_for_idfc_session_slot(reason="session limit on login")
                    if attempt < max_attempts:
                        print(f"\n✗ Login failed on attempt {attempt}. Retrying with new captcha...")
                        time.sleep(2)
                        try:
                            wait.until(EC.presence_of_element_located((By.ID, "j_username")))
                        except Exception:
                            print("Refreshing page for new captcha...")
                            prepare_idfc_login_page(driver)
                            wait.until(EC.presence_of_element_located((By.ID, "j_username")))
                    else:
                        print(f"\n✗ Login failed after {max_attempts} attempts in this browser session.")

            # If login succeeded, proceed with the rest of the process
            if login_success:
                print(f"\n✓ Login successful on attempt {attempt}!")
                
                # Handle popup if it appears
                handle_popup(driver, wait)
                
                # Click on Concessionaire button
                click_concessionaire_button(driver, wait)
                
                # Click on Life Cycle Report(New) link
                click_life_cycle_report(driver, wait)
                
                # Select date type from dropdown - CRITICAL: Must succeed or restart
                print("\n" + "="*60)
                print("CRITICAL: Selecting 'Reader Read Date' from dropdown...")
                print("="*60)
                date_type_selected = select_date_type(driver, wait, date_type="Reader Read Date")
                
                if not date_type_selected:
                    print("\n" + "="*60)
                    print("CRITICAL ERROR: Could not select 'Reader Read Date' from dropdown!")
                    print("="*60)
                    print("This is a critical error. Closing browser and will restart the browser")
                    print("to retry the same plaza (up to max_restarts times).")
                    print("="*60)
                    
                    # Close browser
                    if driver:
                        try:
                            shutdown_idfc_driver(driver)
                            driver = None
                            print("\nBrowser closed.")
                        except:
                            pass
                    
                    # Break out of processing to trigger restart
                    login_success = False
                    break  # Break out of the login success processing block
                
                print("✓ 'Reader Read Date' selected successfully. Proceeding...")
                print("="*60)
                
                # Open date range picker
                open_date_range_picker(driver, wait)
                
                # Select 'from' date if provided
                if from_date:
                    select_from_date(driver, wait, from_date)
                
                # Select 'to' date if provided
                if to_date:
                    select_to_date(driver, wait, to_date)
                
                # Click Apply button to apply the date range
                click_apply_button(driver, wait)
                
                # Select plaza from dropdown
                select_plaza_from_dropdown(driver, wait, plaza_name=plaza_name)
                
                # Click Export to Excel button
                if click_export_excel_button(driver, wait):
                    # Handle the popup and click Yes
                    if handle_download_popup(driver, wait):
                        # Wait 1.5 minutes for file processing to complete
                        print("\nWaiting 1.5 minutes for file processing to complete...")
                        wait_time = 90  # 1.5 minutes = 90 seconds
                        for remaining in range(wait_time, 0, -10):
                            print(f"  Waiting... {remaining} seconds remaining", end='\r')
                            time.sleep(10)
                        print(f"  Waiting complete. Proceeding to download...")
                        
                        # Wait for download table and click top download button
                        if click_top_download_button(driver, wait):
                            if process_download:
                                # Wait for ZIP file to be downloaded (browser stays open during download)
                                print("\nWaiting for ZIP file download to complete...")
                                print("(Browser will remain open during download)")
                                
                                # Detect the downloaded ZIP file (wait up to 2 minutes for download)
                                zip_file = None
                                max_wait_time = 300  # 2 minutes
                                start_wait = time.time()
                                
                                while time.time() - start_wait < max_wait_time:
                                    zip_file = detect_downloaded_zip_file(downloads_folder, max_age_minutes=10)
                                    if zip_file:
                                        print(f"✓ ZIP file download detected: {Path(zip_file).name}")
                                        break
                                    time.sleep(5)  # Check every 5 seconds
                                    remaining = int(max_wait_time - (time.time() - start_wait))
                                    print(f"  Waiting for download... {remaining} seconds remaining", end='\r')
                                
                                # Close browser after download is detected (download is complete)
                                print("\nDownload complete. Closing browser...")
                                if driver:
                                    try:
                                        shutdown_idfc_driver(driver)
                                        driver = None
                                        print("Browser closed successfully.")
                                    except Exception as e:
                                        print(f"Error closing browser: {e}")
                                
                                # Now process the downloaded file (browser is closed)
                                processing_success = False
                                if zip_file:
                                    extracted_folder = None
                                    
                                    try:
                                        print("\n" + "="*60)
                                        print("Processing downloaded ZIP file...")
                                        print("="*60)
                                        
                                        # Extract the ZIP file
                                        extracted_folder = extract_zip_file(zip_file)
                                        
                                        if extracted_folder:
                                            print(f"Extracted to: {extracted_folder}")
                                            
                                            # Run update_db.py script
                                            # Use DB plaza name if provided, otherwise fall back to UI plaza name
                                            db_plaza = plaza_name_db if plaza_name_db else plaza_name
                                            print(f"Running update_db.py for DB plaza name: {db_plaza}")
                                            processing_success = run_update_db_script(extracted_folder, db_plaza, from_date, to_date)
                                        else:
                                            print("Failed to extract ZIP file.")
                                    finally:
                                        # Always clean up files, even if processing failed
                                        cleanup_files(
                                            zip_file_path=zip_file,
                                            extracted_folder_path=extracted_folder,
                                            output_file_path="combined_output.xlsx"
                                        )
                                else:
                                    print("Could not detect downloaded ZIP file within timeout.")
                                
                                # Return success status (True if processing succeeded, False otherwise)
                                return processing_success
                        else:
                            print("Failed to click download button.")
                    else:
                        print("Failed to handle download popup.")
                else:
                    print("Failed to click Export to Excel button.")
                
                # If we reach here, processing didn't complete successfully
                # Close browser if still open
                if driver:
                    try:
                        shutdown_idfc_driver(driver)
                    except:
                        pass
                    driver = None
                return True  # Return True since login succeeded, even if processing failed later
            
            # If login failed after all attempts OR critical error occurred, close browser and restart
            if not login_success:
                # Check if it was a critical error (login succeeded but critical step failed)
                # or if login itself failed
                if driver is None:
                    # Browser was already closed due to critical error
                    print(f"\n✗ Critical error occurred during processing.")
                else:
                    print(f"\n✗ All {max_attempts} login attempts failed in this browser session.")
                
                print("Closing browser and will restart the entire process for the same plaza...")
                
                # Close browser if still open
                if driver:
                    try:
                        shutdown_idfc_driver(driver)
                        driver = None
                    except:
                        pass
                
                # If we haven't reached max restarts, continue to next iteration (restart browser)
                if restart_count < max_restarts:
                    print(f"Will restart browser and retry entire process (restart {restart_count + 1}/{max_restarts})...")
                    time.sleep(2)  # Brief pause before restarting
                    continue
                else:
                    # All restarts exhausted
                    print(f"\n✗ All {max_restarts} browser restart attempts exhausted.")
                    print("Could not complete process after multiple browser restarts.")
                    return False
        
        except Exception as e:
            print(f"Error during login process: {e}")
            import traceback
            traceback.print_exc()
            if driver:
                try:
                    shutdown_idfc_driver(driver)
                    driver = None
                except:
                    pass
            
            # If we haven't reached max restarts, try again
            if restart_count < max_restarts:
                print(f"Error occurred. Will restart browser and retry (restart {restart_count + 1}/{max_restarts})...")
                continue
            else:
                return False
    
    # If we exit the loop, all restarts failed
    return False


def process_downloaded_file(downloads_folder, plaza_name):
    """
    Process a downloaded ZIP file without login.
    
    Args:
        downloads_folder: Path to the Downloads folder
        plaza_name: Name of the plaza (for entity_name in database)
    """
    print("\n" + "="*60)
    print("Detecting downloaded ZIP file...")
    print("="*60)
    
    # Detect the downloaded ZIP file (modified within last 10 minutes)
    zip_file = detect_downloaded_zip_file(downloads_folder, max_age_minutes=10)
    
    if zip_file:
        print(f"\n✓ Found ZIP file: {zip_file}")
        
        # # Extract the ZIP file
        # extracted_folder = extract_rar_file(zip_file)
        # 
        # if extracted_folder:
        #     print(f"Extracted to: {extracted_folder}")
        #     
        #     # Run update_db.py script
        #     run_update_db_script(extracted_folder, plaza_name)
        # else:
        #     print("Failed to extract ZIP file.")
    else:
        print("Could not detect downloaded ZIP file.")


def parse_date_mmddyyyy(date_str):
    """
    Parse date string in dd-mm-year format to datetime object.
    Tries both dd-mm-year and mm-dd-year formats.
    Also handles datetime objects and pandas Timestamps.
    
    Args:
        date_str: Date string in format "dd-mm-year" (e.g., "12-01-2026" = January 12, 2026)
                  or "mm-dd-year" (e.g., "01-12-2026" = January 12, 2026)
                  or datetime object/Timestamp from Excel
    
    Returns:
        datetime object or None if parsing fails
    """
    try:
        if pd.isna(date_str) or str(date_str).strip() == '':
            return None
        
        # Check if it's already a datetime object or pandas Timestamp
        if isinstance(date_str, (datetime, pd.Timestamp)):
            return date_str.to_pydatetime() if isinstance(date_str, pd.Timestamp) else date_str
        
        # Check if it's a datetime-like object
        if hasattr(date_str, 'year') and hasattr(date_str, 'month') and hasattr(date_str, 'day'):
            try:
                return datetime(date_str.year, date_str.month, date_str.day)
            except:
                pass
        
        date_str_clean = str(date_str).strip()
        
        # Try parsing datetime string format first (e.g., "2025-12-10 00:00:00")
        try:
            # Try common datetime formats
            for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
                try:
                    return datetime.strptime(date_str_clean, fmt)
                except ValueError:
                    continue
        except:
            pass
        
        # First try dd-mm-year format (day-month-year)
        # This is the primary format based on user's example: '12-01-2026' = January 12, 2026
        try:
            return datetime.strptime(date_str_clean, "%d-%m-%Y")
        except ValueError:
            pass
        
        # Try dd/mm/yyyy format
        try:
            return datetime.strptime(date_str_clean, "%d/%m/%Y")
        except ValueError:
            pass
        
        # Fallback: Try mm-dd-year format (month-day-year)
        try:
            return datetime.strptime(date_str_clean, "%m-%d-%Y")
        except ValueError:
            pass
        
        # Try mm/dd/yyyy format
        try:
            return datetime.strptime(date_str_clean, "%m/%d/%Y")
        except ValueError:
            pass
        
        print(f"Warning: Could not parse date: {date_str}")
        return None
        
    except Exception as e:
        print(f"Error parsing date '{date_str}': {e}")
        return None


def process_plazas_from_excel(excel_file_path, username, password, downloads_folder):
    """
    Read Excel file and process all plazas that are not marked as 'done'.
    
    Args:
        excel_file_path: Path to the Excel file with plaza information
        username: Username for IDFC login
        password: Password for IDFC login
        downloads_folder: Path to Downloads folder
    
    Returns:
        bool: True if all processing completed, False otherwise
    """
    try:
        # Read Excel file
        print("\n" + "="*60)
        print("Reading Excel file...")
        print("="*60)
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        # Validate required columns
        required_columns = ['Plaza_name_IDFC', 'Plaza_name_DB', 'Done', 'Date']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Error: Missing required columns: {missing_columns}")
            return False
        
        print(f"Found {len(df)} rows in Excel file")
        
        # Extract dates from Date column (first two non-empty values)
        # These dates will be used for all plazas
        dates = []
        for idx, row in df.iterrows():
            date_value = row['Date']
            if pd.notna(date_value):
                date_str = str(date_value).strip()
                print(f"  Reading date from row {idx + 2}: '{date_str}'")
                parsed_date = parse_date_mmddyyyy(date_str)
                if parsed_date:
                    dates.append(parsed_date)
                    print(f"    Parsed as: {parsed_date.strftime('%Y-%m-%d')} ({parsed_date.strftime('%B %d, %Y')})")
                    if len(dates) >= 2:
                        break
                else:
                    print(f"    Warning: Could not parse this date")
        
        if len(dates) < 2:
            print(f"Error: Could not find 2 dates in Date column. Found {len(dates)} date(s).")
            print("Please ensure the Date column has two dates in dd-mm-year format (e.g., '12-01-2026' = January 12, 2026).")
            return False
        
        from_date = dates[0]
        to_date = dates[1]
        
        print(f"\nDate range for all plazas:")
        print(f"  From: {from_date.strftime('%Y-%m-%d')} ({from_date.strftime('%B %d, %Y')})")
        print(f"  To: {to_date.strftime('%Y-%m-%d')} ({to_date.strftime('%B %d, %Y')})")
        
        # Filter rows that are not done and have plaza names (skip rows that only have dates)
        # A row is a plaza row if it has Plaza_name_IDFC and Plaza_name_DB values
        all_plaza_rows = df[
            (df['Plaza_name_IDFC'].notna()) &
            (df['Plaza_name_DB'].notna())
        ]
        
        # Separate done and pending plazas
        done_rows = all_plaza_rows[
            (all_plaza_rows['Done'].notna()) &
            (all_plaza_rows['Done'].astype(str).str.strip().str.lower() == 'done')
        ]
        
        pending_rows = all_plaza_rows[
            (all_plaza_rows['Done'].isna()) |
            (all_plaza_rows['Done'].astype(str).str.strip().str.lower() != 'done')
        ]
        
        print(f"\n" + "="*60)
        print("PLAZA STATUS SUMMARY")
        print("="*60)
        print(f"Total plaza rows found: {len(all_plaza_rows)}")
        print(f"  ✓ Already processed (marked 'done'): {len(done_rows)}")
        print(f"  ⊘ Remaining to process: {len(pending_rows)}")
        print("="*60)
        
        # Show list of already done plazas
        if len(done_rows) > 0:
            print(f"\nAlready processed plazas (will be skipped):")
            for idx, row in done_rows.iterrows():
                plaza_idfc = str(row['Plaza_name_IDFC']).strip() if pd.notna(row['Plaza_name_IDFC']) else "N/A"
                print(f"  - {plaza_idfc}")
        
        # Show list of plazas to process
        if len(pending_rows) > 0:
            print(f"\nPlazas to process:")
            for idx, row in pending_rows.iterrows():
                plaza_idfc = str(row['Plaza_name_IDFC']).strip() if pd.notna(row['Plaza_name_IDFC']) else "N/A"
                plaza_db = str(row['Plaza_name_DB']).strip() if pd.notna(row['Plaza_name_DB']) else "N/A"
                print(f"  - {plaza_idfc} (DB: {plaza_db})")
        
        if len(pending_rows) == 0:
            print("\n✓ All plazas are already marked as 'done'. Nothing to process.")
            return True
        
        print(f"\n" + "="*60)
        print(f"Starting processing of {len(pending_rows)} plaza(s)...")
        print("="*60)
        
        # Process each plaza row
        driver = None
        success_count = 0
        fail_count = 0
        
        for idx, row in pending_rows.iterrows():
            try:
                plaza_name_idfc = str(row['Plaza_name_IDFC']).strip() if pd.notna(row['Plaza_name_IDFC']) else None
                plaza_name_db = str(row['Plaza_name_DB']).strip() if pd.notna(row['Plaza_name_DB']) else None
                
                if not plaza_name_idfc or not plaza_name_db:
                    print(f"\nRow {idx + 2}: Skipping - Missing plaza name (IDFC: {plaza_name_idfc}, DB: {plaza_name_db})")
                    continue
                
                print("\n" + "="*80)
                print(f"Processing Row {idx + 2}/{len(pending_rows)}")
                print("="*80)
                print(f"Plaza (IDFC): {plaza_name_idfc}")
                print(f"Plaza (DB): {plaza_name_db}")
                print(f"Date Range: {from_date.strftime('%m-%d-%Y')} to {to_date.strftime('%m-%d-%Y')}")
                print("="*80)
                
                # Run full process for this plaza
                # Note: login_to_idfc will close browser after download and return success status
                processing_success = login_to_idfc(
                    username=username,
                    password=password,
                    from_date=from_date,
                    to_date=to_date,
                    plaza_name=plaza_name_idfc,   # Use IDFC name for dropdown / UI
                    downloads_folder=downloads_folder,
                    process_download=True,
                    headless=False,
                    plaza_name_db=plaza_name_db   # Use DB name when updating our database
                )
                
                # Check if processing failed due to critical error (e.g., date type selection failed)
                if not processing_success:
                    print("\n" + "="*80)
                    print("PROCESSING FAILED: Could not complete process for this plaza")
                    print("="*80)
                    print("Possible causes:")
                    print("  - Could not select 'Reader Read Date' from dropdown")
                    print("  - Other critical error during processing")
                    print("  - All browser restart attempts exhausted")
                    print("="*80)
                    print("Skipping this plaza and continuing to next plaza...")
                    print("="*80)
                    
                    fail_count += 1
                    # Continue to next plaza instead of stopping
                    continue
                
                # Processing succeeded
                else:
                    # Processing completed successfully
                    # Mark as done in Excel
                    df.at[idx, 'Done'] = 'done'
                    
                    # Save Excel file after each successful processing
                    df.to_excel(excel_file_path, index=False, engine='openpyxl')
                    print(f"\n✓ Row {idx + 2} processed successfully. Marked as 'done' in Excel.")
                    success_count += 1
                    
                    # Small delay before next plaza (browser is already closed)
                    print("Waiting 3 seconds before processing next plaza...")
                    time.sleep(3)
                    
            except Exception as e:
                print(f"\n✗ Error processing row {idx + 2}: {e}")
                import traceback
                traceback.print_exc()
                fail_count += 1
                if driver:
                    try:
                        shutdown_idfc_driver(driver)
                    except:
                        pass
                    driver = None
        
        # Final summary
        print("\n" + "="*80)
        print("PROCESSING SUMMARY")
        print("="*80)
        print(f"Total rows processed: {len(pending_rows)}")
        print(f"  ✓ Successful: {success_count}")
        print(f"  ✗ Failed: {fail_count}")
        print("="*80)
        
        return fail_count == 0
        
    except Exception as e:
        print(f"Error reading/processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main function - process all plazas from Excel file"""
    # Configuration
    USERNAME = "Admin_1"
    PASSWORD = "Nhai@123456"
    downloads_folder = r"C:\Users\Divyesh Parmar\Downloads"
    
    # Excel file path - UPDATE THIS to your Excel file path
    excel_file_path = "Reference_File - updated.xlsx"  # Change this to your actual file path
    
    print("="*80)
    print("IDFC Bot - Batch Processing from Excel")
    print("="*80)
    print(f"Excel file: {excel_file_path}")
    print(f"Downloads folder: {downloads_folder}")
    print("="*80)
    
    # Check if Excel file exists
    if not Path(excel_file_path).exists():
        print(f"\nError: Excel file not found: {excel_file_path}")
        print("Please update 'excel_file_path' in the main() function.")
        return
    
    # Process all plazas from Excel
    success = process_plazas_from_excel(
        excel_file_path=excel_file_path,
        username=USERNAME,
        password=PASSWORD,
        downloads_folder=downloads_folder
    )
    
    if success:
        print("\n" + "="*80)
        print("All plazas processed successfully!")
        print("="*80)
    else:
        print("\n" + "="*80)
        print("Some plazas failed to process. Check the output above.")
        print("="*80)


if __name__ == "__main__":
    main()

