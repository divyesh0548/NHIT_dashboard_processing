"""
IDFC Downloader Bot
Downloads Life Cycle Report files from IDFC portal and saves them to a folder.
Does NOT process or delete the files - just downloads and renames them.
"""

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
from db_config import load_env_file, get_process_date_range_datetimes
from datetime import datetime, timedelta
import shutil
import pandas as pd
import re

# Import all helper functions from IDFC_bot
from IDFC_bot import (
    setup_driver,
    shutdown_idfc_driver,
    probe_release_stale_idfc_session,
    prepare_idfc_login_page,
    wait_for_idfc_session_slot,
    is_idfc_session_limit_message,
    get_idfc_visible_page_text,
    IDFC_LOGIN_URL,
    get_captcha_image,
    get_captcha_from_screenshot,
    handle_popup,
    click_concessionaire_button,
    click_life_cycle_report,
    select_date_type,
    open_date_range_picker,
    get_from_date,
    calculate_months_to_move,
    select_date_in_calendar,
    get_to_date,
    navigate_calendar_months,
    select_from_date,
    select_to_date,
    click_apply_button,
    normalize_plaza_name,
    select_plaza_from_dropdown,
    click_export_excel_button,
    handle_download_popup,
    click_top_download_button,
    detect_downloaded_zip_file,
    attempt_login,
    parse_date_mmddyyyy
)


def split_date_range_to_month_intervals(
    start_dt: datetime,
    end_dt: datetime,
     ) -> list[tuple[datetime, datetime]]:
    """
    Split an inclusive date range into month-bounded intervals for portal limits.

    Example: 15-Nov-2025 to 25-Jan-2026 becomes:
      15-Nov to 30-Nov, 01-Dec to 31-Dec, 01-Jan to 25-Jan
    """
    if end_dt < start_dt:
        raise ValueError("End date must be on or after start date.")

    intervals: list[tuple[datetime, datetime]] = []
    current = start_dt

    while current <= end_dt:
        if current.month == 12:
            next_month_start = datetime(current.year + 1, 1, 1)
        else:
            next_month_start = datetime(current.year, current.month + 1, 1)
        month_end = next_month_start - timedelta(days=1)
        interval_end = min(month_end, end_dt)
        intervals.append((current, interval_end))
        current = interval_end + timedelta(days=1)

    return intervals


def format_duration_label(start_dt: datetime, end_dt: datetime) -> str:
    """Build a compact duration label for filenames (e.g. 15Nov2025-30Nov2025)."""
    return f"{start_dt.strftime('%d%b%Y')}-{end_dt.strftime('%d%b%Y')}"


def sanitize_filename(filename):
    """
    Sanitize a filename by removing invalid characters.
    
    Args:
        filename: The filename to sanitize
    
    Returns:
        str: Sanitized filename
    """
    # Remove invalid characters for Windows filenames
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', filename)
    # Remove leading/trailing spaces and dots
    sanitized = sanitized.strip(' .')
    return sanitized


def move_and_rename_zip_file(
    zip_file_path,
    plaza_name,
    target_folder="Life_cycle_files",
    from_date=None,
    to_date=None,
      ):
    """
    Move ZIP file from Downloads folder to target folder and rename with plaza + duration.
    
    Args:
        zip_file_path: Path to the downloaded ZIP file
        plaza_name: Plaza name (Plaza_name_IDFC) for the subfolder and filename
        target_folder: Folder to move the file to (default: "Life_cycle_files")
        from_date: Interval start date for the filename
        to_date: Interval end date for the filename
    
    Returns:
        str: Path to the moved file, or None if failed
    """
    try:
        zip_path = Path(zip_file_path)
        if not zip_path.exists():
            print(f"Error: ZIP file not found: {zip_file_path}")
            return None
        
        # Save under a Plaza_name_IDFC subfolder; filename stays plaza + duration
        target_path = Path(target_folder) / sanitize_filename(plaza_name)

        # Create target folder if it doesn't exist
        target_path.mkdir(parents=True, exist_ok=True)
        print(f"Target folder: {target_path.absolute()}")
        
        # Sanitize plaza name for filename
        sanitized_plaza = sanitize_filename(plaza_name)

        if from_date and to_date:
            duration_label = sanitize_filename(format_duration_label(from_date, to_date))
            new_filename = f"{sanitized_plaza}_{duration_label}.zip"
        else:
            new_filename = f"{sanitized_plaza}.zip"
        new_file_path = target_path / new_filename
        
        # If file with same name exists, add timestamp
        if new_file_path.exists():
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if from_date and to_date:
                new_filename = f"{sanitized_plaza}_{duration_label}_{timestamp}.zip"
            else:
                new_filename = f"{sanitized_plaza}_{timestamp}.zip"
            new_file_path = target_path / new_filename
            print(f"File with same name exists. Using: {new_filename}")
        
        # Move and rename the file
        print(f"Moving file from: {zip_path.name}")
        print(f"Moving file to: {new_file_path}")
        shutil.move(str(zip_path), str(new_file_path))
        
        print(f"✓ Successfully moved and renamed file: {new_file_path.name}")
        return str(new_file_path)
        
    except Exception as e:
        print(f"Error moving/renaming ZIP file: {e}")
        import traceback
        traceback.print_exc()
        return None


def download_idfc_file(username, password, from_date=None, to_date=None, plaza_name="Raibha", 
                       downloads_folder=r"C:\Users\Divyesh Parmar\Downloads", 
                       target_folder="Life_cycle_files",
                       headless=False, max_attempts=1, max_restarts=15):
    """
    Download Life Cycle Report file from IDFC portal and save it to target folder.
    
    Args:
        username: Username for IDFC login
        password: Password for IDFC login
        from_date: Start date (datetime object)
        to_date: End date (datetime object)
        plaza_name: Plaza name to select and use for folder/filename
        downloads_folder: Path to Downloads folder
        target_folder: Folder to save the downloaded file (default: "Life_cycle_files")
        headless: Run browser in headless mode
        max_attempts: Maximum login attempts per browser session
        max_restarts: Maximum browser restart attempts
    
    Returns:
        bool: True if download successful, False otherwise
    """
    for restart_count in range(1, max_restarts + 1):
        driver = None
        try:
            if restart_count > 1:
                print("\n" + "="*80)
                print(f"RESTARTING BROWSER - Attempt {restart_count} of {max_restarts}")
                print("="*80)
                print("Previous attempts failed. Restarting browser and retrying...")
                wait_for_idfc_session_slot(reason="browser restart")

            # Release any stale server-side session before opening a new browser.
            probe_release_stale_idfc_session(headless=headless)

            print("Setting up browser (fresh profile, no shared cookies/cache)...")
            driver = setup_driver(headless=headless, fresh_profile=True)

            print(f"Preparing login page: {IDFC_LOGIN_URL}")
            prepare_idfc_login_page(driver)

            wait = WebDriverWait(driver, 10)
            
            # Retry login up to max_attempts times
            login_success = False
            for attempt in range(1, max_attempts + 1):
                print(f"\n{'='*50}")
                print(f"Login Attempt {attempt} of {max_attempts}")
                print(f"{'='*50}")
                
                # Attempt login
                success = attempt_login(driver, wait, username, password)
                
                if success:
                    login_success = True
                    break
                else:
                    page_text = get_idfc_visible_page_text(driver)
                    session_limited = is_idfc_session_limit_message(page_text)
                    if session_limited:
                        print(
                            "\nSession limit detected. Clearing portal cookies/cache "
                            "and waiting before retry..."
                        )
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
            
            # If login succeeded, proceed with download
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
                    print("Closing browser and will restart the browser")
                    print("to retry the same plaza (up to max_restarts times).")
                    print("="*60)
                    
                    if driver:
                        try:
                            shutdown_idfc_driver(driver)
                            driver = None
                            print("\nBrowser closed.")
                        except:
                            pass
                    
                    login_success = False
                    break
                
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
                        wait_time = 180  # 1.5 minutes = 90 seconds
                        for remaining in range(wait_time, 0, -10):
                            print(f"  Waiting... {remaining} seconds remaining", end='\r')
                            time.sleep(10)
                        print(f"  Waiting complete. Proceeding to download...")
                        
                        # Wait for download table and click top download button
                        if click_top_download_button(driver, wait):
                            # Wait for ZIP file to be downloaded (browser stays open during download)
                            print("\nWaiting for ZIP file download to complete...")
                            print("(Browser will remain open during download)")
                            
                            # Detect the downloaded ZIP file (wait up to 2 minutes for download)
                            zip_file = None
                            max_wait_time = 120  # 2 minutes
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
                            
                            # Move and rename the downloaded file
                            if zip_file:
                                print("\n" + "="*60)
                                print("Moving and renaming downloaded file...")
                                print("="*60)
                                
                                moved_file = move_and_rename_zip_file(
                                    zip_file_path=zip_file,
                                    plaza_name=plaza_name,
                                    target_folder=target_folder,
                                    from_date=from_date,
                                    to_date=to_date,
                                )
                                
                                if moved_file:
                                    print(f"\n✓ File successfully saved to: {moved_file}")
                                    return True
                                else:
                                    print("\n✗ Failed to move/rename file.")
                                    return False
                            else:
                                print("Could not detect downloaded ZIP file within timeout.")
                                return False
                        else:
                            print("Failed to click download button.")
                            return False
                    else:
                        print("Failed to handle download popup.")
                        return False
                else:
                    print("Failed to click Export to Excel button.")
                    return False
                
                # If we reach here, download didn't complete successfully
                if driver:
                    try:
                        shutdown_idfc_driver(driver)
                    except:
                        pass
                    driver = None
                return False
            
            # If login failed after all attempts, close browser and restart
            if not login_success:
                if driver is None:
                    print(f"\n✗ Critical error occurred during processing.")
                else:
                    print(f"\n✗ All {max_attempts} login attempts failed in this browser session.")
                
                print("Closing browser and will restart the entire process for the same plaza...")
                
                if driver:
                    try:
                        shutdown_idfc_driver(driver)
                        driver = None
                    except:
                        pass
                
                if restart_count < max_restarts:
                    print(f"Will restart browser and retry entire process (restart {restart_count + 1}/{max_restarts})...")
                    wait_for_idfc_session_slot(reason="before next browser restart")
                    continue
                else:
                    print(f"\n✗ All {max_restarts} browser restart attempts exhausted.")
                    print("Could not complete download after multiple browser restarts.")
                    return False
        
        except Exception as e:
            print(f"Error during download process: {e}")
            import traceback
            traceback.print_exc()
            if driver:
                try:
                    shutdown_idfc_driver(driver)
                    driver = None
                except:
                    pass
            
            if restart_count < max_restarts:
                print(f"Error occurred. Will restart browser and retry (restart {restart_count + 1}/{max_restarts})...")
                wait_for_idfc_session_slot(reason="after error")
                continue
            else:
                return False
    
    # If we exit the loop, all restarts failed
    return False


def parse_config_date(date_value: str) -> datetime:
    """Parse a configured date string (YYYY-MM-DD or dd-mm-yyyy)."""
    date_str = str(date_value).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue

    parsed = parse_date_mmddyyyy(date_str)
    if parsed:
        return parsed

    raise ValueError(f"Could not parse date: {date_value}")


def process_plazas_from_excel(
    excel_file_path,
    username,
    password,
    downloads_folder,
    from_date,
    to_date,
    target_folder="Life_cycle_files",
   ):
    """
    Read Excel file and download files for all plazas that are not marked as 'done'.
    
    Args:
        excel_file_path: Path to the Excel file with plaza information
        username: Username for IDFC login
        password: Password for IDFC login
        downloads_folder: Path to Downloads folder
        from_date: Start date (datetime object)
        to_date: End date (datetime object)
        target_folder: Folder to save downloaded files (default: "Life_cycle_files")
    
    Returns:
        bool: True if all processing completed, False otherwise
    """
    try:
        if isinstance(from_date, str):
            from_date = parse_config_date(from_date)
        if isinstance(to_date, str):
            to_date = parse_config_date(to_date)

        # Read Excel file
        print("\n" + "="*60)
        print("Reading Excel file...")
        print("="*60)
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        # Validate required columns
        required_columns = ['Plaza_name_IDFC', 'Plaza_name_DB', 'Done']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Error: Missing required columns: {missing_columns}")
            return False
        
        print(f"Found {len(df)} rows in Excel file")

        month_intervals = split_date_range_to_month_intervals(from_date, to_date)

        print(f"\nOverall date range for all plazas:")
        print(f"  From: {from_date.strftime('%Y-%m-%d')} ({from_date.strftime('%B %d, %Y')})")
        print(f"  To: {to_date.strftime('%Y-%m-%d')} ({to_date.strftime('%B %d, %Y')})")
        print(f"\nSplit into {len(month_intervals)} month interval(s) (portal max ~1 month each):")
        for interval_start, interval_end in month_intervals:
            print(
                f"  {interval_start.strftime('%d-%b-%Y')} to "
                f"{interval_end.strftime('%d-%b-%Y')}"
            )
        
        # Filter rows that are not done and have plaza names
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
        
        if len(pending_rows) == 0:
            print("\n✓ All plazas are already marked as 'done'. Nothing to process.")
            return True
        
        print(f"\n" + "="*60)
        print(f"Starting download of {len(pending_rows)} plaza(s)...")
        print(f"Target folder: {target_folder}")
        print("="*60)
        
        # Process each plaza row
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
                print(
                    f"Overall Range: {from_date.strftime('%d-%b-%Y')} to "
                    f"{to_date.strftime('%d-%b-%Y')} "
                    f"({len(month_intervals)} month interval(s))"
                )
                print("="*80)

                plaza_interval_failures = 0
                for interval_index, (interval_start, interval_end) in enumerate(
                    month_intervals, start=1
                ):
                    print("\n" + "-" * 60)
                    print(
                        f"Month interval {interval_index}/{len(month_intervals)}: "
                        f"{interval_start.strftime('%d-%b-%Y')} to "
                        f"{interval_end.strftime('%d-%b-%Y')}"
                    )
                    print("-" * 60)

                    download_success = download_idfc_file(
                        username=username,
                        password=password,
                        from_date=interval_start,
                        to_date=interval_end,
                        plaza_name=plaza_name_idfc,
                        downloads_folder=downloads_folder,
                        target_folder=target_folder,
                        headless=False,
                    )

                    if not download_success:
                        print(
                            f"\n✗ Download failed for interval "
                            f"{interval_start.strftime('%d-%b-%Y')} to "
                            f"{interval_end.strftime('%d-%b-%Y')}."
                        )
                        plaza_interval_failures += 1
                        continue

                    print(
                        f"\n✓ Interval {interval_index}/{len(month_intervals)} saved as "
                        f"{plaza_name_idfc}_{format_duration_label(interval_start, interval_end)}.zip"
                    )
                    print("Waiting 8 seconds before next interval...")
                    time.sleep(8)

                if plaza_interval_failures > 0:
                    print("\n" + "="*80)
                    print(
                        f"PLAZA INCOMPLETE: {plaza_interval_failures} of "
                        f"{len(month_intervals)} month interval(s) failed"
                    )
                    print("="*80)
                    print("Skipping 'done' mark and continuing to next plaza...")
                    fail_count += 1
                    continue

                df.at[idx, 'Done'] = 'done'
                df.to_excel(excel_file_path, index=False, engine='openpyxl')
                print(
                    f"\n✓ Row {idx + 2}: all {len(month_intervals)} interval(s) downloaded. "
                    "Marked as 'done' in Excel."
                )
                success_count += 1

                print("Waiting 8 seconds before processing next plaza...")
                time.sleep(8)
                
            except Exception as e:
                print(f"\n✗ Error processing row {idx + 2}: {e}")
                import traceback
                traceback.print_exc()
                fail_count += 1
        
        # Final summary
        print("\n" + "="*80)
        print("DOWNLOAD SUMMARY")
        print("="*80)
        print(f"Total rows processed: {len(pending_rows)}")
        print(f"  ✓ Successful: {success_count}")
        print(f"  ✗ Failed: {fail_count}")
        print(f"Files saved under: {Path(target_folder).absolute()}")
        print("="*80)
        
        return fail_count == 0
        
    except Exception as e:
        print(f"Error reading/processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main function - download files for all plazas from Excel file"""
    # Configuration
    load_env_file()
    USERNAME = os.getenv("IDFC_USERNAME")
    PASSWORD = os.getenv("IDFC_PASSWORD")

    if not USERNAME or not PASSWORD:
        print(
            "\nError: IDFC_USERNAME and/or IDFC_PASSWORD not set. "
            "Add them to the project-root .env file."
        )
        return

    downloads_folder = r"C:\Users\Divyesh Parmar\Downloads"
    target_folder = "Life_cycle_files_from_July07_to_August09"  # Folder to save downloaded files

    # Excel file path - UPDATE THIS to your Excel file path
    excel_file_path = "Reference_File - download.xlsx"  # Change this to your actual file path

    try:
        from_date, to_date = get_process_date_range_datetimes()
    except (RuntimeError, ValueError) as exc:
        print(f"\nError: {exc}")
        return

    print("="*80)
    print("IDFC Downloader Bot - Download Files Only")
    print("="*80)
    print(f"Excel file: {excel_file_path}")
    print(f"Date range: {from_date.strftime('%Y-%m-%d')} to {to_date.strftime('%Y-%m-%d')}")
    print(f"Downloads folder: {downloads_folder}")
    print(f"Target folder: {target_folder}")
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
        downloads_folder=downloads_folder,
        from_date=from_date,
        to_date=to_date,
        target_folder=target_folder
    )
    
    if success:
        print("\n" + "="*80)
        print("All files downloaded successfully!")
        print("="*80)
    else:
        print("\n" + "="*80)
        print("Some downloads failed. Check the output above.")
        print("="*80)


if __name__ == "__main__":
    main()

