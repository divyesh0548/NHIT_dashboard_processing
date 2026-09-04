import pandas as pd
from datetime import datetime
import sys
import os
import psycopg2
import glob

from db_config import get_db_connection_kwargs


READER_READ_TIME_COLUMN_NAMES = [
    'Reader Read Time',
    'Reader Read Time ',
    'Reader Read Time\t',
    'ReaderReadTime',
    'ReaderReadTime ',
    'ReaderReadTime\t',
    'Reader_Read_Time',
    'Reader_Read_Time ',
    'Reader_Read_Time\t',
    'READER READ TIME',
    'READER READ TIME ',
    'READER READ TIME\t',
    'Reader Read Time.',
    'Reader Read Time. ',
    'Reader Read Time.\t',
    'Transaction Date/Time',
    'Transaction Date/Time ',
    'Transaction Date/Time\t',
    'Transaction Date/Time.',
    'Transaction Date/Time. ',
    'Transaction Date/Time.\t',
    'Transaction_Date_Time',
    'Transaction_Date_Time ',
    'Transaction_Date_Time\t',
    'TRANSACTION DATE/TIME',
    'TRANSACTION DATE/TIME ',
    'TRANSACTION DATE/TIME\t',
    'Transaction Date',
    'Transaction Date ',
    'Transaction Date\t',
    'Transaction_Date',
    'Transaction_Date ',
    'Transaction_Date\t',
    'TRANSACTION DATE',
    'TRANSACTION DATE ',
    'TRANSACTION DATE\t',
]

VIOLATION_COLUMN_NAMES = [
    'Violation Amt',
    'Violation Amt ',
    'Violation Amt\t',
    'Violation Amt.',
    'Violation Amt. ',
    'Violation Amt.\t',
    'Violation_Amt',
    'Violation_Amt ',
    'Violation_Amt\t',
    'Violation_Amt.',
    'Violation_Amt. ',
    'Violation_Amt.\t',
    'VIOLATION AMT',
    'VIOLATION AMT ',
    'VIOLATION AMT\t',
    'VIOLATION AMT.',
    'VIOLATION AMT. ',
    'VIOLATION AMT.\t',
    'Violation Amount',
    'Violation Amount ',
    'Violation Amount\t',
    'Violation_Amount',
    'Violation_Amount ',
    'Violation_Amount\t',
    'VIOLATION AMOUNT',
    'VIOLATION AMOUNT ',
    'VIOLATION AMOUNT\t',
    'violation_amt',
    'violation_amt ',
    'violation_amt\t',
    'ViolationAmt',
    'ViolationAmt ',
    'ViolationAmt\t',
    'Adjustment Amount',
    'Adjustment Amount ',
    'Adjustment Amount\t',
    'Adjustment_Amount',
    'Adjustment_Amount ',
    'Adjustment_Amount\t',
    'ADJUSTMENT AMOUNT',
    'ADJUSTMENT AMOUNT ',
    'ADJUSTMENT AMOUNT\t',
    'Debit Adjustment Amount (Rs.)',
    'Debit Adjustment Amount (Rs.) ',
    'Debit Adjustment Amount (Rs.)\t',
    'DEBIT ADJUSTMENT AMOUNT (RS.)',
    'DEBIT ADJUSTMENT AMOUNT (RS.) ',
    'DEBIT ADJUSTMENT AMOUNT (RS.)\t',
]

DATE_COLUMN_NAMES = [
    'Date',
    'Date ',
    'Date\t',
    'DATE',
    'DATE ',
    'DATE\t',
]

TOTAL_VIOLATION_AMT_COLUMN_NAMES = [
    'Total Violation amt',
    'Total Violation amt ',
    'Total Violation amt\t',
    'Total_Violation_amt',
    'Total_Violation_amt ',
    'Total_Violation_amt\t',
    'TOTAL VIOLATION AMT',
    'TOTAL VIOLATION AMT ',
    'TOTAL VIOLATION AMT\t',
]

PLAZA_NAME_IDFC_COLUMN_NAMES = [
    'Plaza_name_IDFC',
    'Plaza_name_IDFC ',
    'Plaza_name_IDFC\t',
    'Plaza Name IDFC',
    'Plaza Name IDFC ',
    'Plaza Name IDFC\t',
    'PLAZA NAME IDFC',
    'PLAZA NAME IDFC ',
    'PLAZA NAME IDFC\t',
]

PLAZA_NAME_DB_COLUMN_NAMES = [
    'Plaza_name_DB',
    'Plaza_name_DB ',
    'Plaza_name_DB\t',
    'Plaza Name DB',
    'Plaza Name DB ',
    'Plaza Name DB\t',
    'PLAZA NAME DB',
    'PLAZA NAME DB ',
    'PLAZA NAME DB\t',
]

DONE_COLUMN_NAMES = [
    'Done',
    'Done ',
    'Done\t',
    'DONE',
    'DONE ',
    'DONE\t',
]


def find_column(df, column_names):
    """
    Find a column in DataFrame by trying multiple possible names.
    """
    for col_name in column_names:
        for col in df.columns:
            if str(col).strip().lower() == col_name.lower():
                return col

        for col in df.columns:
            if col_name.lower() in str(col).lower():
                return col

    return None


def read_tabular_file(file_path, **kwargs):
    """
    Read Excel or CSV files using the appropriate pandas reader.
    """
    suffix = os.path.splitext(str(file_path))[1].lower()
    if suffix == ".csv":
        csv_kwargs = {"dtype": str}
        csv_kwargs.update(kwargs)
        return pd.read_csv(file_path, **csv_kwargs)
    return pd.read_excel(file_path, **kwargs)


def find_columns_in_first_100_rows(file_path):
    """
    Search first 100 rows to find 'Reader Read Time' and 'Violation amt' columns.
    Handles newlines, extra whitespace, and searches both headers and data rows.
    Returns tuple of (reader_read_time_col, violation_amt_col, header_row_idx) or (None, None, None) if not found.
    """
    try:
        # Read first 100 rows without assuming header row
        df_sample = read_tabular_file(file_path, nrows=100, header=None)
        
        reader_read_time_col_idx = None
        violation_amt_col_idx = None
        header_row_idx = None
        
        # Search through first 100 rows to find header row
        for row_idx in range(min(100, len(df_sample))):
            found_reader = False
            found_violation = False
            temp_reader_col = None
            temp_violation_col = None
            
            # Check each cell in this row
            for col_idx in range(len(df_sample.columns)):
                cell_value = df_sample.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    # Clean the cell value: remove newlines, carriage returns, and extra whitespace
                    col_clean = str(cell_value).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').strip()
                    col_clean_normalized = ' '.join(col_clean.split()).lower()  # Normalize whitespace
                    
                    if any(
                        alias.lower().strip() in col_clean_normalized
                        or col_clean_normalized in alias.lower().strip()
                        for alias in READER_READ_TIME_COLUMN_NAMES
                    ):
                        temp_reader_col = col_idx
                        found_reader = True
                    
                    if any(
                        alias.lower().strip() in col_clean_normalized
                        or col_clean_normalized in alias.lower().strip()
                        for alias in VIOLATION_COLUMN_NAMES
                    ):
                        temp_violation_col = col_idx
                        found_violation = True
            
            # If we found both columns in this row, this is likely the header row
            if found_reader and found_violation:
                reader_read_time_col_idx = temp_reader_col
                violation_amt_col_idx = temp_violation_col
                header_row_idx = row_idx
                break
        
        # If we found columns, return the column indices and header row
        if reader_read_time_col_idx is not None and violation_amt_col_idx is not None:
            return reader_read_time_col_idx, violation_amt_col_idx, header_row_idx
        
        # Fallback: try reading with default header (row 0)
        try:
            df_with_headers = read_tabular_file(file_path, nrows=1)
            for col in df_with_headers.columns:
                col_clean = str(col).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').strip()
                col_clean_normalized = ' '.join(col_clean.split()).lower()
                
                if any(
                    alias.lower().strip() in col_clean_normalized
                    or col_clean_normalized in alias.lower().strip()
                    for alias in READER_READ_TIME_COLUMN_NAMES
                ):
                    if reader_read_time_col_idx is None:
                        # Find the index of this column
                        reader_read_time_col_idx = list(df_with_headers.columns).index(col)
                        header_row_idx = 0
                
                if any(
                    alias.lower().strip() in col_clean_normalized
                    or col_clean_normalized in alias.lower().strip()
                    for alias in VIOLATION_COLUMN_NAMES
                ):
                    if violation_amt_col_idx is None:
                        violation_amt_col_idx = list(df_with_headers.columns).index(col)
                        if header_row_idx is None:
                            header_row_idx = 0
            
            if reader_read_time_col_idx is not None and violation_amt_col_idx is not None:
                return reader_read_time_col_idx, violation_amt_col_idx, header_row_idx
        except:
            pass
        
        return None, None, None
    except Exception as e:
        print(f"Error reading file to find columns: {e}")
        import traceback
        traceback.print_exc()
        return None, None, None


def extract_date_from_datetime(datetime_str):
    """
    Extract date from datetime string.
    Supports formats like '2025-12-15 00:00:58' and '12/11/2025 18:44:15'.
    Returns date string in format '2025-12-15' or None if parsing fails
    """
    if pd.isna(datetime_str) or datetime_str == '':
        return None
    
    try:
        # Convert to string if not already
        datetime_str = str(datetime_str).strip()

        # Try known explicit formats first.
        for fmt in (
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d',
            '%d/%m/%Y %H:%M:%S',
            '%d/%m/%Y',
        ):
            try:
                return datetime.strptime(datetime_str, fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue

        # Fallback for Excel/native datetime values and other parseable inputs.
        date_obj = pd.to_datetime(datetime_str, errors='coerce', dayfirst=True)
        if pd.notna(date_obj):
            return date_obj.strftime('%Y-%m-%d')

        return None
    except Exception:
        return None


def print_date_wise_violation_totals(date_totals, file_path=None):
    """
    Print date-wise violation totals for a processed file.
    """
    if not date_totals:
        return

    if file_path:
        print(f"\nDate-wise total violation amount for: {os.path.basename(file_path)}")
    else:
        print("\nDate-wise total violation amount:")

    for date in sorted(date_totals.keys()):
        try:
            formatted_date = datetime.strptime(date, '%Y-%m-%d').strftime('%d-%m-%Y')
        except Exception:
            formatted_date = date
        print(f"  {formatted_date}: {date_totals[date]}")


def process_excel_file(input_file_path, output_file_path=None):
    """
    Process Excel file to calculate date-wise totals of Violation amt.
    
    Args:
        input_file_path: Path to input Excel file
        output_file_path: Path to output Excel file (optional; file is written only when provided)
    """
    if not os.path.exists(input_file_path):
        print(f"Error: File '{input_file_path}' not found.")
        return False
    
    # Find columns in first 100 rows
    print("Searching for columns in first 100 rows...")
    reader_read_time_col_idx, violation_amt_col_idx, header_row_idx = find_columns_in_first_100_rows(input_file_path)
    
    if reader_read_time_col_idx is None:
        print("Error: 'Reader Read Time' column not found in first 100 rows.")
        print("\nDebug: Available columns in first few rows:")
        try:
            df_debug = read_tabular_file(input_file_path, nrows=5, header=None)
            for row_idx in range(min(5, len(df_debug))):
                print(f"\n  Row {row_idx}:")
                for col_idx in range(min(10, len(df_debug.columns))):
                    cell_val = df_debug.iloc[row_idx, col_idx]
                    if pd.notna(cell_val):
                        print(f"    Col {col_idx}: '{cell_val}' (repr: {repr(str(cell_val))})")
        except Exception as e:
            print(f"  Error: {e}")
        return False
    
    if violation_amt_col_idx is None:
        print("Error: 'Violation amt' column not found in first 100 rows.")
        print("\nDebug: Available columns in first few rows:")
        try:
            df_debug = read_tabular_file(input_file_path, nrows=5, header=None)
            for row_idx in range(min(5, len(df_debug))):
                print(f"\n  Row {row_idx}:")
                for col_idx in range(min(10, len(df_debug.columns))):
                    cell_val = df_debug.iloc[row_idx, col_idx]
                    if pd.notna(cell_val):
                        print(f"    Col {col_idx}: '{cell_val}' (repr: {repr(str(cell_val))})")
        except Exception as e:
            print(f"  Error: {e}")
        return False
    
    print(f"Found columns at indices: {reader_read_time_col_idx} and {violation_amt_col_idx} (header row: {header_row_idx})")
    
    # Read the entire file with the correct header row
    try:
        if header_row_idx is not None:
            df = read_tabular_file(input_file_path, header=header_row_idx)
            # Get column names by index
            reader_read_time_col = df.columns[reader_read_time_col_idx]
            violation_amt_col = df.columns[violation_amt_col_idx]
        else:
            # Fallback: read without header and use indices directly
            df = read_tabular_file(input_file_path, header=None)
            reader_read_time_col = reader_read_time_col_idx
            violation_amt_col = violation_amt_col_idx
        
        print(f"Using columns: '{reader_read_time_col}' and '{violation_amt_col}'")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    # Dictionary to store date-wise totals
    date_totals = {}
    
    # Process rows starting from where columns were found
    start_row = 0
    empty_row_count = 0
    max_empty_rows = 100
    
    print("Processing rows...")
    
    for idx, row in df.iterrows():
        # Get values from the identified columns
        reader_read_time = row[reader_read_time_col]
        violation_amt = row[violation_amt_col]
        
        # Check if Reader Read Time is empty/null
        if pd.isna(reader_read_time) or str(reader_read_time).strip() == '':
            # Increment empty row count and check if we should stop
            empty_row_count += 1
            if empty_row_count >= max_empty_rows:
                print(f"Stopped: Found {max_empty_rows} consecutive empty rows in 'Reader Read Time' column.")
                break
            continue  # Skip rows with empty Reader Read Time
        
        # Reset empty row count if Reader Read Time is not empty
        empty_row_count = 0
        
        # Extract date from Reader Read Time
        date_str = extract_date_from_datetime(reader_read_time)
        
        if date_str is None:
            print(f"Date extraction failed for row {idx}: {reader_read_time}")
            continue  # Skip rows where date extraction fails

        # print(f"Date extracted for row {idx}: {date_str}")
        
        # Convert Violation amt to native Python float (avoid np.float64 in DB binds)
        # If violation_amt is empty, treat it as 0
        try:
            if pd.isna(violation_amt) or str(violation_amt).strip() == '':
                amount = 0.0
            else:
                amount = pd.to_numeric(violation_amt, errors='coerce')
            if pd.isna(amount):
                amount = 0.0
            else:
                amount = float(amount)
        except Exception:
            amount = 0.0
        
        # Add to date-wise totals (even if amount is 0, we still record the date)
        if date_str in date_totals:
            date_totals[date_str] += amount
        else:
            date_totals[date_str] = amount
    
    # Check if we have any data
    if not date_totals:
        print("No valid data found to process.")
        return None

    print(f"\nProcessing complete!")
    print(f"Total dates processed: {len(date_totals)}")
    print_date_wise_violation_totals(date_totals, input_file_path)

    if output_file_path:
        output_data = []
        for date in sorted(date_totals.keys()):
            try:
                date_obj = datetime.strptime(date, '%Y-%m-%d')
                formatted_date = date_obj.strftime('%d-%m-%Y')
            except Exception:
                formatted_date = date

            output_data.append({
                'Date': formatted_date,
                'Total Violation amt': date_totals[date]
            })

        output_df = pd.DataFrame(output_data)
        try:
            output_df.to_excel(output_file_path, index=False)
            print(f"Output saved to: {output_file_path}")
        except Exception as e:
            print(f"Error writing output file: {e}")
            return None

    return date_totals


def process_folder_excel_files(folder_path, output_file_path=None):
    """
    Process all Excel/CSV files in a folder and return combined date-wise totals.
    
    Args:
        folder_path: Path to folder containing Excel/CSV files
        output_file_path: Path to output Excel file (optional; file is written only when provided)
    
    Returns:
        dict: Combined date_totals dictionary with all dates and their totals, or None if error
    """
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' not found.")
        return None
    
    if not os.path.isdir(folder_path):
        print(f"Error: '{folder_path}' is not a directory.")
        return None
    
    # Find all supported files in the folder
    file_patterns = ['*.xlsx', '*.xls', '*.csv', '*.XLSX', '*.XLS', '*.CSV']
    report_files = []
    for pattern in file_patterns:
        report_files.extend(glob.glob(os.path.join(folder_path, pattern)))

    if not report_files:
        print(f"No Excel/CSV files found in folder: {folder_path}")
        return None

    # Remove duplicates and sort
    report_files = sorted(list(set(report_files)))

    print(f"\nFound {len(report_files)} Excel/CSV file(s) in folder: {folder_path}")
    print("=" * 60)
    
    # Combined date_totals dictionary
    combined_date_totals = {}
    processed_count = 0
    failed_count = 0
    
    # Process each supported file
    for idx, report_file in enumerate(report_files, 1):
        print(f"\n[{idx}/{len(report_files)}] Processing: {os.path.basename(report_file)}")
        print("-" * 60)
        
        # Process file (without generating individual output files)
        date_totals = process_excel_file(report_file, output_file_path=None)
        
        if date_totals:
            # Merge date_totals into combined dictionary
            for date, amount in date_totals.items():
                if date in combined_date_totals:
                    combined_date_totals[date] += amount
                else:
                    combined_date_totals[date] = amount
            processed_count += 1
            print(f"✓ Successfully processed: {os.path.basename(report_file)}")
            print(f"  Dates found: {len(date_totals)}")
        else:
            failed_count += 1
            print(f"✗ Failed to process: {os.path.basename(report_file)}")
    
    # Check if we have any combined data
    if not combined_date_totals:
        print(f"\nNo valid data found in any of the Excel/CSV files.")
        return None
    
    print("\n" + "=" * 60)
    print(f"Processing Summary:")
    print(f"  Total files: {len(report_files)}")
    print(f"  Successfully processed: {processed_count}")
    print(f"  Failed: {failed_count}")
    print(f"  Total unique dates: {len(combined_date_totals)}")
    print("=" * 60)

    if output_file_path:
        output_data = []
        for date in sorted(combined_date_totals.keys()):
            try:
                date_obj = datetime.strptime(date, '%Y-%m-%d')
                formatted_date = date_obj.strftime('%d-%m-%Y')
            except Exception:
                formatted_date = date

            output_data.append({
                'Date': formatted_date,
                'Total Violation amt': combined_date_totals[date]
            })

        output_df = pd.DataFrame(output_data)
        try:
            output_df.to_excel(output_file_path, index=False)
            print(f"\n✓ Combined output saved to: {output_file_path}")
            print(f"  Total dates in output: {len(combined_date_totals)}")
        except Exception as e:
            print(f"Error writing combined output file: {e}")
            import traceback
            traceback.print_exc()
            return None

    return combined_date_totals


def get_db_connection():
    """Create and return a PostgreSQL database connection"""
    try:
        conn = psycopg2.connect(**get_db_connection_kwargs())
        return conn
    except RuntimeError as e:
        print(f"Database configuration error: {e}")
        return None
    except psycopg2.Error as e:
        print(f"Error connecting to database: {e}")
        return None


def update_violation_amt_in_db(plaza_name, date_totals):
    """
    Update violation_amt column in submissions table based on date-wise totals.
    
    Args:
        plaza_name: Name of the plaza (entity_name in database)
        date_totals: Dictionary with date (YYYY-MM-DD) as key and violation amount as value
    
    Returns:
        dict: Statistics about updates (updated_count, skipped_count, etc.)
    """
    if not date_totals:
        print(f"No date totals provided for plaza: {plaza_name}")
        return {'updated': 0, 'skipped': 0, 'already_exists': 0}
    
    conn = get_db_connection()
    if not conn:
        print(f"Failed to connect to database for plaza: {plaza_name}")
        return {'updated': 0, 'skipped': 0, 'already_exists': 0}
    
    stats = {'updated': 0, 'skipped': 0, 'already_exists': 0, 'no_valid_rows': 0}
    
    try:
        cursor = conn.cursor()
        
        for date_str, violation_amount in date_totals.items():
            # Cast numpy/pandas scalars to native float so psycopg2 binds correctly
            try:
                violation_amount = float(violation_amount)
            except (TypeError, ValueError):
                print(f"  Skipping {plaza_name} - {date_str}: invalid violation amount {violation_amount!r}")
                stats['skipped'] += 1
                continue

            # Skip if violation amount is 0 or negative
            if violation_amount <= 0:
                continue
            
            # Check if any row for this date+plaza already has violation_amt
            cursor.execute("""
                SELECT COUNT(*) 
                FROM public.submissions
                WHERE entity_name = %s 
                AND date = %s 
                AND violation_amt IS NOT NULL
            """, (plaza_name, date_str))
            
            existing_count = cursor.fetchone()[0]
            if existing_count > 0:
                print(f"  Skipping {plaza_name} - {date_str}: violation_amt already exists in {existing_count} row(s)")
                stats['already_exists'] += 1
                continue
            
            # Get all submissions for this plaza and date
            cursor.execute("""
                SELECT id, file_link, exception_value
                FROM public.submissions
                WHERE entity_name = %s 
                AND date = %s
                ORDER BY id
            """, (plaza_name, date_str))
            
            rows = cursor.fetchall()
            
            if not rows:
                print(f"  No submissions found for {plaza_name} - {date_str}")
                stats['skipped'] += 1
                continue
            
            # Filter rows based on conditions:
            # 1. file_link must NOT be NULL and NOT be empty
            # 2. exception_value must NOT be NULL, NOT be empty, and NOT be 0 (check both int and float)
            valid_rows = []
            for row in rows:
                row_id = row[0]
                file_link = row[1]
                exception_value = row[2]
                
                # Check file_link condition
                # Commented out because file link can be null and exception value can be non-zero.
                # if file_link is None or (isinstance(file_link, str) and file_link.strip() == ''):
                #     continue
                
                # Check exception_value condition
                if exception_value is None:
                    continue
                
                # Check if exception_value is empty string
                if isinstance(exception_value, str) and exception_value.strip() == '':
                    continue
                
                # Check if exception_value is 0 (handle both int and float)
                try:
                    if isinstance(exception_value, (int, float)):
                        if exception_value == 0:
                            continue
                    elif isinstance(exception_value, str):
                        # Try to convert string to number
                        num_value = float(exception_value.strip())
                        if num_value == 0:
                            continue
                except (ValueError, AttributeError):
                    # If we can't determine if it's 0, skip this row to be safe
                    continue
                
                # This row is valid
                valid_rows.append(row_id)
            
            # If no valid rows found, skip this date+plaza
            if not valid_rows:
                print(f"  Skipping {plaza_name} - {date_str}: No valid rows (all have empty file_link or exception_value is null/0)")
                stats['no_valid_rows'] += 1
                continue
            
            # Update violation_amt in the first valid row only
            update_id = valid_rows[0]
            cursor.execute("""
                UPDATE public.submissions
                SET violation_amt = %s
                WHERE id = %s
            """, (violation_amount, update_id))
            
            conn.commit()
            print(f"  Updated {plaza_name} - {date_str}: violation_amt = {violation_amount} (row ID: {update_id})")
            stats['updated'] += 1
        
        cursor.close()
        conn.close()
        
        print(f"\nUpdate summary for {plaza_name}:")
        print(f"  Updated: {stats['updated']}")
        print(f"  Skipped (already exists): {stats['already_exists']}")
        print(f"  Skipped (no valid rows): {stats['no_valid_rows']}")
        print(f"  Skipped (other): {stats['skipped']}")
        
        return stats
        
    except psycopg2.Error as e:
        print(f"Database error while updating violation_amt for {plaza_name}: {e}")
        if conn:
            conn.rollback()
            conn.close()
        return stats
    except Exception as e:
        print(f"Unexpected error while updating violation_amt for {plaza_name}: {e}")
        import traceback
        traceback.print_exc()
        if conn:
            conn.close()
        return stats


def delete_violation_amt(plaza_name, from_date, to_date):
    """
    Delete (set to NULL) all violation_amt values for a specific plaza within a date range.
    
    This function is independent and can be run separately from the main flow.
    
    Args:
        plaza_name: Name of the plaza (entity_name in database)
        from_date: Start date (datetime object or string in 'YYYY-MM-DD' format)
        to_date: End date (datetime object or string in 'YYYY-MM-DD' format)
    
    Returns:
        dict: Statistics about the deletion operation with keys:
            - deleted: Number of rows where violation_amt was set to NULL
            - not_found: Number of rows that didn't have violation_amt (already NULL)
            - total_affected: Total number of rows in the date range
            - success: Boolean indicating if operation was successful
    """
    print("\n" + "="*80)
    print("DELETE VIOLATION_AMT FUNCTION")
    print("="*80)
    
    # Convert dates to string format if they are datetime objects
    if isinstance(from_date, datetime):
        from_date_str = from_date.strftime('%Y-%m-%d')
    else:
        from_date_str = str(from_date).strip()
    
    if isinstance(to_date, datetime):
        to_date_str = to_date.strftime('%Y-%m-%d')
    else:
        to_date_str = str(to_date).strip()
    
    print(f"Plaza Name: {plaza_name}")
    print(f"Date Range: {from_date_str} to {to_date_str}")
    print("="*80)
    
    # Connect to database
    conn = get_db_connection()
    if not conn:
        print("Failed to connect to database.")
        return {'deleted': 0, 'not_found': 0, 'total_affected': 0, 'success': False}
    
    stats = {'deleted': 0, 'not_found': 0, 'total_affected': 0, 'success': False}
    
    try:
        cursor = conn.cursor()
        
        # First, count total rows in the date range
        cursor.execute("""
            SELECT COUNT(*) 
            FROM public.submissions
            WHERE entity_name = %s 
            AND date >= %s 
            AND date <= %s
        """, (plaza_name, from_date_str, to_date_str))
        
        total_rows = cursor.fetchone()[0]
        stats['total_affected'] = total_rows
        
        print(f"\nTotal rows found in date range: {total_rows}")
        
        if total_rows == 0:
            print("No rows found for the specified plaza and date range.")
            stats['success'] = True
            cursor.close()
            conn.close()
            return stats
        
        # Count rows that have violation_amt (not NULL)
        cursor.execute("""
            SELECT COUNT(*) 
            FROM public.submissions
            WHERE entity_name = %s 
            AND date >= %s 
            AND date <= %s
            AND violation_amt IS NOT NULL
        """, (plaza_name, from_date_str, to_date_str))
        
        rows_with_violation_amt = cursor.fetchone()[0]
        stats['not_found'] = total_rows - rows_with_violation_amt
        
        print(f"Rows with violation_amt (will be deleted): {rows_with_violation_amt}")
        print(f"Rows without violation_amt (already NULL): {stats['not_found']}")
        
        if rows_with_violation_amt == 0:
            print("\nNo violation_amt values to delete. All values are already NULL.")
            stats['success'] = True
            cursor.close()
            conn.close()
            return stats
        
        # Ask for confirmation (optional - you can remove this if you want automatic deletion)
        print(f"\n⚠️  WARNING: This will set violation_amt to NULL for {rows_with_violation_amt} row(s).")
        print("Proceeding with deletion...")
        
        # Update all rows: set violation_amt to NULL
        cursor.execute("""
            UPDATE public.submissions
            SET violation_amt = NULL
            WHERE entity_name = %s 
            AND date >= %s 
            AND date <= %s
            AND violation_amt IS NOT NULL
        """, (plaza_name, from_date_str, to_date_str))
        
        deleted_count = cursor.rowcount
        conn.commit()
        
        stats['deleted'] = deleted_count
        stats['success'] = True
        
        print(f"\n✓ Successfully deleted violation_amt for {deleted_count} row(s).")
        print("="*80)
        
        cursor.close()
        conn.close()
        
        return stats
        
    except psycopg2.Error as e:
        print(f"\n✗ Database error while deleting violation_amt: {e}")
        if conn:
            conn.rollback()
            conn.close()
        stats['success'] = False
        return stats
    except Exception as e:
        print(f"\n✗ Unexpected error while deleting violation_amt: {e}")
        import traceback
        traceback.print_exc()
        if conn:
            conn.close()
        stats['success'] = False
        return stats


def read_processed_excel_and_update_db(input_file_path, plaza_name):
    if not os.path.exists(input_file_path):
        print(f"Error: File '{input_file_path}' not found.")
        return {'updated': 0, 'skipped': 0, 'already_exists': 0}
    
    try:
        # Read the processed Excel file
        df = pd.read_excel(input_file_path)
        
        # Check if required columns exist
        date_col = find_column(df, DATE_COLUMN_NAMES)
        total_violation_amt_col = find_column(df, TOTAL_VIOLATION_AMT_COLUMN_NAMES)

        if not date_col or not total_violation_amt_col:
            print(f"Error: File must contain 'Date' and 'Total Violation amt' columns.")
            print(f"Found columns: {list(df.columns)}")
            return {'updated': 0, 'skipped': 0, 'already_exists': 0}
        
        # Convert date format and create date_totals dictionary
        date_totals = {}
        for idx, row in df.iterrows():
            date_str = str(row[date_col]).strip()
            violation_amt = row[total_violation_amt_col]
            
            # Convert date from DD-MM-YYYY to YYYY-MM-DD
            try:
                # Try parsing DD-MM-YYYY format first
                if '-' in date_str and len(date_str.split('-')) == 3:
                    date_parts = date_str.split('-')
                    if len(date_parts[0]) == 2:  # DD-MM-YYYY format
                        date_obj = datetime.strptime(date_str, '%d-%m-%Y')
                        formatted_date = date_obj.strftime('%Y-%m-%d')
                    else:  # YYYY-MM-DD format
                        formatted_date = date_str
                else:
                    # Try parsing as date object
                    date_obj = pd.to_datetime(date_str, errors='coerce')
                    if pd.notna(date_obj):
                        formatted_date = date_obj.strftime('%Y-%m-%d')
                    else:
                        print(f"Warning: Could not parse date '{date_str}' in row {idx+1}, skipping...")
                        continue
            except Exception as e:
                print(f"Warning: Error parsing date '{date_str}' in row {idx+1}: {e}, skipping...")
                continue
            
            # Convert violation amount to numeric
            try:
                amount = pd.to_numeric(violation_amt, errors='coerce')
                if pd.isna(amount):
                    print(f"Warning: Invalid violation amount in row {idx+1}, skipping...")
                    continue
                date_totals[formatted_date] = float(amount)
            except Exception as e:
                print(f"Warning: Error converting violation amount in row {idx+1}: {e}, skipping...")
                continue
        
        if not date_totals:
            print("No valid data found in the Excel file.")
            return {'updated': 0, 'skipped': 0, 'already_exists': 0}
        
        print(f"\nRead {len(date_totals)} date entries from file: {input_file_path}")
        print(f"Updating database for plaza: {plaza_name}")
        
        # Update database
        stats = update_violation_amt_in_db(plaza_name, date_totals)
        
        return stats
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        import traceback
        traceback.print_exc()
        return {'updated': 0, 'skipped': 0, 'already_exists': 0}


def process_and_update_db(input_file_path, plaza_name, output_file_path=None):
    """
    Process Excel file and update database with violation amounts.
    
    Args:
        input_file_path: Path to input Excel file
        plaza_name: Name of the plaza (entity_name in database)
        output_file_path: Path to output Excel file (optional)
    
    Returns:
        dict: Statistics about updates
    """
    print(f"\nProcessing file for plaza: {plaza_name}")
    print(f"File: {input_file_path}")
    
    # Process Excel file and get date-wise totals
    date_totals = process_excel_file(input_file_path, output_file_path)
    
    if date_totals is None:
        print(f"Failed to process file for plaza: {plaza_name}")
        return {'updated': 0, 'skipped': 0, 'already_exists': 0}
    
    # Update database
    print(f"\nUpdating database for plaza: {plaza_name}")
    stats = update_violation_amt_in_db(plaza_name, date_totals)
    
    return stats


def process_plazas_from_reference_excel(excel_file_path, from_date, to_date, base_folder="All_life_cycles"):
    """
    Process plazas from reference Excel file row by row.
    Similar to IDFC downloader flow: skip rows marked 'done', process pending rows,
    and mark 'Done' only after successful DB update.

    Args:
        excel_file_path: Path to reference Excel file
        from_date: Hardcoded start date for delete_violation_amt
        to_date: Hardcoded end date for delete_violation_amt
        base_folder: Base folder containing extracted life cycle folders

    Returns:
        bool: True if all pending rows processed successfully, False otherwise
    """
    try:
        print("\n" + "=" * 80)
        print("Reading reference Excel file...")
        print("=" * 80)
        df = pd.read_excel(excel_file_path, engine='openpyxl')
    except Exception as e:
        print(f"Error reading reference Excel file: {e}")
        import traceback
        traceback.print_exc()
        return False

    plaza_name_idfc_col = find_column(df, PLAZA_NAME_IDFC_COLUMN_NAMES)
    plaza_name_db_col = find_column(df, PLAZA_NAME_DB_COLUMN_NAMES)
    done_col = find_column(df, DONE_COLUMN_NAMES)

    missing_columns = []
    if not plaza_name_idfc_col:
        missing_columns.append('Plaza_name_IDFC')
    if not plaza_name_db_col:
        missing_columns.append('Plaza_name_DB')
    if not done_col:
        missing_columns.append('Done')

    if missing_columns:
        print(f"Error: Missing required columns in reference file: {missing_columns}")
        return False

    all_plaza_rows = df[
        (df[plaza_name_idfc_col].notna()) &
        (df[plaza_name_db_col].notna())
    ]

    done_rows = all_plaza_rows[
        (all_plaza_rows[done_col].notna()) &
        (all_plaza_rows[done_col].astype(str).str.strip().str.lower() == 'done')
    ]

    pending_rows = all_plaza_rows[
        (all_plaza_rows[done_col].isna()) |
        (all_plaza_rows[done_col].astype(str).str.strip().str.lower() != 'done')
    ]

    print("\n" + "=" * 80)
    print("PLAZA STATUS SUMMARY")
    print("=" * 80)
    print(f"Total plaza rows found: {len(all_plaza_rows)}")
    print(f"  Already processed (done): {len(done_rows)}")
    print(f"  Remaining to process: {len(pending_rows)}")
    print("=" * 80)

    if len(pending_rows) == 0:
        print("All plazas already marked as 'done'. Nothing to process.")
        return True

    success_count = 0
    fail_count = 0

    for idx, row in pending_rows.iterrows():
        try:
            plaza_name_idfc = str(row[plaza_name_idfc_col]).strip()
            plaza_name_db = str(row[plaza_name_db_col]).strip().lower()
            if not plaza_name_idfc or not plaza_name_db:
                print(f"Missing plaza name in row {idx + 2}, skipping.")
                fail_count += 1
                continue
            folder_path = os.path.join(base_folder, plaza_name_idfc)

            print("\n" + "=" * 80)
            print(f"Processing row {idx + 2}")
            print("=" * 80)
            print(f"Plaza_name_IDFC (folder): {plaza_name_idfc}")
            print(f"Plaza_name_DB (database): {plaza_name_db}")
            print(f"Folder path: {folder_path}")
            print(f"Hardcoded date range: {from_date} to {to_date}")
            print("=" * 80)

            if not os.path.isdir(folder_path):
                print(f"Folder not found, skipping row {idx + 2}: {folder_path}")
                fail_count += 1
                continue

            # Step 1: delete old violation_amt for date range
            delete_stats = delete_violation_amt(plaza_name_db, from_date, to_date)
            if not delete_stats.get('success', False):
                print(f"Failed delete step for row {idx + 2}. Skipping.")
                fail_count += 1
                continue

            # Step 2: process extracted files
            combined_totals = process_folder_excel_files(folder_path)
            if combined_totals is None:
                print(f"Failed processing folder for row {idx + 2}. Skipping.")
                fail_count += 1
                continue

            # Step 3: update DB directly from in-memory totals
            print(f"\nUpdating database for plaza: {plaza_name_db}")
            stats = update_violation_amt_in_db(plaza_name_db, combined_totals)
            if not isinstance(stats, dict):
                print(f"Failed DB update step for row {idx + 2}. Skipping.")
                fail_count += 1
                continue

            # Mark done only after successful completion of all steps
            df.at[idx, done_col] = 'done'
            df.to_excel(excel_file_path, index=False, engine='openpyxl')
            print(f"Row {idx + 2} processed successfully. Marked as 'done'.")
            success_count += 1

        except Exception as e:
            print(f"Error processing row {idx + 2}: {e}")
            import traceback
            traceback.print_exc()
            fail_count += 1

    print("\n" + "=" * 80)
    print("PROCESSING SUMMARY")
    print("=" * 80)
    print(f"Total pending rows: {len(pending_rows)}")
    print(f"  Successful: {success_count}")
    print(f"  Failed: {fail_count}")
    print("=" * 80)

    return fail_count == 0


def main():
    """
    Main function to run the script.
    NOTE: This function only processes Excel files and does NOT update the database.
    Database updates must be done separately by calling update_violation_amt_in_db().
    """
    # if len(sys.argv) < 2:
    #     print("Usage: python extraction.py <input_excel_file> [output_excel_file]")
    #     print("\nExample:")
    #     print("  python extraction.py test_excel.xlsx")
    #     print("  python extraction.py test_excel.xlsx output.xlsx")
    #     print("\nNOTE: This script only processes files. Database updates must be done separately.")
    #     return

    # Hardcoded date range (Excel Date column is intentionally ignored)
    # from_date = "2025-11-11"
    from_date = "2026-07-07"
    to_date = "2026-08-09"

    # Reference file (same file used by IDFC flow)
    excel_file_path = "Reference_File.xlsx"

    success = process_plazas_from_reference_excel(
        excel_file_path=excel_file_path,
        from_date=from_date,
        to_date=to_date,
        base_folder="YES July7_to_Aug9"
        # base_folder="ICICI Bank Files"
    )

    if success:
        print("\nAll pending plazas processed successfully.")
    else:
        print("\nSome plazas failed. Check logs above.")
    
    # input_file = sys.argv[1]
    # output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Only process file - NO database update
    #process_excel_file(input_file, output_file)


# Safety: Only run main() when script is executed directly, not when imported
if __name__ == "__main__":
    main()
