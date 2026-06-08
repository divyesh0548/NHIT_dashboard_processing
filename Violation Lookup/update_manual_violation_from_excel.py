"""
Update Manual Violation from Excel File

This script processes the violation_lookup_output Excel file and updates the database.
Input file has columns:
  - 'Reader Read Time': datetime in format '2026-01-13 03:20:33'
  - 'Violation Amt': the violation amount to set in manual_violation column

For each row:
1. Extracts date from 'Reader Read Time' column (e.g., '2026-01-13 03:20:33' -> '2026-01-13')
2. Uses the given plaza_name and extracted date to find matching row in submissions table
3. Updates manual_violation column in ONE row (keeping violation_amt unchanged)
"""

import pandas as pd
import psycopg2
from datetime import datetime
from pathlib import Path
import sys


ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from db_config import get_db_connection_kwargs


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


def find_column(df, column_names):
    """
    Find a column in DataFrame by trying multiple possible names.
    
    Args:
        df: DataFrame
        column_names: List of possible column names to try
    
    Returns:
        str: Column name if found, None otherwise
    """
    for col_name in column_names:
        # Try exact match (case-insensitive)
        for col in df.columns:
            if str(col).strip().lower() == col_name.lower():
                return col
        
        # Try partial match
        for col in df.columns:
            if col_name.lower() in str(col).lower():
                return col
    
    return None


def extract_date_from_reader_read_time(reader_read_time_value):
    """
    Extract date in yyyy-mm-dd format from 'Reader Read Time' column.
    
    The datetime format is: '2026-01-13 03:20:33'
    Extract the date part: '2026-01-13'
    
    Args:
        reader_read_time_value: Datetime value (string, datetime object, or pandas Timestamp)
    
    Returns:
        str: Date in yyyy-mm-dd format, or None if parsing fails
    """
    try:
        if pd.isna(reader_read_time_value) or reader_read_time_value is None:
            return None
        
        # If it's already a datetime object or pandas Timestamp, extract date directly
        if isinstance(reader_read_time_value, (datetime, pd.Timestamp)):
            return reader_read_time_value.strftime('%Y-%m-%d')
        
        # Convert to string for parsing
        value_str = str(reader_read_time_value).strip()
        
        if not value_str or value_str.lower() in ['nan', 'none', '']:
            return None
        
        # Try parsing different datetime formats
        # Format 1: '2026-01-13 03:20:33' (yyyy-mm-dd HH:MM:SS)
        try:
            # Extract date part (before space)
            date_part = value_str.split()[0]
            # Validate by parsing
            parsed_date = datetime.strptime(date_part, "%Y-%m-%d")
            return parsed_date.strftime('%Y-%m-%d')
        except (ValueError, IndexError):
            pass
        
        # Format 2: Try parsing the full string with datetime
        try:
            parsed_date = datetime.strptime(value_str, "%Y-%m-%d %H:%M:%S")
            return parsed_date.strftime('%Y-%m-%d')
        except ValueError:
            pass
        
        # Format 3: Try pandas to_datetime as fallback
        try:
            parsed_date = pd.to_datetime(value_str)
            if isinstance(parsed_date, pd.Timestamp):
                return parsed_date.strftime('%Y-%m-%d')
        except:
            pass
        
        return None
        
    except (ValueError, TypeError, IndexError):
        return None


def update_manual_violation(conn, plaza_name, date_str, violation_amount):
    """
    Update manual_violation column for ONE row (keeping violation_amt unchanged).
    STRICT RULE: Only ONE row per plaza_name and date combination should be updated.
    If a row with manual_violation already exists for this plaza_name and date, update that same row.
    Otherwise, update the first row found (by ID).
    
    Args:
        conn: Database connection
        plaza_name: Entity name (plaza name)
        date_str: Date in yyyy-mm-dd format
        violation_amount: Amount to set in manual_violation column
    
    Returns:
        bool: True if updated, False otherwise
    """
    try:
        cursor = conn.cursor()
        
        # First, check if there's already a row with manual_violation set for this plaza and date
        # This ensures we always update the SAME row if called multiple times
        cursor.execute("""
            SELECT id, manual_violation
            FROM public.submissions
            WHERE entity_name = %s 
            AND date = %s
            AND manual_violation IS NOT NULL
            ORDER BY id
            LIMIT 1
        """, (plaza_name, date_str))
        
        row = cursor.fetchone()
        
        # If no row with manual_violation exists, find the first row (by ID) for this plaza and date
        if not row:
            cursor.execute("""
                SELECT id, manual_violation
                FROM public.submissions
                WHERE entity_name = %s 
                AND date = %s
                ORDER BY id
                LIMIT 1
            """, (plaza_name, date_str))
            
            row = cursor.fetchone()
        
        if row:
            # Update the manual_violation column for this row
            update_id = row[0]
            current_value = row[1]
            cursor.execute("""
                UPDATE public.submissions
                SET manual_violation = %s
                WHERE id = %s
            """, (violation_amount, update_id))
            
            conn.commit()
            cursor.close()
            print(f"    ✓ Updated manual_violation:")
            print(f"      Row ID: {update_id}")
            print(f"      Previous value: {current_value}")
            print(f"      New value: {violation_amount}")
            print(f"      Plaza: {plaza_name}, Date: {date_str}")
            return True
        else:
            cursor.close()
            print(f"    ⊘ No row found for plaza and date. Skipping update.")
            print(f"      Plaza: {plaza_name}, Date: {date_str}, Amount: {violation_amount}")
            return False
        
    except psycopg2.Error as e:
        print(f"  Database error updating manual_violation: {e}")
        if conn:
            conn.rollback()
        return False
    except Exception as e:
        print(f"  Error updating manual_violation: {e}")
        return False


def process_excel_and_update_db(excel_file_path, plaza_name):
    """
    Process the violation_lookup_output Excel file and update database.
    
    Input file columns:
      - 'Reader Read Time': datetime in format '2026-01-13 03:20:33'
      - 'Violation Amt': the amount to set as manual_violation
    
    For each row:
      1. Extract date from 'Reader Read Time' (e.g., '2026-01-13 03:20:33' -> '2026-01-13')
      2. Check Violation Amt is valid and non-zero
      3. Update manual_violation in ONE row for plaza_name + date
    
    Args:
        excel_file_path: Path to violation_lookup_output Excel file
        plaza_name: Plaza name (entity_name in database)
    
    Returns:
        dict: Statistics about updates
    """
    # Check if Excel file exists
    if not Path(excel_file_path).exists():
        print(f"Error: Excel file not found: {excel_file_path}")
        return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
    
    # Connect to database
    conn = get_db_connection()
    if not conn:
        print("Failed to connect to database")
        return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
    
    try:
        # Read Excel file
        print(f"\nReading Excel file: {excel_file_path}")
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        print(f"Total rows in Excel file: {len(df)}")
        print(f"Columns found: {list(df.columns)}")
        
        # Find 'Reader Read Time' column
        reader_read_time_col = find_column(df, [
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
            'Reader Read Time.\t'
        ])
        
        if not reader_read_time_col:
            print("Error: Could not find 'Reader Read Time' column.")
            print(f"Available columns: {list(df.columns)}")
            return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
        
        print(f"Found Reader Read Time column: '{reader_read_time_col}'")
        
        # Find 'Violation Amt' column
        violation_amt_col = find_column(df, [
            'Violation Amt',
            'Violation Amt.',
            'Violation_Amt',
            'VIOLATION AMT',
            'Violation Amount',
            'Violation_Amount',
            'VIOLATION AMOUNT',
            'violation_amt',
            'ViolationAmt'
        ])
        
        if not violation_amt_col:
            print("Error: Could not find 'Violation Amt' column.")
            print(f"Available columns: {list(df.columns)}")
            return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
        
        print(f"Found Violation Amt column: '{violation_amt_col}'")
        
        # Find 'CCH TXN NO' column
        cch_txn_no_col = find_column(df, [
            'CCH TXN NO',
            'CCH TXN NO.',
            'CCH_TXN_NO',
            'CCH_TXN_NO.',
            'CCHTXNNO',
            'CCHTxnNo',
            'CCH Transaction No',
            'CCH Transaction No.',
            'CCH_Transaction_No',
            'CCH_Transaction_No.',
            'CCH Transaction Number',
            'CCH_Transaction_Number',
            'CCH TransactionNumber',
            'CCH TXN NUMBER',
            'CCHTXN NUMBER',
            'CCHTXN_NUMBER'
        ])
        
        if not cch_txn_no_col:
            print("Error: Could not find 'CCH TXN NO' column.")
            print(f"Available columns: {list(df.columns)}")
            return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
        
        print(f"Found CCH TXN NO column: '{cch_txn_no_col}'")
        
        # Filter out rows with empty/NaN Reader Read Time or Violation Amt
        valid_rows = df.dropna(subset=[reader_read_time_col, violation_amt_col])
        print(f"\nValid rows (non-empty Reader Read Time and Violation Amt): {len(valid_rows)}")
        
        if len(valid_rows) == 0:
            print("No valid rows to process. Exiting.")
            return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
        
        # Remove duplicate rows based on 'CCH TXN NO' column
        # Keep first occurrence, mark others as duplicates
        print(f"\nChecking for duplicate 'CCH TXN NO' values...")
        initial_count = len(valid_rows)
        
        # Identify duplicates (keep='first' means first occurrence is kept, others are marked as duplicates)
        duplicate_mask = valid_rows.duplicated(subset=[cch_txn_no_col], keep='first')
        unique_rows = valid_rows[~duplicate_mask].copy()
        duplicate_rows = valid_rows[duplicate_mask].copy()
        
        unique_count = len(unique_rows)
        duplicate_count = len(duplicate_rows)
        
        print(f"  Total rows: {initial_count}")
        print(f"  Unique rows (first occurrence): {unique_count}")
        print(f"  Duplicate rows (removed): {duplicate_count}")
        
        if duplicate_count > 0:
            print(f"\n  Duplicate 'CCH TXN NO' values found:")
            duplicate_cch_txn_nos = duplicate_rows[cch_txn_no_col].unique()
            for dup_value in duplicate_cch_txn_nos[:10]:  # Show first 10
                print(f"    - {dup_value}")
            if len(duplicate_cch_txn_nos) > 10:
                print(f"    ... and {len(duplicate_cch_txn_nos) - 10} more")
        
        # Process unique rows first
        print(f"\n{'='*80}")
        print("STEP 1: Processing UNIQUE rows (first occurrence of each CCH TXN NO)")
        print(f"{'='*80}")
        
        stats_unique = process_rows_and_update_db(conn, unique_rows, reader_read_time_col, violation_amt_col, plaza_name, "unique")
        
        # Process duplicate rows separately
        stats_duplicate = {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
        if duplicate_count > 0:
            print(f"\n{'='*80}")
            print("STEP 2: Processing DUPLICATE rows (removed duplicates)")
            print(f"{'='*80}")
            stats_duplicate = process_rows_and_update_db(conn, duplicate_rows, reader_read_time_col, violation_amt_col, plaza_name, "duplicate")
        
        # Combine statistics
        stats = {
            'processed': stats_unique['processed'] + stats_duplicate['processed'],
            'updated': stats_unique['updated'] + stats_duplicate['updated'],
            'skipped': stats_unique['skipped'] + stats_duplicate['skipped'],
            'errors': stats_unique['errors'] + stats_duplicate['errors']
        }
        
        return stats
        
    except Exception as e:
        print(f"Error processing Excel file: {e}")
        import traceback
        traceback.print_exc()
        return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
    
    finally:
        if conn:
            conn.close()


def process_rows_and_update_db(conn, rows_df, reader_read_time_col, violation_amt_col, plaza_name, row_type="unique"):
    """
    Process rows and update database.
    
    Args:
        conn: Database connection
        rows_df: DataFrame with rows to process
        reader_read_time_col: Column name for Reader Read Time
        violation_amt_col: Column name for Violation Amt
        plaza_name: Plaza name (entity_name in database)
        row_type: Type of rows being processed ("unique" or "duplicate")
    
    Returns:
        dict: Statistics about updates
    """
    try:
        stats = {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}
        
        # Track dates already updated to avoid updating the same date twice
        # For each unique date, we accumulate the total violation amount
        date_violation_map = {}
        
        for idx, row in rows_df.iterrows():
            try:
                reader_read_time = row[reader_read_time_col]
                violation_amt = row[violation_amt_col]
                
                # Extract date from Reader Read Time
                date_str = extract_date_from_reader_read_time(reader_read_time)
                
                if not date_str:
                    print(f"\nRow {idx + 2}: Could not extract date from '{reader_read_time}'. Skipping.")
                    stats['skipped'] += 1
                    continue
                
                # Validate violation amount
                try:
                    violation_value = float(violation_amt)
                except (ValueError, TypeError):
                    print(f"\nRow {idx + 2}: Invalid Violation Amt '{violation_amt}'. Skipping.")
                    stats['skipped'] += 1
                    continue
                
                if violation_value == 0:
                    print(f"\nRow {idx + 2}: Violation Amt is 0. Skipping.")
                    stats['skipped'] += 1
                    continue
                
                # Accumulate violation amounts per date
                if date_str in date_violation_map:
                    date_violation_map[date_str] += violation_value
                else:
                    date_violation_map[date_str] = violation_value
                
                stats['processed'] += 1
                
            except Exception as e:
                print(f"\nRow {idx + 2}: Error processing row: {e}")
                stats['errors'] += 1
                continue
        
        # Now update database once per unique date
        print(f"\nUpdating database for {row_type} rows...")
        print(f"Unique dates to update: {len(date_violation_map)}")
        
        for date_str, total_violation in sorted(date_violation_map.items()):
            print(f"\nDate: {date_str}, Total Violation Amt: {total_violation}")
            
            if update_manual_violation(conn, plaza_name, date_str, total_violation):
                stats['updated'] += 1
            else:
                stats['skipped'] += 1
        
        print(f"\n{row_type.capitalize()} rows processing summary:")
        print(f"  Rows processed: {stats['processed']}")
        print(f"  Unique dates updated: {stats['updated']}")
        print(f"  Skipped: {stats['skipped']}")
        print(f"  Errors: {stats['errors']}")
        
        return stats
        
    except Exception as e:
        print(f"Error processing {row_type} rows: {e}")
        import traceback
        traceback.print_exc()
        return {'processed': 0, 'updated': 0, 'skipped': 0, 'errors': 0}


def clear_manual_violation(plaza_name, date_from, date_to):
    """
    CLEANUP FUNCTION: Clear all manual_violation values for a given plaza_name and date range.
    This function is independent and should be run separately to clean up false values.
    
    Args:
        plaza_name: Entity name (plaza name)
        date_from: Start date in yyyy-mm-dd format
        date_to: End date in yyyy-mm-dd format (inclusive)
    
    Returns:
        dict: Statistics about cleared values
    """
    print("="*80)
    print("CLEANUP FUNCTION: Clear Manual Violation Values")
    print("="*80)
    print(f"Plaza Name: {plaza_name}")
    print(f"Date Range: {date_from} to {date_to}")
    print("="*80)
    
    # Connect to database
    conn = get_db_connection()
    if not conn:
        print("Failed to connect to database")
        return {'cleared': 0, 'errors': 0}
    
    try:
        cursor = conn.cursor()
        
        # First, count how many rows will be affected
        cursor.execute("""
            SELECT COUNT(*)
            FROM public.submissions
            WHERE entity_name = %s 
            AND date >= %s 
            AND date <= %s
            AND manual_violation IS NOT NULL
        """, (plaza_name, date_from, date_to))
        
        count_result = cursor.fetchone()
        rows_to_clear = count_result[0] if count_result else 0
        
        if rows_to_clear == 0:
            print(f"\nNo rows found with manual_violation values for {plaza_name} in date range {date_from} to {date_to}")
            cursor.close()
            conn.close()
            return {'cleared': 0, 'errors': 0}
        
        print(f"\nFound {rows_to_clear} row(s) with manual_violation values to clear")
        
        # Clear all manual_violation values for this plaza and date range
        cursor.execute("""
            UPDATE public.submissions
            SET manual_violation = NULL
            WHERE entity_name = %s 
            AND date >= %s 
            AND date <= %s
            AND manual_violation IS NOT NULL
        """, (plaza_name, date_from, date_to))
        
        rows_cleared = cursor.rowcount
        conn.commit()
        cursor.close()
        conn.close()
        
        print(f"\n✓ Successfully cleared {rows_cleared} manual_violation value(s)")
        print(f"  Plaza: {plaza_name}")
        print(f"  Date Range: {date_from} to {date_to}")
        
        return {'cleared': rows_cleared, 'errors': 0}
        
    except psycopg2.Error as e:
        print(f"\n✗ Database error clearing manual_violation: {e}")
        if conn:
            conn.rollback()
            conn.close()
        return {'cleared': 0, 'errors': 1}
    except Exception as e:
        print(f"\n✗ Error clearing manual_violation: {e}")
        import traceback
        traceback.print_exc()
        if conn:
            conn.close()
        return {'cleared': 0, 'errors': 1}


def main():
    """Main function"""
    # Normal operation: Get plaza name and Excel file path from command line arguments or use defaults
    if len(sys.argv) >= 3:
        plaza_name = sys.argv[1]
        excel_file_path = sys.argv[2]
    else:
        # Default values (update these as needed)
        plaza_name = "bahadrabad"
        excel_file_path = "daroda_combined_2025-11-23_to_2026-02-15.xlsx"
    
    # STEP 1: Clear existing manual_violation values FIRST before updating
    print("\n" + "="*80)
    print("STEP 1: Clearing existing manual_violation values")
    print("="*80)
    clear_result = clear_manual_violation(plaza_name, "2025-11-23", "2026-05-25")
    
    if clear_result['errors'] > 0:
        print("\n✗ Error during cleanup. Aborting to prevent inconsistent data.")
        return
    
    # print("\n✓ Cleanup complete. Proceeding to update...\n")
    
    # # STEP 2: Process Excel file and update database
    # print("="*80)
    # print("STEP 2: Update Manual Violation from Excel File")
    # print("="*80)
    # print(f"Plaza Name: {plaza_name}")
    # print(f"Excel File: {excel_file_path}")
    # print("="*80)
    
    # # Process Excel file and update database
    # stats = process_excel_and_update_db(excel_file_path, plaza_name)
    
    # print(f"\n{'='*80}")
    # print("FINAL SUMMARY")
    # print(f"{'='*80}")
    # print(f"  Total rows processed: {stats['processed']}")
    # print(f"  Total unique dates updated: {stats['updated']}")
    # print(f"  Total skipped: {stats['skipped']}")
    # print(f"  Total errors: {stats['errors']}")
    # print(f"{'='*80}")
    
    # if stats['updated'] > 0:
    #     print(f"\n✓ Successfully updated {stats['updated']} manual_violation values in database.")
    # else:
    #     print(f"\n⊘ No updates were made.")


if __name__ == "__main__":
    main()
