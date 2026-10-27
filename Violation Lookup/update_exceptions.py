"""
Script to:
1. Remove duplicate rows from an Excel file based on 'CCH TXN NO' column.
2. For each unique (date, shift) combination, count the number of rows.
3. Update the 'exceptions' column in the submissions table for the given plaza_name, date, and shift.
   - Only ONE row is updated per (plaza_name, date, shift) combination.
   - If 'exceptions' already has a value in any row for that combination, skip the update.

Usage:
    python update_exceptions.py
    (Configure excel_file_path and plaza_name in the __main__ block)
"""

import pandas as pd
import psycopg2
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


def remove_duplicates(excel_file_path: str) -> pd.DataFrame:
    """
    Read the Excel file, remove duplicate rows based on 'CCH TXN NO',
    update the file, and return the deduplicated DataFrame.

    Args:
        excel_file_path: Path to the input Excel file

    Returns:
        Deduplicated DataFrame
    """
    print(f"\n{'='*80}")
    print("STEP 1: Removing duplicate rows based on 'CCH TXN NO'")
    print(f"{'='*80}")

    df = pd.read_excel(excel_file_path, engine='openpyxl')
    total_rows_before = len(df)
    print(f"Total rows before deduplication: {total_rows_before}")

    if 'CCH TXN NO' not in df.columns:
        print("ERROR: 'CCH TXN NO' column not found in the Excel file.")
        print(f"Available columns: {list(df.columns)}")
        return df

    from Violation_lookup import normalize_txn_id_for_comparison

    # Normalize IDs for dedup (whitespace, backticks, special chars removed)
    df['_cch_txn_norm'] = df['CCH TXN NO'].apply(normalize_txn_id_for_comparison)

    # Separate rows with empty/null normalized CCH TXN NO
    empty_mask = df['_cch_txn_norm'] == ''
    df_empty = df[empty_mask]
    df_non_empty = df[~empty_mask]

    empty_count = len(df_empty)
    non_empty_count = len(df_non_empty)
    print(f"Rows with empty/null 'CCH TXN NO': {empty_count} (will be kept as-is)")
    print(f"Rows with non-empty 'CCH TXN NO': {non_empty_count}")

    # Count duplicates using normalized IDs (so `123 and 123 match as duplicates)
    duplicate_mask = df_non_empty.duplicated(subset=['_cch_txn_norm'], keep='first')
    duplicate_count = duplicate_mask.sum()

    if duplicate_count > 0:
        # Show some duplicate CCH TXN NOs
        duplicate_values = df_non_empty.loc[duplicate_mask, '_cch_txn_norm'].unique()
        print(f"Found {duplicate_count} duplicate row(s) based on 'CCH TXN NO'.")
        print(f"Number of unique duplicate CCH TXN NOs: {len(duplicate_values)}")
        # Print up to 10 examples
        for i, val in enumerate(duplicate_values[:10]):
            count = df_non_empty[df_non_empty['_cch_txn_norm'] == val].shape[0]
            print(f"  - '{val}' appears {count} time(s)")
        if len(duplicate_values) > 10:
            print(f"  ... and {len(duplicate_values) - 10} more")

        # Remove duplicates only from non-empty rows, keep first occurrence
        df_non_empty = df_non_empty.drop_duplicates(subset=['_cch_txn_norm'], keep='first')

        # Recombine: empty rows + deduplicated non-empty rows
        df = pd.concat([df_non_empty, df_empty], ignore_index=True)
        df = df.drop(columns=['_cch_txn_norm'])
        total_rows_after = len(df)
        print(f"\nTotal rows after deduplication: {total_rows_after}")
        print(f"Rows removed: {total_rows_before - total_rows_after}")

        # Save updated file
        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        print(f"Updated file saved: {excel_file_path}")
    else:
        print("No duplicate rows found based on 'CCH TXN NO'. File unchanged.")
        if '_cch_txn_norm' in df.columns:
            df = df.drop(columns=['_cch_txn_norm'])

    return df


def update_exceptions_in_db(df: pd.DataFrame, plaza_name: str):
    """
    For each unique (date, shift) combination in the DataFrame, count the rows
    and update the 'exceptions' column in the submissions table.

    Rules:
    - Match on entity_name = plaza_name, date, and shift.
    - If 'exceptions' already has a value in ANY row for that combination, skip.
    - Otherwise, update only ONE row (the first by id).

    Args:
        df: Deduplicated DataFrame with 'date' and 'shift' columns
        plaza_name: Plaza name to match against entity_name in DB
    """
    print(f"\n{'='*80}")
    print("STEP 2: Updating 'exceptions' in submissions table")
    print(f"{'='*80}")
    print(f"Plaza Name: {plaza_name}")

    # Validate required columns
    if 'date' not in df.columns or 'shift' not in df.columns:
        print("ERROR: 'date' and/or 'shift' columns not found in the Excel file.")
        print(f"Available columns: {list(df.columns)}")
        return

    # Group by (date, shift) and count rows
    grouped = df.groupby(['date', 'shift']).size().reset_index(name='row_count')
    print(f"\nFound {len(grouped)} unique (date, shift) combinations:")
    print("-" * 50)
    for _, row in grouped.iterrows():
        print(f"  date={row['date']}, shift={row['shift']}, row_count={row['row_count']}")
    print("-" * 50)

    # Connect to database
    conn = get_db_connection()
    if not conn:
        print("Failed to connect to database.")
        return

    stats = {
        'updated': 0,
        'skipped_already_exists': 0,
        'skipped_no_rows': 0,
        'errors': 0
    }

    try:
        cursor = conn.cursor()

        for _, row in grouped.iterrows():
            date_val = str(row['date']).strip()
            shift_val = str(row['shift']).strip()
            row_count = int(row['row_count'])

            print(f"\n  Processing: date={date_val}, shift={shift_val}, exceptions={row_count}")

            try:
                # Check if any row already has 'exceptions' set for this combination
                cursor.execute("""
                    SELECT id, exceptions
                    FROM public.submissions
                    WHERE entity_name = %s
                    AND date = %s
                    AND shift = %s
                    AND exceptions IS NOT NULL
                    AND CAST(exceptions AS TEXT) != ''
                    ORDER BY id
                    LIMIT 1
                """, (plaza_name, date_val, shift_val))

                existing_row = cursor.fetchone()

                if existing_row:
                    print(f"    SKIPPED: 'exceptions' already has value '{existing_row[1]}' "
                          f"in row id={existing_row[0]}")
                    stats['skipped_already_exists'] += 1
                    continue

                # Find the first row (by id) for this combination to update
                cursor.execute("""
                    SELECT id
                    FROM public.submissions
                    WHERE entity_name = %s
                    AND date = %s
                    AND shift = %s
                    ORDER BY id
                    LIMIT 1
                """, (plaza_name, date_val, shift_val))

                target_row = cursor.fetchone()

                if not target_row:
                    print(f"    SKIPPED: No rows found in DB for entity_name='{plaza_name}', "
                          f"date='{date_val}', shift='{shift_val}'")
                    stats['skipped_no_rows'] += 1
                    continue

                target_id = target_row[0]

                # Update only this one row
                cursor.execute("""
                    UPDATE public.submissions
                    SET exceptions = %s
                    WHERE id = %s
                """, (str(row_count), target_id))

                conn.commit()
                print(f"    UPDATED: row id={target_id}, exceptions set to {row_count}")
                stats['updated'] += 1

            except Exception as e:
                conn.rollback()
                print(f"    ERROR: {e}")
                stats['errors'] += 1

    except Exception as e:
        print(f"Database error: {e}")
    finally:
        conn.close()

    # Print summary
    print(f"\n{'='*80}")
    print("SUMMARY")
    print(f"{'='*80}")
    print(f"Total (date, shift) combinations: {len(grouped)}")
    print(f"Updated:                {stats['updated']}")
    print(f"Skipped (already set):  {stats['skipped_already_exists']}")
    print(f"Skipped (no DB rows):   {stats['skipped_no_rows']}")
    print(f"Errors:                 {stats['errors']}")
    print(f"{'='*80}")


def clear_exceptions(plaza_name: str, from_date: str, to_date: str):
    """
    Standalone cleanup function: Set 'exceptions' to NULL for all rows
    matching the given plaza_name and date range.

    Args:
        plaza_name: Plaza name to match against entity_name in DB
        from_date: Start date in yyyy-mm-dd format (inclusive)
        to_date: End date in yyyy-mm-dd format (inclusive)
    """
    print(f"\n{'='*80}")
    print("CLEARING 'exceptions' VALUES")
    print(f"{'='*80}")
    print(f"Plaza Name: {plaza_name}")
    print(f"Date Range: {from_date} to {to_date}")

    conn = get_db_connection()
    if not conn:
        print("Failed to connect to database.")
        return

    try:
        cursor = conn.cursor()

        # First, count how many rows will be affected
        cursor.execute("""
            SELECT COUNT(*)
            FROM public.submissions
            WHERE entity_name = %s
            AND date >= %s
            AND date <= %s
            AND exceptions IS NOT NULL
        """, (plaza_name, from_date, to_date))

        affected_count = cursor.fetchone()[0]
        print(f"Rows with 'exceptions' set: {affected_count}")

        if affected_count == 0:
            print("No rows to clear. Nothing to do.")
            return

        # Confirm and clear
        cursor.execute("""
            UPDATE public.submissions
            SET exceptions = NULL
            WHERE entity_name = %s
            AND date >= %s
            AND date <= %s
            AND exceptions IS NOT NULL
        """, (plaza_name, from_date, to_date))

        conn.commit()
        print(f"Cleared 'exceptions' for {affected_count} row(s).")

    except Exception as e:
        conn.rollback()
        print(f"Error: {e}")
    finally:
        conn.close()

    print(f"{'='*80}")


if __name__ == "__main__":
    # ============================================================================
    # CONFIGURATION
    # ============================================================================
    excel_file_path = "odhaki_paipkhar_combined_2026-06-01_to_2026-07-31.xlsx"
    plaza_name = "odhaki_paipkhar"
    # ============================================================================

    print(f"Input Excel File: {excel_file_path}")
    print(f"Plaza Name: {plaza_name}")

    clear_exceptions(plaza_name, "2026-06-01", "2026-07-31")

    # Step 1: Remove duplicates based on 'CCH TXN NO'
    df = remove_duplicates(excel_file_path)

    # Step 2: Update exceptions in DB for each (date, shift) combination
    update_exceptions_in_db(df, plaza_name)

