import pandas as pd
import psycopg2
from plaza_rates import PLAZA_RATES
from pathlib import Path
import sys


ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from db_config import get_db_connection_kwargs


JOURNEY_TYPE_COLUMN_CANDIDATES = ["Journey Type", "Trip Type"]


def find_first_matching_column(df: pd.DataFrame, column_candidates):
    for candidate in column_candidates:
        for col in df.columns:
            if str(col).strip().lower() == candidate.lower():
                return col

    for candidate in column_candidates:
        for col in df.columns:
            if candidate.lower() in str(col).lower():
                return col

    return None


def get_rate_section_from_journey_type(journey_type_value) -> str:
    journey_type = (
        str(journey_type_value).strip().lower()
        if pd.notna(journey_type_value)
        else ""
    )

    if "local" in journey_type:
        return "local"
    if "cont" in journey_type:
        return "return"
    if "return" in journey_type:
        return "return"
    return "single"


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


def update_exception_values_in_db(df: pd.DataFrame, plaza_name: str):
    """
    Group potential_exception_value by date, sum them, and update exception_value
    in the submissions table for each date.

    Rules:
    - Match on entity_name = plaza_name and date.
    - If exception_value already has a value in ANY row for that date, skip.
    - Otherwise, update only ONE row (the first by id).

    Args:
        df: DataFrame with 'date' and 'potential_exception_value' columns
        plaza_name: Plaza name to match against entity_name in DB
    """
    print(f"\n{'='*80}")
    print("STEP 2: Updating 'exception_value' in submissions table")
    print(f"{'='*80}")
    print(f"Plaza Name: {plaza_name}")

    # Validate required columns
    if 'date' not in df.columns or 'potential_exception_value' not in df.columns:
        print("ERROR: 'date' and/or 'potential_exception_value' columns not found.")
        print(f"Available columns: {list(df.columns)}")
        return

    # Group by date and sum potential_exception_value
    # Only consider non-null values
    df_with_values = df[df['potential_exception_value'].notna()].copy()
    
    if df_with_values.empty:
        print("No rows with potential_exception_value to process.")
        return

    # Convert date to string format (yyyy-mm-dd) if it's a datetime object
    if pd.api.types.is_datetime64_any_dtype(df_with_values['date']):
        df_with_values['date'] = df_with_values['date'].dt.strftime('%Y-%m-%d')
    else:
        # Convert to string and try to parse if needed
        df_with_values['date'] = df_with_values['date'].astype(str).str.strip()

    # Group by date and sum
    grouped = df_with_values.groupby('date')['potential_exception_value'].sum().reset_index(name='total_exception_value')
    
    print(f"\nFound {len(grouped)} unique dates with exception values:")
    print("-" * 50)
    for _, row in grouped.iterrows():
        print(f"  date={row['date']}, total_exception_value={row['total_exception_value']}")
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
            total_value = float(row['total_exception_value'])

            print(f"\n  Processing: date={date_val}, exception_value={total_value}")

            try:
                # Check if any row already has 'exception_value' set for this combination
                cursor.execute("""
                    SELECT id, exception_value
                    FROM public.submissions
                    WHERE entity_name = %s
                    AND date = %s
                    AND exception_value IS NOT NULL
                    AND CAST(exception_value AS TEXT) != ''
                    ORDER BY id
                    LIMIT 1
                """, (plaza_name, date_val))

                existing_row = cursor.fetchone()

                if existing_row:
                    print(f"    SKIPPED: 'exception_value' already has value '{existing_row[1]}' "
                          f"in row id={existing_row[0]}")
                    stats['skipped_already_exists'] += 1
                    continue

                # Find the first row (by id) for this combination to update
                cursor.execute("""
                    SELECT id
                    FROM public.submissions
                    WHERE entity_name = %s
                    AND date = %s
                    ORDER BY id
                    LIMIT 1
                """, (plaza_name, date_val))

                target_row = cursor.fetchone()

                if not target_row:
                    print(f"    SKIPPED: No rows found in DB for entity_name='{plaza_name}', "
                          f"date='{date_val}'")
                    stats['skipped_no_rows'] += 1
                    continue

                target_id = target_row[0]

                # Update only this one row
                cursor.execute("""
                    UPDATE public.submissions
                    SET exception_value = %s
                    WHERE id = %s
                """, (str(total_value), target_id))

                conn.commit()
                print(f"    UPDATED: row id={target_id}, exception_value set to {total_value}")
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
    print("UPDATE SUMMARY")
    print(f"{'='*80}")
    print(f"Total dates processed: {len(grouped)}")
    print(f"Updated:                {stats['updated']}")
    print(f"Skipped (already set):  {stats['skipped_already_exists']}")
    print(f"Skipped (no DB rows):    {stats['skipped_no_rows']}")
    print(f"Errors:                 {stats['errors']}")
    print(f"{'='*80}")


def clear_exception_values(plaza_name: str, from_date: str, to_date: str):
    """
    Standalone cleanup function: Set 'exception_value' to NULL for all rows
    matching the given plaza_name and date range.

    Args:
        plaza_name: Plaza name to match against entity_name in DB
        from_date: Start date in yyyy-mm-dd format (inclusive)
        to_date: End date in yyyy-mm-dd format (inclusive)
    """
    print(f"\n{'='*80}")
    print("CLEARING 'exception_value' VALUES")
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
            AND exception_value IS NOT NULL
        """, (plaza_name, from_date, to_date))

        affected_count = cursor.fetchone()[0]
        print(f"Rows with 'exception_value' set: {affected_count}")

        if affected_count == 0:
            print("No rows to clear. Nothing to do.")
            return

        # Confirm and clear
        cursor.execute("""
            UPDATE public.submissions
            SET exception_value = NULL
            WHERE entity_name = %s
            AND date >= %s
            AND date <= %s
            AND exception_value IS NOT NULL
        """, (plaza_name, from_date, to_date))

        conn.commit()
        print(f"Cleared 'exception_value' for {affected_count} row(s).")

    except Exception as e:
        conn.rollback()
        print(f"Error: {e}")
    finally:
        conn.close()

    print(f"{'='*80}")


def calculate_exception_values(excel_file_path: str, plaza_name: str):
    """
    Read an Excel file, calculate potential_exception_value for each row,
    and save the updated file.

    Args:
        excel_file_path: Path to the input Excel file
        plaza_name: Plaza name to look up rates from PLAZA_RATES
    """
    print(f"\n{'='*80}")
    print("Calculating Potential Exception Values")
    print(f"{'='*80}")
    print(f"Input File: {excel_file_path}")
    print(f"Plaza Name: {plaza_name}")

    # Validate plaza exists in rates
    if plaza_name not in PLAZA_RATES:
        print(f"ERROR: Plaza '{plaza_name}' not found in PLAZA_RATES.")
        print(f"Available plazas: {', '.join(PLAZA_RATES.keys())}")
        return

    rates = PLAZA_RATES[plaza_name]
    print(f"Rates loaded for '{plaza_name}':")
    for section, values in rates.items():
        print(f"  {section}: {values}")

    # Read Excel file
    df = pd.read_excel(excel_file_path, engine='openpyxl')
    total_rows = len(df)
    print(f"\nTotal rows: {total_rows}")

    # Validate required columns
    required_cols = ['Correct_index', 'TC_class_index']
    for col in required_cols:
        if col not in df.columns:
            print(f"ERROR: Required column '{col}' not found.")
            print(f"Available columns: {list(df.columns)}")
            return

    # Check for Journey Type / Trip Type column (optional - defaults to single if missing)
    journey_type_col = find_first_matching_column(df, JOURNEY_TYPE_COLUMN_CANDIDATES)
    has_journey_type = journey_type_col is not None
    if has_journey_type:
        print(f"Journey type column found: '{journey_type_col}'.")
    else:
        print("'Journey Type'/'Trip Type' column not found. Will use 'single' rate for all rows.")

    # Calculate potential_exception_value for each row
    results = []
    skipped = 0
    error_count = 0

    for idx, row in df.iterrows():
        try:
            # Get index values
            correct_idx = row['Correct_index']
            tc_idx = row['TC_class_index']

            # Skip if either index is empty/null
            if pd.isna(correct_idx) or pd.isna(tc_idx):
                results.append(None)
                skipped += 1
                continue

            # Convert to int
            correct_idx = int(float(correct_idx))
            tc_idx = int(float(tc_idx))

            # Determine rate type from Journey Type / Trip Type using keyword precedence:
            # local -> cont -> return -> single
            if has_journey_type:
                rate_section = get_rate_section_from_journey_type(row[journey_type_col])
            else:
                rate_section = "single"

            # Look up rate values
            correct_rate = rates.get(rate_section, {}).get(correct_idx)
            tc_rate = rates.get(rate_section, {}).get(tc_idx)

            if correct_rate is None:
                print(f"  Row {idx + 1}: No rate found for {rate_section}[{correct_idx}] (Correct_index). Skipping.")
                results.append(None)
                skipped += 1
                continue

            if tc_rate is None:
                print(f"  Row {idx + 1}: No rate found for {rate_section}[{tc_idx}] (TC_class_index). Skipping.")
                results.append(None)
                skipped += 1
                continue

            # Subtract: rate[Correct_index] - rate[TC_class_index]
            exception_value = correct_rate - tc_rate
            results.append(exception_value)

        except (ValueError, TypeError) as e:
            print(f"  Row {idx + 1}: Error processing - {e}")
            results.append(None)
            error_count += 1

    # Add results to DataFrame
    df['potential_exception_value'] = results

    # Print summary stats
    calculated = total_rows - skipped - error_count
    print(f"\n{'='*80}")
    print("SUMMARY")
    print(f"{'='*80}")
    print(f"Total rows:       {total_rows}")
    print(f"Calculated:       {calculated}")
    print(f"Skipped:          {skipped}")
    print(f"Errors:           {error_count}")

    # Show some sample results
    non_null_results = df[df['potential_exception_value'].notna()]
    if not non_null_results.empty:
        print(f"\nSample results (first 15 rows with values):")
        print("-" * 80)
        display_cols = ['Correct_index', 'TC_class_index']
        if has_journey_type:
            display_cols.append(journey_type_col)
        display_cols.append('potential_exception_value')
        print(non_null_results[display_cols].head(15).to_string(index=False))

    # Show distribution of exception values
    if not non_null_results.empty:
        print(f"\nException value distribution:")
        print(f"  Non-zero values: {(non_null_results['potential_exception_value'] != 0).sum()}")
        print(f"  Zero values:     {(non_null_results['potential_exception_value'] == 0).sum()}")
        print(f"  Positive values: {(non_null_results['potential_exception_value'] > 0).sum()}")
        print(f"  Negative values: {(non_null_results['potential_exception_value'] < 0).sum()}")

    # Save updated file
    df.to_excel(excel_file_path, index=False, engine='openpyxl')
    print(f"\nUpdated file saved: {excel_file_path}")
    print(f"{'='*80}")

    # Step 2: Update exception_value in DB grouped by date
    update_exception_values_in_db(df, plaza_name)


if __name__ == "__main__":
    # ============================================================================
    # CONFIGURATION
    # ============================================================================
    excel_file_path = "chalageri_combined_2025-11-23_to_2026-02-15.xlsx"
    plaza_name = "bahadrabad"
    # ============================================================================

    clear_exception_values(plaza_name, "2025-11-23", "2026-05-25")

    # calculate_exception_values(excel_file_path, plaza_name)

    

