"""
Clear exceptions and exception_value for a plaza/date range, then repopulate
them from a combined Excel report.

Usage:
    python refresh_exceptions_from_file.py
    (Configure excel_file_path, plaza_name, from_date, and to_date below)
"""

from update_exceptions import (
    clear_exceptions,
    remove_duplicates,
    update_exceptions_in_db,
)
from exception_value_count import clear_exception_values, calculate_exception_values


def refresh_exceptions_from_file(
    excel_file_path: str,
    plaza_name: str,
    from_date: str,
    to_date: str,
):
    """
    Clear exceptions and exception_value in the DB for the date range, then
    update both from the given Excel file.

    Args:
        excel_file_path: Path to the combined Excel report
        plaza_name: Plaza name matching entity_name in the DB
        from_date: Start date in yyyy-mm-dd format (inclusive)
        to_date: End date in yyyy-mm-dd format (inclusive)
    """
    print(f"\n{'=' * 80}")
    print("REFRESH EXCEPTIONS AND EXCEPTION VALUES FROM FILE")
    print(f"{'=' * 80}")
    print(f"Excel File:  {excel_file_path}")
    print(f"Plaza Name:  {plaza_name}")
    print(f"Date Range:  {from_date} to {to_date}")
    print(f"{'=' * 80}")

    print("\nStep 1: Clearing existing 'exceptions' values in DB")
    clear_exceptions(plaza_name, from_date, to_date)

    print("\nStep 2: Clearing existing 'exception_value' values in DB")
    clear_exception_values(plaza_name, from_date, to_date)

    print("\nStep 3: Removing duplicate rows from Excel (by 'CCH TXN NO')")
    df = remove_duplicates(excel_file_path)

    print("\nStep 4: Updating 'exceptions' in DB from Excel")
    update_exceptions_in_db(df, plaza_name)

    print("\nStep 5: Calculating and updating 'exception_value' in DB from Excel")
    calculate_exception_values(excel_file_path, plaza_name)

    print(f"\n{'=' * 80}")
    print("REFRESH COMPLETE")
    print(f"{'=' * 80}")


if __name__ == "__main__":
    # ============================================================================
    # CONFIGURATION
    # ============================================================================
    excel_file_path = "marripalem_combined_2026-05-26_to_2026-07-06.xlsx"
    plaza_name = "marripalem"
    from_date = "2026-05-26"
    to_date = "2026-07-06"
    # ============================================================================

    refresh_exceptions_from_file(
        excel_file_path,
        plaza_name,
        from_date,
        to_date,
    )
