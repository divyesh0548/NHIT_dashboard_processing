"""
Data Extraction from Database

This module provides a function to extract data from the submissions table
in the snt_form database based on date range and file_status filter.
"""

import psycopg2
import pandas as pd
from datetime import datetime

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


def extract_submissions_data(from_date, to_date, output_file=None):
    """
    Extract data from submissions table based on date range and file_status.
    
    Args:
        from_date (str): Start date in format "YYYY-MM-DD" (e.g., "2026-01-16")
        to_date (str): End date in format "YYYY-MM-DD" (e.g., "2026-01-20")
        output_file (str, optional): Path to save the extracted data as Excel file.
                                     If provided, data will be saved to this file.
                                     Example: "extracted_data.xlsx"
    
    Returns:
        pandas.DataFrame: DataFrame containing the extracted data with columns:
            - entity_name
            - name
            - email
            - date
            - shift
            - exception_value
            - total_transactions
            - revenue
            - violation_amt
        Returns None if connection fails or query execution fails.
    
    Example:
        >>> df = extract_submissions_data("2026-01-16", "2026-01-20", "output.xlsx")
        >>> print(df.head())
    """
    # Validate date format
    try:
        datetime.strptime(from_date, "%Y-%m-%d")
        datetime.strptime(to_date, "%Y-%m-%d")
    except ValueError as e:
        print(f"Error: Invalid date format. Expected format: YYYY-MM-DD. Error: {e}")
        return None
    
    # Connect to database
    conn = get_db_connection()
    if not conn:
        print("Failed to connect to database")
        return None
    
    # Columns to extract
    columns = [
        'entity_name',
        'name',
        'email',
        'date',
        'shift',
        'exception_value',
        'total_transactions',
        'revenue',
        'violation_amt'
    ]
    
    # Build SQL query
    columns_str = ', '.join(columns)
    query = f"""
        SELECT {columns_str}
        FROM submissions
        WHERE file_status = 1
        AND date >= %s
        AND date <= %s
        ORDER BY date, entity_name
    """
    
    try:
        # Execute query and fetch results
        df = pd.read_sql_query(query, conn, params=(from_date, to_date))
        
        print(f"Successfully extracted {len(df)} rows from submissions table")
        print(f"Date range: {from_date} to {to_date}")
        
        # Save to Excel file if output_file is provided
        if output_file:
            try:
                df.to_excel(output_file, index=False, engine='openpyxl')
                print(f"Data saved to Excel file: {output_file}")
            except Exception as e:
                print(f"Warning: Failed to save to Excel file: {e}")
                print("Make sure openpyxl is installed: pip install openpyxl")
        
        return df
    
    except psycopg2.Error as e:
        print(f"Error executing query: {e}")
        return None
    
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None
    
    finally:
        # Close database connection
        if conn:
            conn.close()


if __name__ == "__main__":
    # Example usage
    from_date = "2026-01-10"
    to_date = "2026-01-20"
    output_excel = "extracted_submissions_data.xlsx"  # Output Excel file name
    
    print("=" * 60)
    print("Extracting data from submissions table")
    print("=" * 60)
    
    df = extract_submissions_data(from_date, to_date, output_file=output_excel)
    
    if df is not None and not df.empty:
        print(f"\nExtracted {len(df)} rows")
        print(f"Data saved to: {output_excel}")
        print("\nFirst few rows:")
        print(df.head())
    elif df is not None:
        print("\nNo data found for the specified date range")
    else:
        print("\nFailed to extract data")

