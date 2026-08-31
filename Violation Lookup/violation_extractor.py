
import psycopg2
import pandas as pd
import requests
import os
import tempfile
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List
import sys


ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from db_config import get_db_connection_kwargs


def normalize_column_name(col_name: str) -> str:
    """
    Normalize column name for matching by:
    - Converting to lowercase
    - Removing spaces and underscores
    - Removing special characters (keeping alphanumeric only)
    
    Args:
        col_name: Original column name
        
    Returns:
        Normalized column name for comparison
    """
    if pd.isna(col_name):
        return ""
    # Convert to string and lowercase
    normalized = str(col_name).lower().strip()
    # Remove spaces, underscores, and special characters (keep alphanumeric)
    normalized = re.sub(r'[^a-z0-9]', '', normalized)
    return normalized


def get_column_aliases() -> Dict[str, List[str]]:
    """
    Define alternative column names for each target column.
    Each target column can have multiple alternative names that refer to the same data.
    
    Returns:
        Dictionary mapping target column names to lists of alternative names
    """
    return {
        'TRANSACTION DATE': [
            'Transaction Date',
            'Transaction_Date',
            'DATE',
            'Date',
            'Date & Time',
            'TransactionDate',
            'TxnDate'
        ],
        'TRANSACTION NO.': [
            'Transaction No.',
            'Transaction_No.',
            'Transaction No',
            'Transaction_No',
            'TRANSACTION NO',
            'Transaction ID',
            'TRANSACTION NUMBER',
            'Transaction Number',
            'Transaction_Number',
            'TXN NUMBER',
            'Txn Number',
            'Txn_Number',
            'TransactionNo',
            'TxnNo'
        ],
        'MVC (TLC CLASS)': [
            'MVC (TLC Class)',
            'MVC(TLC CLASS)',
            'MVC(TLC Class)',
            'MVC_TLC_CLASS',
            'MVC TLC CLASS',
            'MVC TLC Class',
            'MVC_TLC_Class',
            'TC Class',
            'TLC CLASS',
            'TLC Class',
            'TLC_CLASS',
            'TLC_Class',
            'MVC',
            'Mvc',
            'TLC',
            'Tlc'
        ],
        'Veh Reg No.': [
            'Veh Reg No',
            'Veh_Reg_No.',
            'Veh_Reg_No',
            'VEH REG NO.',
            'Veh Reg No.',
            'Vehicle Registration No.',
            'Vehicle Registration No',
            'Vehicle_Registration_No.',
            'Vehicle_Registration_No',
            'VRN',
            'Vrn',
            'VEHICLE REG NO.',
            'Vehicle Reg No.',
            'Vehicle_Reg_No.',
            'Vehicle Reg No',
            'Vehicle_Reg_No',
            'VehRegNo',
            'VehicleRegNo'
        ],
        'Weight': [
            'WEIGHT',
            'weight',
            'WEIGHT (KG)',
            'Weight (kg)',
            'Weight (KG)',
            'Weight_kg',
            'Weight_KG',
            'WEIGHT_KG',
            'WEIGHT_KG',
            'WeightKg',
            'WeightKG'
        ],
        'Weight_Based_Class': [
            'Weight Based Class',
            'Weight_Based_Class',
            'WeightBasedClass',
            'WEIGHT BASED CLASS',
            'Weight Based Class',
            'Weight_Based_Class',
            'WBC',
            'Wbc'
        ],
        'class_accuracy': [
            'Class Accuracy',
            'Class_Accuracy',
            'ClassAccuracy',
            'CLASS ACCURACY',
            'class accuracy',
            'Accuracy',
            'accuracy',
            'ACCURACY'
        ],
        'TC_class_index': [
            'TC Class Index',
            'TC_Class_Index',
            'TCClassIndex',
            'TC CLASS INDEX',
            'TC class index',
            'TC_Class_Index',
            'TC Index',
            'TC_Index',
            'TCIndex'
        ],
        'Correct_index': [
            'Correct Index',
            'Correct_Index',
            'CorrectIndex',
            'CORRECT INDEX',
            'correct index',
            'Correct Index',
            'Correct_Index',
            'CI',
            'Ci'
        ],
        'result': [
            'Result',
            'RESULT',
            'Result Status',
            'Result_Status',
            'ResultStatus',
            'RESULT STATUS',
            'Status',
            'status',
            'STATUS'
        ],
        'CCH TXN NO': [
            'CCH TXN NO',
            'ICD TX ID',
            'CCH TRANSACTION ID',
            'CCH TXN NO.',
            'CCH_TRANSACTION_ID',
        ]
    }


def find_matching_column(target_col: str, available_columns: List[str], 
                        column_aliases: Optional[Dict[str, List[str]]] = None) -> Optional[str]:
    """
    Find a matching column from available columns that matches the target column.
    Handles variations in spaces, underscores, and case.
    Also tries alternative column names defined in column_aliases in sequential order.
    Stops at the first match found.
    
    Args:
        target_col: Target column name to find
        available_columns: List of available column names in the Excel file
        column_aliases: Optional dictionary mapping target columns to alternative names.
                       If not provided, will use get_column_aliases()
        
    Returns:
        Matching column name from available_columns, or None if not found
    """
    if column_aliases is None:
        column_aliases = get_column_aliases()
    
    # First, try exact match (normalized) of the target column itself
    target_normalized = normalize_column_name(target_col)
    
    for col in available_columns:
        if normalize_column_name(col) == target_normalized:
            return col
    
    # If not found, try alternative names in the exact sequence they are written
    # Stop immediately at the first match found
    if target_col in column_aliases:
        for alias in column_aliases[target_col]:
            alias_normalized = normalize_column_name(alias)
            # Check this alias against all available columns
            for col in available_columns:
                if normalize_column_name(col) == alias_normalized:
                    # Found a match - return immediately (don't check remaining aliases)
                    return col
    
    return None


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


def extract_and_combine_excel_files(entity_name: str, from_date: str = None, to_date: str = None, output_file: Optional[str] = None, specific_dates: list = None) -> Optional[str]:
    """
    Extract Excel files from submissions table based on entity_name and date range or specific dates,
    then combine them into a single output Excel file with only specific columns.
    
    Extracts only these columns from Excel files (handles variations in spaces, underscores, and case):
    - TRANSACTION DATE
    - TRANSACTION NO.
    - MVC (TLC CLASS)
    - Veh Reg No.
    - Weight
    - Weight_Based_Class
    - class_accuracy
    - TC_class_index
    - Correct_index
    - result
    - CCH TXN NO
    
    Additionally, 'date' and 'shift' values are taken from the submissions table (not from Excel files)
    and added to every row of the output.
    
    Args:
        entity_name (str): The entity name to filter by (matches entity_name column)
        from_date (str, optional): Start date in format "YYYY-MM-DD" (used for date range mode)
        to_date (str, optional): End date in format "YYYY-MM-DD" (used for date range mode)
        output_file (str, optional): Path to save the combined Excel file.
                                     If not provided, generates a filename based on entity_name and dates.
        specific_dates (list, optional): List of specific dates in "YYYY-MM-DD" format.
                                         If provided, only these exact dates are queried (not a range).
    
    Returns:
        str: Path to the output Excel file if successful, None otherwise.
    
    Example (date range):
        >>> output_path = extract_and_combine_excel_files(
        ...     "maigalganj", 
        ...     from_date="2025-12-23", 
        ...     to_date="2025-12-25",
        ...     output_file="combined_output.xlsx"
        ... )
    
    Example (specific dates):
        >>> output_path = extract_and_combine_excel_files(
        ...     "maigalganj",
        ...     specific_dates=["2025-12-05", "2025-12-10", "2026-01-15"],
        ...     output_file="combined_output.xlsx"
        ... )
    """
    # Define target columns to extract from Excel files
    target_columns = [
        'TRANSACTION DATE',
        'TRANSACTION NO.',
        'MVC (TLC CLASS)',
        'Veh Reg No.',
        'Weight',
        'Weight_Based_Class',
        'class_accuracy',
        'TC_class_index',
        'Correct_index',
        'result',
        'CCH TXN NO'
    ]
    
    # These columns come from the DB submissions table, not from Excel files
    db_columns = ['date', 'shift']
    
    # Final output column order: db columns first, then Excel columns
    output_columns = db_columns + target_columns
    # Validate inputs: either specific_dates or from_date+to_date must be provided
    if specific_dates:
        # Validate each specific date
        for d in specific_dates:
            try:
                datetime.strptime(d, "%Y-%m-%d")
            except ValueError as e:
                print(f"Error: Invalid date '{d}'. Expected format: YYYY-MM-DD. Error: {e}")
                return None
        date_display = f"specific dates: {specific_dates}"
    elif from_date and to_date:
        try:
            datetime.strptime(from_date, "%Y-%m-%d")
            datetime.strptime(to_date, "%Y-%m-%d")
        except ValueError as e:
            print(f"Error: Invalid date format. Expected format: YYYY-MM-DD. Error: {e}")
            return None
        date_display = f"{from_date} to {to_date}"
    else:
        print("Error: Either specific_dates or both from_date and to_date must be provided.")
        return None
    
    # Connect to database
    conn = get_db_connection()
    if not conn:
        print("Failed to connect to database")
        return None
    
    try:
        cursor = conn.cursor()
        
        if specific_dates:
            # Query for specific dates using IN clause
            placeholders = ','.join(['%s'] * len(specific_dates))
            query = f"""
                SELECT file_link, date, shift
                FROM submissions
                WHERE entity_name = %s
                AND date IN ({placeholders})
                AND file_status = 1
                AND file_link IS NOT NULL
                AND file_link != ''
                ORDER BY date, shift
            """
            cursor.execute(query, (entity_name, *specific_dates))
        else:
            # Query for date range
            query = """
                SELECT file_link, date, shift
                FROM submissions
                WHERE entity_name = %s
                AND date >= %s
                AND date <= %s
                AND file_status = 1
                AND file_link IS NOT NULL
                AND file_link != ''
                ORDER BY date, shift
            """
            cursor.execute(query, (entity_name, from_date, to_date))
        
        rows = cursor.fetchall()
        
        if not rows:
            print(f"No records found for entity_name='{entity_name}' for {date_display}")
            return None
        
        # Each row is (file_link, date, shift)
        file_entries = [(row[0], row[1], row[2]) for row in rows]
        print(f"Found {len(file_entries)} file(s) to download and combine")
        print(f"Target columns to extract from Excel: {', '.join(target_columns)}")
        print(f"Columns from DB: {', '.join(db_columns)}")
        
        # List to store all dataframes
        all_dataframes = []
        temp_files = []
        column_mapping = {}  # Maps target column to actual column name found in first file
        
        # Download and read each Excel file
        for idx, (file_link, db_date, db_shift) in enumerate(file_entries, 1):
            try:
                print(f"[{idx}/{len(file_entries)}] Downloading: {file_link}")
                print(f"  DB date: {db_date}, DB shift: {db_shift}")
                
                # Download file
                response = requests.get(file_link, timeout=60)
                response.raise_for_status()
                
                # Save to temporary file
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                temp_file.write(response.content)
                temp_file.close()
                temp_files.append(temp_file.name)
                
                # Read Excel file (pandas automatically uses first row as headers)
                print(f"  Reading Excel file...")
                df = pd.read_excel(temp_file.name, engine='openpyxl')
                
                if df.empty:
                    print(f"  Warning: File is empty, skipping...")
                    continue
                
                available_columns = list(df.columns)
                
                # Get column aliases for alternative name matching
                column_aliases = get_column_aliases()
                
                # Find matching columns for this file (handles variations in column names)
                file_column_mapping = {}
                for target_col in target_columns:
                    matched_col = find_matching_column(target_col, available_columns, column_aliases)
                    if matched_col:
                        file_column_mapping[target_col] = matched_col
                    else:
                        file_column_mapping[target_col] = None
                
                # First file: print mapping details
                if not column_mapping:
                    print(f"  Available columns in file: {len(available_columns)}")
                    print(f"  Available column names: {', '.join(available_columns[:10])}{'...' if len(available_columns) > 10 else ''}")
                    for target_col in target_columns:
                        matched_col = file_column_mapping.get(target_col)
                        if matched_col:
                            # Check if it matched via alias
                            aliases = column_aliases.get(target_col, [])
                            if matched_col != target_col and matched_col not in aliases:
                                # Normalized match (different formatting)
                                print(f"    Found: '{target_col}' -> '{matched_col}' (normalized match)")
                            elif matched_col in aliases:
                                # Matched via alias
                                print(f"    Found: '{target_col}' -> '{matched_col}' (via alias)")
                            else:
                                # Exact match
                                print(f"    Found: '{target_col}' -> '{matched_col}' (exact match)")
                        else:
                            # Show available alternatives for debugging
                            aliases = column_aliases.get(target_col, [])
                            if aliases:
                                print(f"    Missing: '{target_col}' (tried aliases: {', '.join(aliases[:3])}{'...' if len(aliases) > 3 else ''})")
                            else:
                                print(f"    Missing: '{target_col}' (will be added as empty)")
                    
                    found_count = sum(1 for v in file_column_mapping.values() if v is not None)
                    print(f"  Matched {found_count}/{len(target_columns)} target columns")
                    column_mapping = file_column_mapping.copy()  # Store for reference
                
                # Extract only the target columns using this file's mapping
                extracted_data = {}
                for target_col in target_columns:
                    actual_col = file_column_mapping.get(target_col)
                    if actual_col and actual_col in df.columns:
                        # Always use target_col as the column name in output (e.g., 'CCH TXN NO')
                        # even if an alias was found (e.g., 'ICD TX ID')
                        extracted_data[target_col] = df[actual_col].values
                    else:
                        # Add as empty column if not found
                        extracted_data[target_col] = [None] * len(df)
                
                # Create dataframe with only target columns in the correct order
                # Column names will be exactly as specified in target_columns (e.g., 'CCH TXN NO')
                df_extracted = pd.DataFrame(extracted_data)
                
                # Check if first data row looks like a header row (duplicate header in data)
                if len(df_extracted) > 0:
                    first_row_values = [str(val).strip().upper() if pd.notna(val) else '' for val in df_extracted.iloc[0].values]
                    target_columns_normalized = [normalize_column_name(col) for col in target_columns]
                    first_row_normalized = [normalize_column_name(str(val)) for val in df_extracted.iloc[0].values]
                    
                    # Check if first row matches header structure
                    matches = sum(1 for i, val in enumerate(first_row_normalized) 
                                if i < len(target_columns_normalized) and val == target_columns_normalized[i])
                    match_ratio = matches / len(target_columns_normalized) if len(target_columns_normalized) > 0 else 0
                    
                    # If first row looks like a header (high match ratio), remove it
                    if match_ratio >= 0.5:
                        print(f"  Removing duplicate header row from data (match ratio: {match_ratio:.2%})...")
                        df_extracted = df_extracted.iloc[1:].reset_index(drop=True)
                
                # Only keep rows that have data (not all NaN)
                df_extracted = df_extracted.dropna(how='all')
                
                if not df_extracted.empty:
                    # Add date and shift from DB submissions table to all rows
                    df_extracted.insert(0, 'date', str(db_date))
                    df_extracted.insert(1, 'shift', str(db_shift))
                    
                    all_dataframes.append(df_extracted)
                    print(f"  Added {len(df_extracted)} rows with {len(output_columns)} columns (incl. DB date & shift)")
                else:
                    print(f"  Warning: No data rows found after processing, skipping...")
                    
            except requests.exceptions.RequestException as e:
                print(f"  Error downloading file: {e}")
                continue
            except Exception as e:
                print(f"  Error processing file: {e}")
                continue
        
        # Clean up temporary files
        for temp_file in temp_files:
            try:
                os.remove(temp_file)
            except:
                pass
        
        if not all_dataframes:
            print("No data to combine. All files were empty or failed to process.")
            return None
        
        # Combine all dataframes
        print(f"\nCombining {len(all_dataframes)} file(s)...")
        
        # Ensure all dataframes have the same column structure (output_columns)
        for i, df in enumerate(all_dataframes):
            # Ensure all output columns exist
            for col in output_columns:
                if col not in df.columns:
                    df[col] = None
            # Reorder to match output columns order
            all_dataframes[i] = df[output_columns]
        
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        
        # Final check: ensure output has exact column structure (output_columns)
        # Rename any columns that might have different names to match output_columns
        # This ensures 'CCH TXN NO' is always used even if an alias was found
        column_rename_map = {}
        for target_col in output_columns:
            # Check if target_col exists (case-insensitive check)
            found_col = None
            for col in combined_df.columns:
                if normalize_column_name(col) == normalize_column_name(target_col):
                    if col != target_col:
                        found_col = col
                        break
            
            if found_col:
                column_rename_map[found_col] = target_col
        
        if column_rename_map:
            combined_df = combined_df.rename(columns=column_rename_map)
        
        # Ensure columns are in the correct order
        if list(combined_df.columns) != output_columns:
            combined_df = combined_df[output_columns]
        
        print(f"Total rows in combined file: {len(combined_df)}")
        print(f"Output structure: {len(combined_df.columns)} columns ({', '.join(output_columns)})")
        
        # Generate output filename if not provided
        if not output_file:
            safe_entity = entity_name.replace(' ', '_').replace('/', '_')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if specific_dates:
                output_file = f"{safe_entity}_combined_remaining_dates_{timestamp}.xlsx"
            else:
                output_file = f"{safe_entity}_combined_{from_date}_to_{to_date}_{timestamp}.xlsx"
        
        # Save to Excel file
        try:
            combined_df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"\nSuccessfully saved combined file to: {output_file}")
            return output_file
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return None
        
    except psycopg2.Error as e:
        print(f"Database error: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        # Close database connection
        if conn:
            conn.close()


if __name__ == "__main__":
    # List of entity names to process
    entity_names = [
        # "aroli",
        # "bassi",
        # "dhaneshwar",
        # "hattargi",
        # "khemana",
        # "kognoli",
        # "pullur",
        # "undavariya",
        # "daroda",
        # "kelapur",
        # "raibha",
        # "raksha",
        # "balibhasa",
        # "bankapur",
        # "boharipar",
        # "chalageri",
        # "dahalapara",
        # "hebbalu",
        # "kherwasani",
        # "madai",
        # "mohtara",
        # "usaka",
        # "galia",
        # "khawasa",
        # "odhaki_paipkhar",
        # "patgaon",
        # "bahadrabad",
        # "bhojpuree",
        # "chhapar",
        # "dukkavanipalem",
        # # "madapam",
        # "mahasamudram",
        # # "maigalganj",
        # "mudhipar",
        # "tarapongi",
        # "veeravalli",
        # # "faridpur",
        "gadanki",
        # "marripalem",
        # "nathavalasa",
    ]

    
    # ============================================================================
    # DATE CONFIGURATION
    # ============================================================================
    # OPTION 1: Use a date range (uncomment these three lines and comment out OPTION 2)
    from_date = "2025-11-23"
    to_date = "2026-05-25"
    specific_dates = None
    
    # OPTION 2: Use specific dates only (uncomment this line and comment out OPTION 1)
    # specific_dates = [
    #     "2025-11-23", "2025-11-24", "2025-11-25", "2025-11-26", "2025-11-27", "2025-11-28", "2025-11-29", "2025-11-30",
    #     "2026-02-01", "2026-02-02", "2026-02-03", "2026-02-04", "2026-02-05", "2026-02-06", "2026-02-07", "2026-02-08", "2026-02-09",
    # ]
    # from_date = None
    # to_date = None
    # ============================================================================
    
    print("\n" + "=" * 80)
    print("Extracting and Combining Excel Files for Multiple Entities")
    print("=" * 80)
    if specific_dates:
        print(f"Specific Dates: {specific_dates}")
    else:
        print(f"Date Range: {from_date} to {to_date}")
    print(f"Total Entities to Process: {len(entity_names)}")
    print("=" * 80)
    
    # Track results
    success_count = 0
    failed_count = 0
    failed_entities = []
    
    # Process each entity
    for idx, entity_name in enumerate(entity_names, 1):
        print(f"\n{'='*80}")
        print(f"Processing Entity {idx}/{len(entity_names)}: {entity_name}")
        print(f"{'='*80}")
        
        # Generate output file name
        if specific_dates:
            output_file = f"{entity_name}_combined_remaining_dates.xlsx"
        else:
            output_file = f"{entity_name}_combined_{from_date}_to_{to_date}.xlsx"
        
        print(f"Entity Name: {entity_name}")
        print(f"Output File: {output_file}")
        print("-" * 80)
        
        # Extract and combine files (function handles both date range and specific dates)
        output_path = extract_and_combine_excel_files(
            entity_name,
            from_date=from_date,
            to_date=to_date,
            output_file=output_file,
            specific_dates=specific_dates
        )
        
        if output_path:
            print(f"\n✓ Success! Output file: {output_path}")
            success_count += 1
        else:
            print(f"\n✗ Failed to extract and combine files for {entity_name}")
            failed_count += 1
            failed_entities.append(entity_name)
    
    # Print summary
    print(f"\n{'='*80}")
    print("PROCESSING SUMMARY")
    print(f"{'='*80}")
    print(f"Total Entities: {len(entity_names)}")
    print(f"Successful: {success_count}")
    print(f"Failed: {failed_count}")
    
    if failed_entities:
        print(f"\nFailed Entities:")
        for entity in failed_entities:
            print(f"  - {entity}")
    
    print(f"{'='*80}")

