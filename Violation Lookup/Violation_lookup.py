"""
Violation Lookup Program
Takes a folder and an Excel file as input.
1. Processes single Excel file with 'CCH TXN NO' column (trailing spaces removed)
2. Processes all Excel files in folder with 'Agency Txn Id' column
3. Compares CCH TXN NO from single file with Agency Txn Id from folder files to find matches
   (special characters removed before comparison; only letters and digits are kept)
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import sys
import glob


# Standard output column names written to the combined report
STANDARD_AGENCY_TXN_ID_COL = "Agency Txn Id"
STANDARD_VIOLATION_AMT_COL = "Violation Amt"
STANDARD_READER_READ_TIME_COL = "Reader Read Time"
STANDARD_JOURNEY_TYPE_COL = "Journey Type"
AGENCY_TXN_ID_COLUMN_NAMES = [
    'Agency Txn Id',
    'Agency Txn Id ',
    'Agency Txn Id\t',
    'Agency Txn Id.',
    'Agency Txn Id. ',
    'Agency Txn Id.\t',
    'Agency_Txn_Id',
    'Agency_Txn_Id ',
    'Agency_Txn_Id\t',
    'Agency_Txn_Id.',
    'Agency_Txn_Id. ',
    'Agency_Txn_Id.\t',
    'AGENCY TXN ID',
    'AGENCY TXN ID ',
    'AGENCY TXN ID\t',
    'AGENCY TXN ID.',
    'AGENCY TXN ID. ',
    'AGENCY TXN ID.\t',
    'Agency Transaction Id',
    'Agency Transaction Id ',
    'Agency Transaction Id\t',
    'Agency Transaction Id.',
    'Agency Transaction Id. ',
    'Agency Transaction Id.\t',
    'Agency_Transaction_Id',
    'Agency_Transaction_Id ',
    'Agency_Transaction_Id\t',
    'Agency_Transaction_Id.',
    'Agency_Transaction_Id. ',
    'Agency_Transaction_Id.\t',
    'Agency Transaction ID',
    'Agency Transaction ID ',
    'Agency Transaction ID\t',
    'Agency Tran Id',
    'Agency Tran Id ',
    'Agency Tran Id\t',
    'Agency_Tran_Id',
    'Agency_Tran_Id ',
    'Agency_Tran_Id\t',
    'AGENCY TRAN ID',
    'AGENCY TRAN ID ',
    'AGENCY TRAN ID\t',
    'Agency_Transaction_ID',
    'Agency_Transaction_ID ',
    'Agency_Transaction_ID\t',
    'AgencyTxnId',
    'AgencyTxnId ',
    'AgencyTxnId\t',
    'AgencyTxnID',
    'AgencyTxnID ',
    'AgencyTxnID\t',
    'Agency TxnID',
    'Agency TxnID ',
    'Agency TxnID\t',
    'Agency_TxnID',
    'Agency_TxnID ',
    'Agency_TxnID\t'
]
CCH_TXN_ID_COLUMN_NAMES = [
    'CCH Tran Id',
    'CCH Tran Id ',
    'CCH Tran Id\t',
    'CCH Tran Id.',
    'CCH Tran Id. ',
    'CCH Tran Id.\t',
    'CCH_Txn_Id',
    'CCH_Txn_Id ',
    'CCH_Txn_Id\t',
    'CCH_Txn_Id.',
    'CCH_Txn_Id. ',
    'CCH_Txn_Id.\t',
    'CCH TXN ID',
    'CCH TXN ID ',
    'CCH TXN ID\t',
    'CCH TXN ID.',
    'CCH TXN ID. ',
    'CCH TXN ID.\t',
    'CCH Transaction Id',
    'CCH Transaction Id ',
    'CCH Transaction Id\t',
    'CCH Transaction Id.',
    'CCH Transaction Id. ',
    'CCH Transaction Id.\t',
    'CCH_Transaction_Id',
    'CCH_Transaction_Id ',
    'CCH_Transaction_Id\t',
    'CCH_Transaction_Id.',
    'CCH_Transaction_Id. ',
    'CCH_Transaction_Id.\t',
    'CCH Transaction ID',
    'CCH Transaction ID ',
    'CCH Transaction ID\t',
    'CCH Tran ID',
    'CCH Tran ID ',
    'CCH Tran ID\t',
    'CCH_Tran_ID',
    'CCH_Tran_ID ',
    'CCH_Tran_ID\t',
    'CCHTxnId',
    'CCHTxnId ',
    'CCHTxnId\t',
    'CCHTxnID',
    'CCHTxnID ',
    'CCHTxnID\t',
]
JOURNEY_TYPE_COLUMN_NAMES = [
    'Journey Type',
    'Journey Type ',
    'Journey Type\t',
    'JourneyType',
    'JourneyType ',
    'JourneyType\t',
    'Journey_Type',
    'Journey_Type ',
    'Journey_Type\t',
    'JOURNEY TYPE',
    'JOURNEY TYPE ',
    'JOURNEY TYPE\t',
    'Journey Type.',
    'Journey Type. ',
    'Journey Type.\t',
    'Trip Type',
    'Trip Type ',
    'Trip Type\t',
    'TripType',
    'TripType ',
    'TripType\t',
    'Trip_Type',
    'Trip_Type ',
    'Trip_Type\t',
    'TRIP TYPE',
    'TRIP TYPE ',
    'TRIP TYPE\t',
    'Trip Type.',
    'Trip Type. ',
    'Trip Type.\t',
    'Fare Type',
    'Fare Type ',
    'Fare Type\t',
    'FareType',
    'FareType ',
    'FareType\t',
    'Fare_Type',
    'Fare_Type ',
    'Fare_Type\t',
    'FARE TYPE',
    'FARE TYPE ',
    'FARE TYPE\t',
    'Fare Type.',
    'Fare Type. ',
    'Fare Type.\t',
]


def get_date_format_from_user():
    """
    Ask user to select the date format used in the Excel file.
    
    Returns:
        str: 'dd/mm/yyyy' or 'mm/dd/yyyy'
    """
    print("\n" + "="*80)
    print("DATE FORMAT SELECTION")
    print("="*80)
    print("Please select the date format used in your Excel file:")
    print("  1. dd/mm/yyyy (e.g., 02/12/2025 means 2nd December 2025)")
    print("  2. mm/dd/yyyy (e.g., 12/2/2025 means 12th February 2025)")
    print("="*80)
    
    while True:
        choice = input("Enter your choice (1 or 2): ").strip()
        if choice == '1':
            print("Selected format: dd/mm/yyyy")
            return 'dd/mm/yyyy'
        elif choice == '2':
            print("Selected format: mm/dd/yyyy")
            return 'mm/dd/yyyy'
        else:
            print("Invalid choice. Please enter 1 or 2.")


def extract_ddmmyyyy(date_value, date_format='dd/mm/yyyy'):
    """
    Extract date in ddmmyyyy format from various datetime formats.
    Used for single Excel file processing (TRANSACTION DATE column).
    
    Handles formats like:
    - '02-12-2025 03:11:11' --> 02122025
    - '02/12/2025 05:20:39' --> 02122025 (if date_format='dd/mm/yyyy')
    - '12/2/2025 11:11:34' --> 02122025 (if date_format='mm/dd/yyyy' - month/day/year format)
    - '08-Dec-2025 02:07:44 PM' --> 08122025
    
    Args:
        date_value: Date value from Excel (can be string, datetime, or pandas Timestamp)
        date_format: 'dd/mm/yyyy' or 'mm/dd/yyyy' - format used in Excel file
    
    Returns:
        str: Date in ddmmyyyy format, or None if parsing fails
    """
    try:
        # Handle None, NaN, or empty values
        if pd.isna(date_value) or date_value is None:
            return None
        
        # If it's already a datetime object or pandas Timestamp, extract directly
        if isinstance(date_value, (datetime, pd.Timestamp)):
            day = date_value.day
            month = date_value.month
            year = date_value.year
            return f"{day:02d}{month:02d}{year}"
        
        # Convert to string for parsing
        date_str = str(date_value).strip()
        
        if not date_str or date_str.lower() in ['nan', 'none', '']:
            return None
        
        # Try parsing different formats
        parsed_date = None
        
        # Format 1: '02-12-2025 03:11:11' (dd-mm-yyyy with time)
        try:
            parsed_date = datetime.strptime(date_str.split()[0], "%d-%m-%Y")
        except (ValueError, IndexError):
            pass
        
        # Format 2: Handle slash-separated dates based on user-selected format
        if not parsed_date:
            try:
                date_part = date_str.split()[0]
                if date_format == 'dd/mm/yyyy':
                    # Try dd/mm/yyyy format first
                    parsed_date = datetime.strptime(date_part, "%d/%m/%Y")
                else:
                    # Try mm/dd/yyyy format first
                    parsed_date = datetime.strptime(date_part, "%m/%d/%Y")
            except (ValueError, IndexError):
                # If the selected format fails, try the other format as fallback
                if date_format == 'dd/mm/yyyy':
                    try:
                        parsed_date = datetime.strptime(date_part, "%m/%d/%Y")
                    except (ValueError, IndexError):
                        pass
                else:
                    try:
                        parsed_date = datetime.strptime(date_part, "%d/%m/%Y")
                    except (ValueError, IndexError):
                        pass
        
        # Format 3: '08-Dec-2025 02:07:44 PM' (dd-Mon-yyyy with time)
        if not parsed_date:
            try:
                date_part = date_str.split()[0]
                parsed_date = datetime.strptime(date_part, "%d-%b-%Y")
            except (ValueError, IndexError):
                pass
        
        if not parsed_date:
            try:
                date_part = date_str.split()[0]
                parsed_date = datetime.strptime(date_part, "%Y-%m-%d")
            except (ValueError, IndexError):
                pass
        
        if not parsed_date:
            try:
                parsed_date = pd.to_datetime(date_str)
                if isinstance(parsed_date, pd.Timestamp):
                    parsed_date = parsed_date.to_pydatetime()
            except:
                pass
        
        if parsed_date:
            day = parsed_date.day
            month = parsed_date.month
            year = parsed_date.year
            return f"{day:02d}{month:02d}{year}"
        else:
            return None
            
    except Exception as e:
        return None


def extract_ddmmyyyy_from_yyyymmdd(date_value):
    """
    Extract date in ddmmyyyy format from yyyy-mm-dd HH:MM:SS format.
    
    Handles format like:
    - '2025-12-01 00:01:24' --> 01122025
    
    Args:
        date_value: Date value from Excel (can be string, datetime, or pandas Timestamp)
    
    Returns:
        str: Date in ddmmyyyy format, or None if parsing fails
    """
    try:
        # Handle None, NaN, or empty values
        if pd.isna(date_value) or date_value is None:
            return None
        
        # If it's already a datetime object or pandas Timestamp, extract directly
        if isinstance(date_value, (datetime, pd.Timestamp)):
            day = date_value.day
            month = date_value.month
            year = date_value.year
            return f"{day:02d}{month:02d}{year}"
        
        # Convert to string for parsing
        date_str = str(date_value).strip()
        
        if not date_str or date_str.lower() in ['nan', 'none', '']:
            return None
        
        # Try parsing yyyy-mm-dd format first (primary format for Reader Read Time)
        parsed_date = None
        
        # Format: '2025-12-01 00:01:24' (yyyy-mm-dd with time)
        try:
            # Extract date part (before space)
            date_part = date_str.split()[0]
            # Parse yyyy-mm-dd format
            parsed_date = datetime.strptime(date_part, "%Y-%m-%d")
        except (ValueError, IndexError):
            # Try parsing the full string with datetime
            try:
                parsed_date = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                pass
        
        # If still not parsed, try pandas to_datetime as fallback
        if not parsed_date:
            try:
                parsed_date = pd.to_datetime(date_str)
                if isinstance(parsed_date, pd.Timestamp):
                    parsed_date = parsed_date.to_pydatetime()
            except:
                pass
        
        if parsed_date:
            # Format as ddmmyyyy
            day = parsed_date.day
            month = parsed_date.month
            year = parsed_date.year
            return f"{day:02d}{month:02d}{year}"
        else:
            return None
            
    except Exception as e:
        return None


def extract_yyyy_mm_dd(date_value):
    """
    Extract date in yyyy-mm-dd format from various datetime formats.

    Handles formats like:
    - '2025-12-01 00:01:24' --> 2025-12-01
    - '28-04-2026  18:13:41' --> 2026-04-28
    - '10/11/2025 17:06:01' --> 2025-11-10
    - datetime/Timestamp objects --> yyyy-mm-dd

    Args:
        date_value: Date value from Excel (can be string, datetime, or pandas Timestamp)

    Returns:
        str: Date in yyyy-mm-dd format, or None if parsing fails
    """
    try:
        if pd.isna(date_value) or date_value is None:
            return None

        if isinstance(date_value, (datetime, pd.Timestamp)):
            return date_value.strftime("%Y-%m-%d")

        date_str = re.sub(r"\s+", " ", str(date_value).strip())
        if not date_str or date_str.lower() in ['nan', 'none', '']:
            return None

        parsed_date = None
        date_part = date_str.split(" ", 1)[0]

        date_only_formats = (
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%d-%b-%Y",
            "%d.%m.%Y",
            "%m/%d/%Y",
        )
        datetime_formats = (
            "%d-%m-%Y %H:%M:%S",
            "%d/%m/%Y %H:%M:%S",
            "%Y-%m-%d %H:%M:%S",
            "%d-%b-%Y %H:%M:%S",
            "%d-%m-%Y %H:%M",
            "%d/%m/%Y %H:%M",
        )

        for fmt in datetime_formats:
            try:
                parsed_date = datetime.strptime(date_str, fmt)
                break
            except (ValueError, TypeError):
                continue

        if not parsed_date:
            for fmt in date_only_formats:
                try:
                    parsed_date = datetime.strptime(date_part, fmt)
                    break
                except (ValueError, TypeError):
                    continue

        if not parsed_date:
            try:
                parsed_date = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
                if isinstance(parsed_date, pd.Timestamp):
                    parsed_date = parsed_date.to_pydatetime()
                elif pd.isna(parsed_date):
                    parsed_date = None
            except Exception:
                parsed_date = None

        return parsed_date.strftime("%Y-%m-%d") if parsed_date else None

    except Exception:
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


def normalize_txn_id_for_comparison(value) -> str:
    """
    Normalize CCH TXN NO / Agency Txn Id values before comparison.

    - Strips all whitespace (spaces, tabs, non-breaking spaces)
    - Removes Excel text-prefix backticks (`) and other special characters
    - Keeps letters and digits only (case preserved)
    - Coerces whole-number floats (e.g. 12345.0) to integer strings

    Example: '`  ABC-123/45  ' -> 'ABC12345'
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    if isinstance(value, bool):
        text = str(value)
    elif isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if value.is_integer():
            text = str(int(value))
        else:
            text = str(value)
    else:
        text = str(value)

    # Excel often prefixes text cells with a grave accent/backtick
    text = text.replace("`", "")
    # Remove all whitespace characters
    text = re.sub(r"\s+", "", text.strip())
    # Remove remaining special characters; keep alphanumeric only
    return re.sub(r"[^A-Za-z0-9]", "", text)


def detect_lookup_txn_id_column(column_values, prefer_cch_txn_id=False):
    """
    Detect lookup transaction id column names, preferring Agency Txn Id or
    CCH Tran Id based on the configured priority.

    Args:
        column_values: Iterable of header strings or DataFrame columns

    Returns:
        tuple: (matched_value, source_type) where source_type is 'agency', 'cch', or None
    """
    values = list(column_values)

    column_groups = (
        [('cch', CCH_TXN_ID_COLUMN_NAMES), ('agency', AGENCY_TXN_ID_COLUMN_NAMES)]
        if prefer_cch_txn_id
        else [('agency', AGENCY_TXN_ID_COLUMN_NAMES), ('cch', CCH_TXN_ID_COLUMN_NAMES)]
    )

    for source_type, column_names in column_groups:
        for column_name in column_names:
            for value in values:
                value_text = str(value).strip() if value is not None else ''
                if not value_text:
                    continue
                value_lower = value_text.lower()
                column_name_lower = column_name.lower().strip()
                if column_name_lower in value_lower or value_lower in column_name_lower:
                    if len(value_lower) >= 5 or len(column_name_lower) >= 5:
                        return value, source_type

    return None, None


def resolve_lookup_txn_id_column(df, prefer_cch_txn_id=False):
    """
    Resolve the lookup transaction id column on a DataFrame.
    The preferred source is controlled by prefer_cch_txn_id. The matched
    column is renamed to the standard Agency Txn Id column so downstream
    logic stays unchanged.

    Returns:
        tuple: (resolved_column_name, source_type) or (None, None)
    """
    column_groups = (
        [('cch', CCH_TXN_ID_COLUMN_NAMES), ('agency', AGENCY_TXN_ID_COLUMN_NAMES)]
        if prefer_cch_txn_id
        else [('agency', AGENCY_TXN_ID_COLUMN_NAMES), ('cch', CCH_TXN_ID_COLUMN_NAMES)]
    )

    for source_type, column_names in column_groups:
        matched_column = find_column(df, column_names)
        if matched_column:
            if matched_column != STANDARD_AGENCY_TXN_ID_COL:
                df.rename(columns={matched_column: STANDARD_AGENCY_TXN_ID_COL}, inplace=True)
            return STANDARD_AGENCY_TXN_ID_COL, source_type

    return None, None


def resolve_settlement_type_column(df):
    """
    Resolve the best settlement/status column for lifecycle filtering.
    Avoid matching unrelated columns that merely contain the word 'Status'.

    Returns:
        str | None: Resolved column name
    """
    preferred_column_names = [
        'Settlement Type',
        'SettlementType',
        'Settlement_Type',
        'NPCI Settlement Status',
        'NPCI_Settlement_Status',
        'Transaction Status',
        'TransactionStatus',
        'Transaction_Status',
    ]

    for col_name in preferred_column_names:
        for col in df.columns:
            if str(col).strip().lower() == col_name.lower():
                return col

    for col_name in preferred_column_names:
        for col in df.columns:
            col_text = str(col).strip().lower()
            name_text = col_name.lower()
            if name_text in col_text or col_text in name_text:
                return col

    for col in df.columns:
        if str(col).strip().lower() == 'status':
            return col

    return None


def process_single_excel_file(excel_file_path, date_format='dd/mm/yyyy'):
    """
    Process a single Excel file with 'CCH TXN NO' column.
    This is for the comparison Excel file.
    
    Args:
        excel_file_path: Path to Excel file
        date_format: Not used anymore, kept for compatibility
    
    Returns:
        tuple: (list of CCH TXN NO values, DataFrame with normalized CCH TXN NO column)
    """
    try:
        # Read Excel file
        print(f"\nReading single Excel file: {excel_file_path}")
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        print(f"Total rows in Excel file: {len(df)}")
        print(f"Columns found: {list(df.columns)}")
        
        # Find CCH TXN NO column
        cch_txn_no_column_names = [
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
        ]
        
        cch_txn_no_col = find_column(df, cch_txn_no_column_names)
        if not cch_txn_no_col:
            print(f"Error: Could not find CCH TXN NO column.")
            print(f"Available columns: {list(df.columns)}")
            return [], None
        
        print(f"Found CCH TXN NO column: '{cch_txn_no_col}'")
        
        # Process each row - use existing column, just normalize values for comparison
        results = []
        
        for idx, row in df.iterrows():
            try:
                # Get CCH TXN NO value
                cch_txn_no_value = row[cch_txn_no_col]
                
                # Handle NaN or empty values
                if pd.isna(cch_txn_no_value) or str(cch_txn_no_value).strip() == '':
                    continue
                
                cch_txn_no_norm = normalize_txn_id_for_comparison(cch_txn_no_value)
                if not cch_txn_no_norm:
                    continue

                # Keep Excel cell readable; comparison uses normalized form
                df.at[idx, cch_txn_no_col] = str(cch_txn_no_value).strip().lstrip("`").strip()
                results.append(cch_txn_no_norm)
                
            except Exception as e:
                continue
        
        print(f"Extracted {len(results)} CCH TXN NO values from single Excel file")
        return results, df
        
    except Exception as e:
        print(f"Error processing single Excel file: {e}")
        import traceback
        traceback.print_exc()
        return [], None


def find_header_row(excel_file_path, max_rows=100, prefer_cch_txn_id=False):
    """
    Find the header row in an Excel file by searching through the first max_rows rows.
    
    Args:
        excel_file_path: Path to Excel file
        max_rows: Maximum number of rows to search (default: 100)
    
    Returns:
        int: Row index of header row (0-based), or None if not found
    """
    try:
        file_name = Path(excel_file_path).name
        print(f"    Searching for header row in: {file_name}")

        file_path = Path(excel_file_path)
        suffix = file_path.suffix.lower()

        # Read file without headers to search for header row
        if suffix == ".csv":
            df_raw = pd.read_csv(excel_file_path, header=None, nrows=max_rows, dtype=str)
        else:
            df_raw = pd.read_excel(excel_file_path, engine='openpyxl', header=None, nrows=max_rows)
        
        print(f"    Total rows to check: {len(df_raw)}")
        
        violation_column_names = [
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
        
        # Search through rows to find header
        for row_idx in range(min(max_rows, len(df_raw))):
            row_values = [str(val).strip() if pd.notna(val) else '' for val in df_raw.iloc[row_idx].values]
            
            # Print first 20 rows for debugging
            if row_idx < 20:
                print(f"    Row {row_idx + 1}: {row_values[:10]}")  # Print first 10 columns
            
            # Check if this row contains both required column names
            lookup_txn_id_match, lookup_txn_id_source = detect_lookup_txn_id_column(
                row_values,
                prefer_cch_txn_id=prefer_cch_txn_id,
            )
            found_lookup_txn_id_col = lookup_txn_id_match is not None
            found_violation_col = False
            violation_match = None
            
            for val in row_values:
                # Skip empty strings - they cause false matches
                if not val or val.strip() == '':
                    continue
                
                val_lower = val.lower().strip()
                
                # Check for violation column - require meaningful match
                for violation_name in violation_column_names:
                    violation_name_lower = violation_name.lower().strip()
                    if violation_name_lower and val_lower:
                        if violation_name_lower in val_lower or val_lower in violation_name_lower:
                            # Additional check: make sure it's a meaningful match (at least 5 characters)
                            if len(val_lower) >= 5 or len(violation_name_lower) >= 5:
                                found_violation_col = True
                                violation_match = val
                                break
                
                # We need BOTH lookup txn id and violation amount columns to be found
                if found_lookup_txn_id_col and found_violation_col:
                    print(f"    ✓ Header found at row {row_idx + 1} (0-based index: {row_idx})")
                    if lookup_txn_id_source == 'cch':
                        print(
                            f"      Lookup txn id column match: '{lookup_txn_id_match}' "
                            f"(using CCH fallback, will normalize to '{STANDARD_AGENCY_TXN_ID_COL}')"
                        )
                    else:
                        print(f"      Lookup txn id column match: '{lookup_txn_id_match}'")
                    print(f"      Violation column match: '{violation_match}'")
                    print(f"    Header row values: {row_values}")
                    return row_idx
            
            # Print partial matches for debugging (first 20 rows only)
            if row_idx < 20:
                if found_lookup_txn_id_col:
                    if lookup_txn_id_source == 'cch':
                        print(
                            f"    Row {row_idx + 1}: Found fallback CCH Tran Id column "
                            f"('{lookup_txn_id_match}')"
                        )
                    else:
                        print(
                            f"    Row {row_idx + 1}: Found Agency Txn Id column "
                            f"('{lookup_txn_id_match}')"
                        )
                if found_violation_col:
                    print(f"    Row {row_idx + 1}: Found Violation column ('{violation_match}')")
        
        print(f"    ✗ Header row not found in first {min(max_rows, len(df_raw))} rows")
        print(f"    Showing first row as reference: {[str(val).strip() if pd.notna(val) else '' for val in df_raw.iloc[0].values[:10]]}")
        return None
        
    except Exception as e:
        print(f"    Error finding header row: {e}")
        import traceback
        traceback.print_exc()
        return None


def process_excel_file_from_folder(excel_file_path, prefer_cch_txn_id=False):
    """
    Process an Excel file from folder with 'Agency Txn Id' column.
    Also extracts 'Violation Amt', 'Reader Read Time', and 'Journey Type' columns if available.
    Searches for header row in first 100 rows if not found in first row.
    
    Args:
        excel_file_path: Path to Excel file
    
    Returns:
        list: List of tuples (agency_txn_id, violation_amt, reader_read_time, journey_type) 
              where violation_amt, reader_read_time, and journey_type can be None
    """
    try:
        # First, try to find header row
        header_row = find_header_row(
            excel_file_path,
            max_rows=100,
            prefer_cch_txn_id=prefer_cch_txn_id,
        )
        file_path = Path(excel_file_path)
        suffix = file_path.suffix.lower()

        if header_row is None:
            # Try reading with default header (row 0)
            if suffix == ".csv":
                df = pd.read_csv(excel_file_path, header=0, dtype=str)
            else:
                df = pd.read_excel(excel_file_path, engine='openpyxl', header=0)
        else:
            # Read with found header row
            if header_row > 0:
                print(f"    Found header at row {header_row + 1} (skipping {header_row} rows)")
            if suffix == ".csv":
                df = pd.read_csv(excel_file_path, header=header_row, dtype=str)
            else:
                df = pd.read_excel(excel_file_path, engine='openpyxl', header=header_row)
        
        agency_txn_id_col, txn_id_source = resolve_lookup_txn_id_column(
            df,
            prefer_cch_txn_id=prefer_cch_txn_id,
        )
        if not agency_txn_id_col:
            print(
                "    Warning: Could not find 'Agency Txn Id' column in this file "
                "or fallback 'CCH Tran Id' column"
            )
            return []

        if txn_id_source == 'cch':
            print(
                f"    Found fallback CCH Tran Id column and renamed it to "
                f"'{STANDARD_AGENCY_TXN_ID_COL}'"
            )
        else:
            print(f"    Found Agency Txn Id column: '{agency_txn_id_col}'")
        
        # Find 'Violation Amt' column (optional) - with variations including trailing spaces/tabs
        violation_column_names = [
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
        
        violation_col = find_column(df, violation_column_names)
        if violation_col:
            if violation_col != STANDARD_VIOLATION_AMT_COL:
                print(f"    Found violation amount column: '{violation_col}' (will write as '{STANDARD_VIOLATION_AMT_COL}')")
            else:
                print(f"    Found violation amount column: '{violation_col}'")
        else:
            print(f"    Warning: Could not find '{STANDARD_VIOLATION_AMT_COL}' column in this file")
        
        # Find 'Reader Read Time' column (optional)
        reader_read_time_column_names = [
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
            'Transaction Date Time',
            'Transaction Date Time ',
            'Transaction Date Time\t',
            'Transaction Date Time.',
            'Transaction Date Time. ',
            'Transaction Date Time.\t',
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
        
        reader_read_time_col = find_column(df, reader_read_time_column_names)
        if reader_read_time_col:
            if reader_read_time_col != STANDARD_READER_READ_TIME_COL:
                print(f"    Found Reader Read Time column: '{reader_read_time_col}' (will write as '{STANDARD_READER_READ_TIME_COL}')")
            else:
                print(f"    Found Reader Read Time column: '{reader_read_time_col}'")
        else:
            print(f"    Warning: Could not find '{STANDARD_READER_READ_TIME_COL}' column in this file")
        
        # Find 'Journey Type' column (optional)
        journey_type_col = find_column(df, JOURNEY_TYPE_COLUMN_NAMES)
        if journey_type_col:
            print(f"    Found Journey/Trip/Fare Type column: '{journey_type_col}'")
        else:
            print(
                "    Warning: Could not find 'Journey Type', 'Trip Type', or 'Fare Type' column in this file"
            )
        
        # Find settlement/status column for filtering
        settlement_type_col = resolve_settlement_type_column(df)
        if not settlement_type_col:
            print(
                "    Warning: Could not find 'Settlement Type', "
                "'NPCI Settlement Status', or 'Transaction Status' column in this file"
            )
            print(f"    All rows will be processed (no filtering applied)")
        else:
            print(f"    Found Settlement Type column: '{settlement_type_col}'")
        
        # Filter rows: only include rows where settlement status is SETTLED or ACCEPTED
        if settlement_type_col:
            initial_row_count = len(df)
            valid_settlement_values = {'SETTLED', 'ACCEPTED'}
            df_filtered = df[
                df[settlement_type_col].astype(str).str.strip().str.upper().isin(valid_settlement_values)
            ].copy()
            filtered_row_count = len(df_filtered)
            print(
                f"    Filtered rows: {initial_row_count} -> {filtered_row_count} "
                f"(keeping only 'SETTLED' or 'ACCEPTED' rows)"
            )
        else:
            df_filtered = df.copy()
            print(f"    No filter applied: processing all {len(df_filtered)} rows")
        
        # Process each row (only filtered rows)
        results = []
        
        for idx, row in df_filtered.iterrows():
            try:
                # Get Agency Txn Id value
                agency_txn_id_value = row[agency_txn_id_col]
                
                # Handle NaN or empty values
                if pd.isna(agency_txn_id_value) or str(agency_txn_id_value).strip() == '':
                    continue
                
                agency_txn_id_str = normalize_txn_id_for_comparison(agency_txn_id_value)
                if not agency_txn_id_str:
                    continue
                
                # Get violation amount if column exists
                violation_amt = None
                if violation_col:
                    violation_value = row[violation_col]
                    if pd.notna(violation_value):
                        # Try to convert to numeric, keep as string if conversion fails
                        try:
                            violation_amt = float(violation_value)
                        except (ValueError, TypeError):
                            violation_amt = str(violation_value).strip()
                            if violation_amt == '':
                                violation_amt = None
                
                # Get Reader Read Time if column exists
                reader_read_time = None
                if reader_read_time_col:
                    reader_read_time_value = row[reader_read_time_col]
                    if pd.notna(reader_read_time_value):
                        reader_read_time = extract_yyyy_mm_dd(reader_read_time_value)
                        if reader_read_time is None:
                            reader_read_time = str(reader_read_time_value).strip()
                
                # Get Journey Type if column exists
                journey_type = None
                if journey_type_col:
                    journey_type_value = row[journey_type_col]
                    if pd.notna(journey_type_value):
                        # Keep as string
                        journey_type = str(journey_type_value).strip()
                        if journey_type == '':
                            journey_type = None
                
                results.append((agency_txn_id_str, violation_amt, reader_read_time, journey_type))
                
            except Exception as e:
                continue
        
        return results
        
    except Exception as e:
        return []


def process_folder(folder_path, prefer_cch_txn_id=False):
    """
    Process all Excel files in a folder.
    
    Args:
        folder_path: Path to folder containing Excel files
    
    Returns:
        tuple: (list of Agency Txn Id strings, list of violation amounts, list of reader read times, list of journey types)
               All lists are parallel (same length, matching indices)
    """
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"Error: Folder not found: {folder_path}")
        return [], [], [], []
    
    if not folder.is_dir():
        print(f"Error: Path is not a folder: {folder_path}")
        return [], [], [], []
    
    # Find all supported files in the folder
    report_files = (
        list(folder.glob("*.xlsx"))
        + list(folder.glob("*.xls"))
        + list(folder.glob("*.csv"))
    )

    if not report_files:
        print(f"No Excel/CSV files found in folder: {folder_path}")
        return [], [], [], []

    print(f"\nFound {len(report_files)} Excel/CSV file(s) in folder")
    print(f"{'='*80}")
    
    all_results = []
    all_violation_amts = []
    all_reader_read_times = []
    all_journey_types = []
    
    for report_file in report_files:
        print(f"\nProcessing: {report_file.name}")
        file_results = process_excel_file_from_folder(
            report_file,
            prefer_cch_txn_id=prefer_cch_txn_id,
        )
        
        # Separate Agency Txn Id strings, violation amounts, reader read times, and journey types
        agency_txn_id_list = [item[0] for item in file_results]
        violation_list = [item[1] for item in file_results]
        reader_read_time_list = [item[2] for item in file_results]
        journey_type_list = [item[3] for item in file_results]
        
        all_results.extend(agency_txn_id_list)
        all_violation_amts.extend(violation_list)
        all_reader_read_times.extend(reader_read_time_list)
        all_journey_types.extend(journey_type_list)
        print(f"  Extracted {len(file_results)} Agency Txn Id values from this file")
    
    return all_results, all_violation_amts, all_reader_read_times, all_journey_types


def detect_output_files(folder_path):
    """
    Detect all violation lookup output files in the folder using keywords.
    Searches for files matching: violation, lookup, output
    
    Args:
        folder_path: Path to folder
    
    Returns:
        list: List of detected output file paths, sorted by index
    """
    folder = Path(folder_path)
    detected_files = []
    
    # Keywords to search for
    keywords = ['violation', 'lookup', 'output']
    
    # Search for Excel files in the folder
    for file_path in folder.glob('*.xlsx'):
        file_name_lower = file_path.name.lower()
        
        # Check if filename contains any of the keywords
        if any(keyword in file_name_lower for keyword in keywords):
            detected_files.append(str(file_path))
    
    # Sort files by index if they have numbered suffixes (e.g., violation_lookup_output_1.xlsx)
    def extract_index(filepath):
        """Extract index number from filename for sorting"""
        match = re.search(r'_(\d+)\.xlsx$', filepath, re.IGNORECASE)
        if match:
            return int(match.group(1))
        return 0  # If no index found, put it first
    
    detected_files.sort(key=extract_index)
    
    return detected_files


def match_and_generate_output(
    excel_file_path,
    df,
    folder_path,
    folder_results,
    violation_amts,
    reader_read_times,
    journey_types,
    write_matched_output=False,
):
    """
    Match CCH TXN NO against folder Agency Txn Id data, update the combined Excel,
    and optionally write violation_lookup_output_{n}.xlsx files.

    Args:
        write_matched_output: If True, write matched-only output Excel files.
            Prefer False — DB update uses the updated combined report, not these files.
    """
    try:
        folder = Path(folder_path)
        
        # ---- STEP A: Build folder lookup dict from in-memory data ----
        print(f"\nBuilding folder lookup dictionary...")
        
        # Build dict: agency_txn_id -> (violation_amt, reader_read_time, journey_type)
        # For duplicates, keep the one with non-zero violation amount
        folder_dict = {}
        for agency_txn_id, viol_amt, reader_read_time, journey_type in zip(folder_results, violation_amts, reader_read_times, journey_types):
            agency_txn_id = normalize_txn_id_for_comparison(agency_txn_id)
            if not agency_txn_id:
                continue
            
            # Convert violation amount to numeric
            violation_value = None
            if viol_amt is not None and pd.notna(viol_amt):
                try:
                    violation_value = float(viol_amt)
                except (ValueError, TypeError):
                    try:
                        violation_value = int(viol_amt)
                    except (ValueError, TypeError):
                        violation_value = None
            
            # Keep entry with non-zero violation if duplicate
            if agency_txn_id not in folder_dict or (violation_value is not None and violation_value != 0):
                folder_dict[agency_txn_id] = (violation_value, reader_read_time, journey_type)
        
        print(f"  Total unique entries in folder: {len(folder_dict)}")
        
        # ---- STEP B: Compare & mark matches, collect matched entries for output ----
        print(f"\nComparing single file rows (CCH TXN NO) against folder data (Agency Txn Id)...")
        print(
            "  (IDs normalized for match: whitespace/backticks/special chars removed, "
            "letters and digits kept)"
        )
        
        # Create/initialize standard columns in single combined report
        # (lifecycle aliases such as Adjustment Amount / Transaction Date/Time are normalized here)
        if STANDARD_JOURNEY_TYPE_COL not in df.columns:
            df[STANDARD_JOURNEY_TYPE_COL] = None
        if STANDARD_VIOLATION_AMT_COL not in df.columns:
            df[STANDARD_VIOLATION_AMT_COL] = None
        if STANDARD_READER_READ_TIME_COL not in df.columns:
            df[STANDARD_READER_READ_TIME_COL] = None
        df['violation_matched'] = ''
        
        yes_count = 0
        no_count = 0
        matched_entries = []  # Collect (cch_txn_no, Violation Amt, Reader Read Time) for output file
        
        # Find the CCH TXN NO column in the DataFrame
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
            print("Error: Could not find CCH TXN NO column in DataFrame")
            return False
        
        for idx, row in df.iterrows():
            cch_txn_no_raw = row.get(cch_txn_no_col, '')
            cch_txn_no = normalize_txn_id_for_comparison(cch_txn_no_raw)
            cch_txn_no_display = (
                str(cch_txn_no_raw).strip().lstrip("`").strip()
                if pd.notna(cch_txn_no_raw)
                else ""
            )
            
            if cch_txn_no and cch_txn_no in folder_dict:
                violation_amt, reader_read_time, journey_type = folder_dict[cch_txn_no]
                
                # Write values to single Excel file for ALL matches (standard column names)
                df.at[idx, STANDARD_JOURNEY_TYPE_COL] = journey_type
                df.at[idx, STANDARD_READER_READ_TIME_COL] = reader_read_time
                df.at[idx, STANDARD_VIOLATION_AMT_COL] = violation_amt
                
                # Check if violation amount exists and is non-zero for violation_matched column
                if violation_amt is not None:
                    try:
                        violation_value = float(violation_amt)
                        if violation_value != 0:
                            df.at[idx, 'violation_matched'] = 'Yes'
                            yes_count += 1
                            # Collect for output file (CCH TXN NO, Violation Amt, Reader Read Time)
                            matched_entries.append(
                                (cch_txn_no_display or cch_txn_no, violation_amt, reader_read_time)
                            )
                        else:
                            df.at[idx, 'violation_matched'] = 'No'
                            no_count += 1
                    except (ValueError, TypeError):
                        df.at[idx, 'violation_matched'] = 'No'
                        no_count += 1
                else:
                    df.at[idx, 'violation_matched'] = 'No'
                    no_count += 1
            else:
                df.at[idx, 'violation_matched'] = 'No'
                no_count += 1
        
        # ---- STEP C: Save updated single file ----
        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        
        print(f"\n✓ Updated single Excel file: {excel_file_path}")
        print(
            f"  Added/updated columns: '{STANDARD_JOURNEY_TYPE_COL}', "
            f"'{STANDARD_VIOLATION_AMT_COL}', '{STANDARD_READER_READ_TIME_COL}', 'violation_matched'"
        )
        print(f"  Marked {yes_count} rows as 'Yes' (matched with non-zero violation amount)")
        print(f"  Marked {no_count} rows as 'No' (no match or zero/empty violation amount)")
        
        # ---- STEP D: Optionally write output file(s) with ONLY matched entries ----
        # Skipped by default — combined report already has the matched data used for DB updates.
        if not write_matched_output:
            print(f"\n  Skipping violation_lookup_output file generation "
                  f"({yes_count} non-zero match(es) already written to combined report).")
            return True

        if not matched_entries:
            print(f"\n  No matched entries to write to output file.")
            return True

        # Remove duplicate matched entries
        df_matched = pd.DataFrame(
            matched_entries,
            columns=['CCH TXN NO', STANDARD_VIOLATION_AMT_COL, STANDARD_READER_READ_TIME_COL],
        )
        df_matched = df_matched.drop_duplicates(subset=['CCH TXN NO']).sort_values('CCH TXN NO').reset_index(drop=True)

        total_matched = len(df_matched)
        print(f"\n  Total unique matched entries for output: {total_matched}")

        MAX_EXCEL_ROWS = 1048576
        MAX_DATA_ROWS = MAX_EXCEL_ROWS - 1  # Account for header row

        if total_matched <= MAX_DATA_ROWS:
            output_file = folder / "violation_lookup_output_1.xlsx"
            df_matched.to_excel(output_file, index=False, engine='openpyxl')

            print(f"✓ Created output file: {output_file}")
            print(f"  Matched rows written: {total_matched}")
        else:
            num_files = (total_matched // MAX_DATA_ROWS) + (1 if total_matched % MAX_DATA_ROWS > 0 else 0)

            print(f"  Matched data exceeds Excel limit. Splitting into {num_files} file(s)...")

            for i in range(num_files):
                start_idx = i * MAX_DATA_ROWS
                end_idx = min((i + 1) * MAX_DATA_ROWS, total_matched)

                df_chunk = df_matched.iloc[start_idx:end_idx].copy()
                output_file = folder / f"violation_lookup_output_{i+1}.xlsx"

                df_chunk.to_excel(output_file, index=False, engine='openpyxl')
                print(f"✓ Created output file {i+1}/{num_files}: {output_file}")
                print(f"  Rows: {len(df_chunk)}")

            print(f"\n✓ Created {num_files} output file(s) total")

        return True
        
    except Exception as e:
        print(f"Error in match_and_generate_output: {e}")
        import traceback
        traceback.print_exc()
        return False


def compare_results(single_file_results, folder_results):
    """
    Compare results from single Excel file with results from folder files.
    
    Args:
        single_file_results: List of CCH TXN NO values from single Excel file
        folder_results: List of Agency Txn Id values from folder files
    
    Returns:
        dict: Comparison statistics
    """
    # Convert to sets for comparison (alphanumeric-normalized IDs)
    single_file_set = {normalize_txn_id_for_comparison(v) for v in single_file_results}
    single_file_set.discard("")
    folder_set = {normalize_txn_id_for_comparison(v) for v in folder_results}
    folder_set.discard("")
    
    # Find matches (values present in both)
    matches = single_file_set.intersection(folder_set)
    
    # Find values only in single file
    only_in_single = single_file_set - folder_set
    
    # Find values only in folder
    only_in_folder = folder_set - single_file_set
    
    return {
        'matches': list(matches),
        'only_in_single': list(only_in_single),
        'only_in_folder': list(only_in_folder),
        'total_single': len(single_file_results),
        'total_folder': len(folder_results),
        'unique_single': len(single_file_set),
        'unique_folder': len(folder_set),
        'match_count': len(matches)
    }


def main():
    """Main function"""
    # Get Excel file and folder path from command line arguments or use defaults

    excel_file_path = "bassi_combined_2025-11-23_to_2026-02-15.xlsx"
    folder_path = "Bassi"
    
    print("="*80)
    print("Violation Lookup Program - Comparison")
    print("="*80)
    print(f"Single Excel file: {excel_file_path}")
    print(f"Folder: {folder_path}")
    print("="*80)
    
    # Check if Excel file exists
    if not Path(excel_file_path).exists():
        print(f"\nError: Excel file not found: {excel_file_path}")
        print("Usage: python Violation_lookup.py <excel_file_path> <folder_path>")
        print("Or update the default paths in the main() function.")
        return
    
    # Check if folder exists
    if not Path(folder_path).exists():
        print(f"\nError: Folder not found: {folder_path}")
        print("Usage: python Violation_lookup.py <excel_file_path> <folder_path>")
        print("Or update the default paths in the main() function.")
        return
    
    # Process single Excel file (no date format needed anymore)
    print(f"\n{'='*80}")
    print("STEP 1: Processing single Excel file")
    print(f"{'='*80}")
    single_file_results, df_single = process_single_excel_file(excel_file_path)
    
    if not single_file_results or df_single is None:
        print("Error: Could not process single Excel file. Exiting.")
        return
    
    # Process all Excel files in the folder
    print(f"\n{'='*80}")
    print("STEP 2: Processing Excel files in folder")
    print(f"{'='*80}")
    folder_results, violation_amts, reader_read_times, journey_types = process_folder(folder_path)
    
    if not folder_results:
        print("Warning: No results from folder files. Skipping comparison.")
        return
    
    # Match, mark violation_matched, and generate output file(s) with ONLY matched entries
    print(f"\n{'='*80}")
    print("STEP 3: Matching & generating output")
    print(f"{'='*80}")
    print("Flow: Compare single file rows (CCH TXN NO) with folder data (Agency Txn Id)")
    print("  -> For ALL matches: Write 'Journey Type', 'Reader Read Time', and 'Violation Amt' to single Excel file")
    print("  -> If Violation Amt is non-zero -> mark 'Yes' in violation_matched + write to output")
    print("  -> Otherwise -> mark 'No' in violation_matched")
    print("  -> Output file(s) contain ONLY matched entries with non-zero Violation Amt (CCH TXN NO + Violation Amt + Reader Read Time)")
    
    success = match_and_generate_output(excel_file_path, df_single, folder_path, folder_results, violation_amts, reader_read_times, journey_types)
    
    if not success:
        print("Error: Could not complete matching and output generation.")
        return
    
    # Compare results for statistics
    print(f"\n{'='*80}")
    print("STEP 4: Comparison Statistics")
    print(f"{'='*80}")
    
    comparison = compare_results(single_file_results, folder_results)
    
    print(f"\nComparison Statistics:")
    print(f"  Total values in single file: {comparison['total_single']} (unique: {comparison['unique_single']})")
    print(f"  Total values in folder files: {comparison['total_folder']} (unique: {comparison['unique_folder']})")
    print(f"  Matches (found in both): {comparison['match_count']}")
    print(f"  Only in single file: {len(comparison['only_in_single'])}")
    print(f"  Only in folder files: {len(comparison['only_in_folder'])}")
    
    # Print 10-15 sample outputs from matches
    if comparison['matches']:
        sample_count = min(15, len(comparison['matches']))
        print(f"\n{'='*80}")
        print(f"Sample matches (showing {sample_count} results):")
        print(f"{'='*80}")
        for i, match in enumerate(comparison['matches'][:sample_count], 1):
            print(f"  {i}. {match}")
        
        if len(comparison['matches']) > sample_count:
            print(f"\n  ... and {len(comparison['matches']) - sample_count} more matches")
    
    print(f"\n{'='*80}")
    print("Processing complete!")
    print(f"{'='*80}")


if __name__ == "__main__":
    main()
