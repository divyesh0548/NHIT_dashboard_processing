"""
Aggregate plaza Excel transaction data into hourly analytics rows and upsert into PostgreSQL.

Configure PLAZA_NAME, COLUMN_MAPPING, normalization lists, and INPUT_FOLDER below.
"""

from __future__ import annotations

import re
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import psycopg2
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from psycopg2 import sql
from psycopg2.extras import execute_values

ROOT_DIR = Path(__file__).resolve().parent.parent
INSIGHTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR))
sys.path.insert(0, str(INSIGHTS_DIR))

from config.db import (
    ensure_analytics_table,
    get_analytics_db_connection_kwargs,
    get_analytics_table_name,
)
from config.mappings import load_mop_mappings, load_vehicle_class_mappings


# ---------------------------------------------------------------------------
# Configuration — update for each plaza / Excel layout
# ---------------------------------------------------------------------------

PLAZA_NAME = "aroli"

INPUT_FOLDER = r"C:\Divyesh\NHIT_dashboard_processing\New Dashboard Insights\Input"

# If True, print prepared hourly rows only — no DB connection or inserts.
DRY_RUN = False

# Excel column name -> logical field
COLUMN_MAPPING = {
    "datetime": "DATE",
    "vehicle_class": "MVC",
    "lane_no": "LANE",
    "mop": "MVC MOP",
}

# Each entry: canonical value -> (db_column, excel_aliases)
# Edit config/vehicle_class.json and config/mop.json to update mappings.
VEHICLE_CLASS_MAPPINGS = load_vehicle_class_mappings()
MOP_MAPPINGS = load_mop_mappings()

TOTAL_TRANSACTION_COLUMN = "total_transaction"

# Add lane aliases here when you define plaza-specific lane parsing rules.
LANE_MAPPINGS: dict[str, tuple[str, list[str]]] = {
    f"L{i:02d}": (f"l{i:02d}", [f"L{i:02d}"]) for i in range(1, 13)
}

EXCEL_EXTENSIONS = {".xlsx", ".xls", ".csv"}

# Try these formats per file; the one that parses the most values wins.
DATETIME_FORMATS_12H = [
    "%d-%b-%Y %I:%M:%S %p",  # 06-Dec-2025 12:00:52 AM
    "%d-%b-%Y %I:%M %p",     # 06-Dec-2025 2:00 PM
    "%m/%d/%Y %I:%M:%S %p",  # 11/24/2025 2:23:35 AM
    "%m/%d/%Y %I:%M %p",     # 11/24/2025 2:23 AM
]

DATETIME_FORMATS_24H = [
    "%d/%m/%Y %H:%M:%S",     # 24/11/2025 14:23:35
    "%d/%m/%Y %H:%M",
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
]

DATETIME_FORMATS = DATETIME_FORMATS_12H + DATETIME_FORMATS_24H

AM_PM_PATTERN = re.compile(r"\b(AM|PM)\b", re.IGNORECASE)

# Hour bucket stored as TEXT in 24-hour form: "2-3" means 02:00:00-02:59:59.
# Use 24-hour labels (not "2-3 AM") so afternoon hours stay unambiguous ("14-15").

HEADER_KEYWORDS = [
    "Transaction ID",
    "Date & Time",
    "Veh Reg No.",
    "TC Class",
    "Lane No",
    "VEH REG NO",
    "TRANSACTION NO",
    "TC_VEH_REG_NO",
    "PLAZA_NAME",
    "MVC_TLC_CLASS",
    "MVC_TLC_MOP",
    "MVC MOP",
    "MVC",
    "Veh Reg Num",
    "PaymentType",
    "VehicleNumber",
    "Journey Type",
    "Vehicle No",
    "Operator Class",
    "DATE",
    "LANE",
]
# Include mapped Excel column names in header detection.
HEADER_KEYWORDS.extend(COLUMN_MAPPING.values())


class UnmappedValueError(Exception):
    """Raised when a value cannot be normalized; execution should stop."""


def normalize_key(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


def is_blank(value) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def build_lookup(normalization: dict[str, list[str]]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for canonical, aliases in normalization.items():
        lookup[normalize_key(canonical)] = canonical
        for alias in aliases:
            lookup[normalize_key(alias)] = canonical
    return lookup


def split_mappings(
    mappings: dict[str, tuple[str, list[str]]],
) -> tuple[dict[str, str], dict[str, list[str]]]:
    """Split combined config into db-column map and excel-alias map."""
    columns = {canonical: db_column for canonical, (db_column, _aliases) in mappings.items()}
    normalization = {canonical: aliases for canonical, (_db_column, aliases) in mappings.items()}
    return columns, normalization


VEHICLE_CLASS_COLUMNS, VEHICLE_CLASS_NORMALIZATION = split_mappings(VEHICLE_CLASS_MAPPINGS)
MOP_COLUMNS, MOP_NORMALIZATION = split_mappings(MOP_MAPPINGS)
LANE_COLUMNS, LANE_NORMALIZATION = split_mappings(LANE_MAPPINGS)

VEHICLE_CLASS_LOOKUP = build_lookup(VEHICLE_CLASS_NORMALIZATION)
MOP_LOOKUP = build_lookup(MOP_NORMALIZATION)
LANE_LOOKUP = build_lookup(LANE_NORMALIZATION)


def create_run_log_file() -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = INSIGHTS_DIR / f"log-{timestamp}.txt"
    log_path.write_text(
        "module1 processing log\n"
        f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"{'=' * 80}\n",
        encoding="utf-8",
    )
    return log_path


def append_run_log(log_path: Path, message: str) -> None:
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(message + "\n")


def is_header_detection_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return "header" in message and "detect" in message


def try_parse_datetime(value, datetime_format: str | None = None):
    try:
        return parse_datetime(value, datetime_format=datetime_format)
    except (ValueError, TypeError):
        return None


def datetime_has_am_pm(text: str) -> bool:
    return bool(AM_PM_PATTERN.search(text))


def datetime_formats_for_text(text: str) -> list[str]:
    """Use 12-hour formats when AM/PM is present; otherwise 24-hour formats."""
    if datetime_has_am_pm(text):
        return DATETIME_FORMATS_12H
    return DATETIME_FORMATS_24H


def try_parse_with_format(value, datetime_format: str):
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    text = str(value).strip()
    if not text:
        return None

    if "%p" in datetime_format and not datetime_has_am_pm(text):
        return None
    if "%p" not in datetime_format and datetime_has_am_pm(text):
        return None

    try:
        return datetime.strptime(text, datetime_format)
    except ValueError:
        return None


def parse_datetime_with_format(value, datetime_format: str) -> datetime:
    if pd.isna(value):
        raise ValueError("Empty datetime value")
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    text = str(value).strip()
    return datetime.strptime(text, datetime_format)


def detect_datetime_format(values) -> str:
    """
    Detect the datetime format used in a file by testing DATETIME_FORMATS
    against non-empty sample values.
    """
    sample = [
        value
        for value in values
        if not pd.isna(value) and str(value).strip()
    ]
    if not sample:
        raise ValueError("No datetime values found to detect format.")

    best_format = None
    best_count = -1

    for fmt in DATETIME_FORMATS:
        parsed_count = sum(
            1 for value in sample if try_parse_with_format(value, fmt) is not None
        )
        if parsed_count > best_count:
            best_count = parsed_count
            best_format = fmt

    if best_format is None or best_count == 0:
        raise ValueError("Could not detect datetime format from file values.")

    if best_count < len(sample):
        print(
            f"  Warning: detected format '{best_format}' parsed "
            f"{best_count}/{len(sample)} datetime sample(s)."
        )

    return best_format


def parse_datetime(value, datetime_format: str | None = None) -> datetime:
    if pd.isna(value):
        raise ValueError("Empty datetime value")
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    if datetime_format:
        return parse_datetime_with_format(value, datetime_format)

    text = str(value).strip()
    for fmt in datetime_formats_for_text(text):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        raise ValueError(f"Unable to parse datetime: {value!r}")
    return parsed.to_pydatetime()


def hour_bucket_label(dt: datetime) -> str:
    """Return 24-hour bucket label, e.g. 2:23 AM -> '2-3', 2:23 PM -> '14-15'."""
    return f"{dt.hour}-{dt.hour + 1}"


def try_normalize_vehicle_class(value):
    key = normalize_key(value)
    if not key:
        return None
    return VEHICLE_CLASS_LOOKUP.get(key)


def try_normalize_lane(value):
    key = normalize_key(value)
    if not key:
        return None
    return LANE_LOOKUP.get(key)


def try_normalize_mop(value):
    key = normalize_key(value)
    if not key:
        return None
    return MOP_LOOKUP.get(key)


def optional_normalize_vehicle_class(value):
    if is_blank(value):
        return None
    return try_normalize_vehicle_class(value)


def optional_normalize_lane(value):
    if is_blank(value):
        return None
    return try_normalize_lane(value)


def optional_normalize_mop(value):
    if is_blank(value):
        return None
    return try_normalize_mop(value)


def safe_parse_datetime(value, datetime_format: str):
    if is_blank(value):
        return pd.NaT
    try:
        return parse_datetime(value, datetime_format=datetime_format)
    except (ValueError, TypeError):
        return pd.NaT


def normalize_lane(value) -> str:
    normalized = try_normalize_lane(value)
    if normalized is None:
        raise UnmappedValueError(f"Unmapped lane value: {value!r}")
    return normalized


def normalize_vehicle_class(value) -> str:
    normalized = try_normalize_vehicle_class(value)
    if normalized is None:
        raise UnmappedValueError(f"Unmapped vehicle class value: {value!r}")
    return normalized


def normalize_mop(value) -> str:
    normalized = try_normalize_mop(value)
    if normalized is None:
        raise UnmappedValueError(f"Unmapped MOP value: {value!r}")
    return normalized


def find_header_row(df: pd.DataFrame) -> int | None:
    """
    Find the header row by matching COLUMN_MAPPING Excel names as exact cell values.

    This avoids false positives from report metadata rows such as
    'FROM DATE' / 'TO DATE' that only contain DATE as a substring.
    """
    required_columns = list(COLUMN_MAPPING.values())
    required_normalized = [normalize_key(column) for column in required_columns]

    best_row = None
    best_score = 0

    for row_index in range(len(df)):
        row_cells = {
            normalize_key(value)
            for value in df.iloc[row_index]
            if not pd.isna(value) and str(value).strip()
        }
        score = sum(1 for column in required_normalized if column in row_cells)
        if score > best_score:
            best_score = score
            best_row = row_index

    if best_score >= 2:
        return best_row

    return None


def remove_duplicate_header_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Remove consecutive duplicate header-like rows from the raw sheet."""
    rows_to_remove: list[int] = []

    for row_index in range(len(df) - 1):
        current_row = df.iloc[row_index]
        next_row = df.iloc[row_index + 1]
        current_str = current_row.fillna("").astype(str)
        next_str = next_row.fillna("").astype(str)

        if not current_str.equals(next_str):
            continue

        has_text = any(
            pd.notna(value) and isinstance(value, str) and str(value).strip()
            for value in current_row
        )
        if has_text:
            rows_to_remove.append(row_index)

    if not rows_to_remove:
        return df

    return df.drop(index=rows_to_remove).reset_index(drop=True)


def unmerge_workbook(input_path: Path, output_path: Path, sheet_name: str | None = None) -> str:
    """Unmerge all merged cells and save to a new workbook."""
    workbook = load_workbook(input_path)
    worksheet = workbook[sheet_name] if sheet_name else workbook.active
    resolved_sheet_name = worksheet.title

    for merged_range in list(worksheet.merged_cells):
        worksheet.unmerge_cells(range_string=str(merged_range))

    workbook.save(output_path)
    return resolved_sheet_name


def remove_empty_columns(filename: Path, sheet_name: str, header_row: int) -> Path:
    """Delete columns with empty headers or no data below the header row."""
    workbook = load_workbook(filename)
    worksheet = workbook[sheet_name]
    max_col = worksheet.max_column
    cols_to_delete: list[int] = []

    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        header_value = worksheet[f"{col_letter}{header_row}"].value

        if header_value is None or str(header_value).strip() == "":
            cols_to_delete.append(col_idx)
            continue

        empty_below = True
        for row_idx in range(header_row + 1, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value not in (None, ""):
                empty_below = False
                break

        if empty_below:
            cols_to_delete.append(col_idx)

    for col_idx in sorted(cols_to_delete, reverse=True):
        worksheet.delete_cols(col_idx, 1)

    output_path = filename.with_name(f"{filename.stem}_cleaned{filename.suffix}")
    workbook.save(output_path)
    return output_path


def read_csv_with_header_detection(path: Path) -> pd.DataFrame:
    """Read CSV after detecting and trimming rows above the real header."""
    keywords = [keyword.upper() for keyword in HEADER_KEYWORDS]

    with open(path, "r", encoding="utf-8", errors="ignore") as file_obj:
        lines = file_obj.readlines()

    detected_header_index = None
    for line_index, line in enumerate(lines):
        upper_line = line.upper()
        keyword_count = sum(1 for keyword in keywords if keyword in upper_line)
        if keyword_count >= 2:
            detected_header_index = line_index
            break

    if detected_header_index is None:
        raise ValueError(f"Header could not be detected in CSV: {path.name}")

    print(f"  Header detected at CSV line: {detected_header_index + 1}")
    trimmed_lines = lines[detected_header_index:]

    temp_clean = path.with_name(f"{path.stem}_clean.csv")
    try:
        with open(temp_clean, "w", encoding="utf-8") as file_obj:
            file_obj.writelines(trimmed_lines)

        try:
            return pd.read_csv(temp_clean, engine="python", sep=None, dtype=str)
        except Exception:
            return pd.read_csv(temp_clean, engine="python", on_bad_lines="skip", dtype=str)
    finally:
        if temp_clean.exists():
            temp_clean.unlink()


def prepare_excel_dataframe(file_path: Path) -> pd.DataFrame:
    """
    Normalize Excel files before column mapping:
    1) remove duplicate header rows
    2) detect header row
    3) unmerge merged cells
    4) remove empty columns
    5) read data using detected header row
    """
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        working_copy = temp_dir_path / file_path.name
        shutil.copy2(file_path, working_copy)

        df_full = pd.read_excel(working_copy, header=None)
        df_full = remove_duplicate_header_rows(df_full)

        header_index = find_header_row(df_full)
        if header_index is None:
            raise ValueError(f"Header row could not be detected in: {file_path.name}")

        print(f"  Header detected at Excel row: {header_index + 1}")

        no_dup_path = temp_dir_path / "no_duplicate_headers.xlsx"
        df_full.to_excel(no_dup_path, index=False, header=False)

        sheet_name = pd.ExcelFile(no_dup_path).sheet_names[0]
        unmerged_path = temp_dir_path / "unmerged.xlsx"
        sheet_name = unmerge_workbook(no_dup_path, unmerged_path, sheet_name=sheet_name)

        cleaned_path = remove_empty_columns(
            unmerged_path,
            sheet_name,
            header_index + 1,
        )

        df = pd.read_excel(
            cleaned_path,
            sheet_name=sheet_name,
            header=header_index,
            dtype=str,
        )

    df.columns = [str(column).strip() for column in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def find_column(df: pd.DataFrame, mapped_name: str) -> str:
    if mapped_name in df.columns:
        return mapped_name

    normalized_target = normalize_key(mapped_name)
    for column in df.columns:
        if normalize_key(column) == normalized_target:
            return column

    available = ", ".join(str(col) for col in df.columns)
    raise KeyError(
        f"Mapped column '{mapped_name}' not found. Available columns: {available}"
    )


def read_excel_file(file_path: Path) -> pd.DataFrame:
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        df = read_csv_with_header_detection(file_path)
    else:
        df = prepare_excel_dataframe(file_path)

    df.columns = [str(column).strip() for column in df.columns]
    return df.dropna(how="all").reset_index(drop=True)


def validate_dataframe(
    df: pd.DataFrame,
    file_name: str,
    datetime_format: str,
) -> list[str]:
    """
    Scan the entire file before any DB work.
    Returns a list of validation error messages; empty list means file is safe to insert.
    """
    datetime_col = find_column(df, COLUMN_MAPPING["datetime"])
    vehicle_col = find_column(df, COLUMN_MAPPING["vehicle_class"])
    lane_col = find_column(df, COLUMN_MAPPING["lane_no"])
    mop_col = find_column(df, COLUMN_MAPPING["mop"])

    invalid_datetimes: dict[str, int] = {}
    unmapped_vehicle_classes: dict[str, int] = {}
    unmapped_lanes: dict[str, int] = {}
    unmapped_mops: dict[str, int] = {}

    for _row_idx, row in df.iterrows():
        datetime_value = row[datetime_col]
        if not is_blank(datetime_value):
            if try_parse_datetime(datetime_value, datetime_format=datetime_format) is None:
                raw_dt = str(datetime_value).strip()
                invalid_datetimes[raw_dt] = invalid_datetimes.get(raw_dt, 0) + 1

        vehicle_value = row[vehicle_col]
        if not is_blank(vehicle_value):
            if try_normalize_vehicle_class(vehicle_value) is None:
                raw_vc = str(vehicle_value).strip()
                unmapped_vehicle_classes[raw_vc] = unmapped_vehicle_classes.get(raw_vc, 0) + 1

        lane_value = row[lane_col]
        if not is_blank(lane_value):
            if try_normalize_lane(lane_value) is None:
                raw_lane = str(lane_value).strip()
                unmapped_lanes[raw_lane] = unmapped_lanes.get(raw_lane, 0) + 1

        mop_value = row[mop_col]
        if not is_blank(mop_value):
            if try_normalize_mop(mop_value) is None:
                raw_mop = str(mop_value).strip()
                unmapped_mops[raw_mop] = unmapped_mops.get(raw_mop, 0) + 1

    errors: list[str] = []
    if invalid_datetimes:
        errors.append(
            f"Invalid datetime values in '{file_name}' ({datetime_col}): "
            + ", ".join(f"{value!r} ({count} row(s))" for value, count in sorted(invalid_datetimes.items()))
        )
    if unmapped_vehicle_classes:
        errors.append(
            f"Unmapped vehicle class values in '{file_name}' ({vehicle_col}): "
            + ", ".join(f"{value!r} ({count} row(s))" for value, count in sorted(unmapped_vehicle_classes.items()))
        )
    if unmapped_lanes:
        errors.append(
            f"Unmapped lane values in '{file_name}' ({lane_col}): "
            + ", ".join(f"{value!r} ({count} row(s))" for value, count in sorted(unmapped_lanes.items()))
        )
    if unmapped_mops:
        errors.append(
            f"Unmapped MOP values in '{file_name}' ({mop_col}): "
            + ", ".join(f"{value!r} ({count} row(s))" for value, count in sorted(unmapped_mops.items()))
        )

    return errors


def print_validation_errors(errors: list[str]) -> None:
    print("\n" + "=" * 80)
    print("VALIDATION FAILED — abnormal data found. DB insertion skipped.")
    print("=" * 80)
    for error in errors:
        print(f"  - {error}")
    print("=" * 80)


def prepare_dataframe(df: pd.DataFrame, datetime_format: str) -> pd.DataFrame:
    datetime_col = find_column(df, COLUMN_MAPPING["datetime"])
    vehicle_col = find_column(df, COLUMN_MAPPING["vehicle_class"])
    lane_col = find_column(df, COLUMN_MAPPING["lane_no"])
    mop_col = find_column(df, COLUMN_MAPPING["mop"])

    prepared = pd.DataFrame()
    prepared["event_dt"] = df[datetime_col].map(
        lambda value: safe_parse_datetime(value, datetime_format)
    )
    prepared["vehicle_class"] = df[vehicle_col].map(optional_normalize_vehicle_class)
    prepared["lane_no"] = df[lane_col].map(optional_normalize_lane)
    prepared["mop"] = df[mop_col].map(optional_normalize_mop)

    # Only rows with a valid, non-empty date are counted in totals and hourly buckets.
    prepared = prepared[prepared["event_dt"].notna()].copy()
    prepared["hour"] = prepared["event_dt"].map(hour_bucket_label)
    prepared["date"] = prepared["event_dt"].dt.date
    return prepared


def aggregate_hourly_rows(df: pd.DataFrame) -> list[dict]:
    rows: list[dict] = []

    for (record_date, hour_label), group in df.groupby(["date", "hour"], sort=True):
        row = {
            "plaza_name": PLAZA_NAME,
            "hour": hour_label,
            "date": record_date,
        }

        for canonical, column_name in VEHICLE_CLASS_COLUMNS.items():
            row[column_name] = int((group["vehicle_class"] == canonical).sum())

        for lane, column_name in LANE_COLUMNS.items():
            row[column_name] = int((group["lane_no"] == lane).sum())

        for mop_bucket, column_name in MOP_COLUMNS.items():
            row[column_name] = int(group["mop"].eq(mop_bucket).sum())

        row[TOTAL_TRANSACTION_COLUMN] = len(group)
        rows.append(row)

    return rows


def print_hourly_rows(rows: list[dict]) -> None:
    """Print final aggregated rows that would be inserted into the DB."""
    if not rows:
        print("  No hourly rows prepared.")
        return

    print("\n  Prepared DB rows:")
    print("  " + "-" * 76)
    for index, row in enumerate(rows, start=1):
        print(f"  Row {index}:")
        for key, value in row.items():
            print(f"    {key}: {value}")
    print("  " + "-" * 76)


def fetch_existing_hour_keys(conn, table_name: str, rows: list[dict]) -> set[tuple]:
    if not rows:
        return set()

    keys = [(row["plaza_name"], row["date"], row["hour"]) for row in rows]
    query = sql.SQL(
        "SELECT plaza_name, date, hour FROM {table} "
        "WHERE (plaza_name, date, hour) IN %s"
    ).format(table=sql.Identifier(table_name))

    with conn.cursor() as cursor:
        cursor.execute(query.as_string(conn), (tuple(keys),))
        return {(plaza_name, record_date, hour) for plaza_name, record_date, hour in cursor.fetchall()}


def upsert_rows(conn, table_name: str, rows: list[dict]) -> tuple[int, int]:
    if not rows:
        return 0, 0

    data_columns = ["plaza_name", "hour", "date", TOTAL_TRANSACTION_COLUMN]
    data_columns.extend(VEHICLE_CLASS_COLUMNS.values())
    data_columns.extend(LANE_COLUMNS.values())
    data_columns.extend(MOP_COLUMNS.values())

    existing_keys = fetch_existing_hour_keys(conn, table_name, rows)
    rows_to_insert: list[dict] = []
    skipped = 0

    for row in rows:
        key = (row["plaza_name"], row["date"], row["hour"])
        if key in existing_keys:
            print(
                f"  SKIP (exists): plaza={row['plaza_name']}, "
                f"date={row['date']}, hour={row['hour']}"
            )
            skipped += 1
        else:
            rows_to_insert.append(row)

    if not rows_to_insert:
        return 0, skipped

    insert_sql = sql.SQL(
        "INSERT INTO {table} ({fields}) VALUES %s"
    ).format(
        table=sql.Identifier(table_name),
        fields=sql.SQL(", ").join(sql.Identifier(column) for column in data_columns),
    )

    values = [
        tuple(row[column] for column in data_columns)
        for row in rows_to_insert
    ]

    with conn.cursor() as cursor:
        execute_values(cursor, insert_sql.as_string(conn), values)
    conn.commit()
    return len(rows_to_insert), skipped


def list_excel_files(folder_path: Path) -> list[Path]:
    """Return Excel/CSV files under folder_path, including all subfolders."""
    return sorted(
        path
        for path in folder_path.rglob("*")
        if path.is_file()
        and path.suffix.lower() in EXCEL_EXTENSIONS
        and not path.name.startswith("~$")
    )


def process_file(
    file_path: Path,
    table_name: str,
    conn=None,
    dry_run: bool = False,
    run_log: Path | None = None,
) -> tuple[int, bool]:
    print(f"\nProcessing file: {file_path}")
    try:
        df = read_excel_file(file_path)
    except ValueError as exc:
        if is_header_detection_error(exc):
            message = f"[SKIP] {file_path} — {exc}"
            print(f"  {message}")
            if run_log is not None:
                append_run_log(run_log, message)
            return 0, True

        print(f"\n  File preparation failed: {exc}")
        raise UnmappedValueError(
            f"File '{file_path.name}' could not be prepared. No rows were inserted."
        ) from exc

    print(f"  Rows read after header detection and normalization: {len(df)}")

    datetime_col = find_column(df, COLUMN_MAPPING["datetime"])
    datetime_format = detect_datetime_format(df[datetime_col].tolist())
    print(f"  Detected datetime format: {datetime_format}")

    validation_errors = validate_dataframe(df, file_path.name, datetime_format)
    if validation_errors:
        print_validation_errors(validation_errors)
        raise UnmappedValueError(
            f"File '{file_path.name}' failed validation. No rows were inserted."
        )

    print("  Validation passed for entire file.")
    prepared = prepare_dataframe(df, datetime_format)
    hourly_rows = aggregate_hourly_rows(prepared)

    if dry_run:
        print("  DRY RUN — skipping DB insert.")
        print_hourly_rows(hourly_rows)
        return len(hourly_rows), False

    print("  Starting DB insertion...")
    inserted, skipped = upsert_rows(conn, table_name, hourly_rows)
    print(f"  Hourly rows inserted: {inserted}, skipped (already exists): {skipped}")
    return inserted, False


def process_folder(folder_path: str, dry_run: bool = DRY_RUN) -> None:
    folder = Path(folder_path)
    if not folder.is_dir():
        raise FileNotFoundError(f"Input folder not found: {folder_path}")

    table_name = get_analytics_table_name()
    files = list_excel_files(folder)
    if not files:
        print(f"No Excel/CSV files found under: {folder} (including subfolders)")
        return

    print(f"Plaza: {PLAZA_NAME}")
    print(f"Target database: {get_analytics_db_connection_kwargs()['database']}")
    print(f"Target table: {table_name}")
    print(f"Files found: {len(files)}")

    run_log = create_run_log_file()
    print(f"Run log: {run_log}")
    append_run_log(
        run_log,
        f"Plaza: {PLAZA_NAME}\nInput folder: {folder}\nFiles found: {len(files)}",
    )

    if dry_run:
        print("DRY RUN enabled — prepared rows will be printed only.")

    total_rows = 0
    skipped_files = 0
    if dry_run:
        for file_path in files:
            try:
                rows, file_skipped = process_file(
                    file_path,
                    table_name,
                    dry_run=True,
                    run_log=run_log,
                )
                if file_skipped:
                    skipped_files += 1
                total_rows += rows
            except UnmappedValueError:
                append_run_log(run_log, "Execution stopped due to validation failure.")
                print("\nExecution stopped. Fix the abnormal values above and rerun.")
                print(f"See log: {run_log}")
                sys.exit(1)
        append_run_log(
            run_log,
            f"\nDry run complete. Total hourly rows prepared: {total_rows}. "
            f"Skipped files: {skipped_files}.",
        )
        print(f"\nDry run complete. Total hourly rows prepared: {total_rows}")
        print(f"Skipped files: {skipped_files}. Log: {run_log}")
        return

    conn = psycopg2.connect(**get_analytics_db_connection_kwargs())
    try:
        count_columns = [TOTAL_TRANSACTION_COLUMN]
        count_columns.extend(VEHICLE_CLASS_COLUMNS.values())
        count_columns.extend(LANE_COLUMNS.values())
        count_columns.extend(MOP_COLUMNS.values())
        ensure_analytics_table(conn, table_name, count_columns)
        for file_path in files:
            try:
                rows, file_skipped = process_file(
                    file_path,
                    table_name,
                    conn=conn,
                    dry_run=False,
                    run_log=run_log,
                )
                if file_skipped:
                    skipped_files += 1
                total_rows += rows
            except UnmappedValueError:
                append_run_log(run_log, "Execution stopped due to validation failure.")
                print("\nExecution stopped. Fix the abnormal values above and rerun.")
                print(f"See log: {run_log}")
                sys.exit(1)
        append_run_log(
            run_log,
            f"\nCompleted. Total hourly rows inserted: {total_rows}. "
            f"Skipped files: {skipped_files}.",
        )
        print(f"\nCompleted. Total hourly rows inserted: {total_rows}")
        print(f"Skipped files: {skipped_files}. Log: {run_log}")
    finally:
        conn.close()


def main() -> None:
    folder_path = INPUT_FOLDER
    if len(sys.argv) > 1:
        folder_path = sys.argv[1].strip().strip('"')

    process_folder(folder_path, dry_run=DRY_RUN)


if __name__ == "__main__":
    main()
