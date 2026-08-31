"""
Aggregate plaza Excel transaction data into hourly analytics rows and upsert into PostgreSQL.

Configure PLAZA_NAME, COLUMN_MAPPING, normalization lists, and INPUT_FOLDER below.
"""

from __future__ import annotations

import os
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
sys.path.insert(0, str(ROOT_DIR))

from db_config import get_db_connection_kwargs, load_env_file


# ---------------------------------------------------------------------------
# Configuration — update for each plaza / Excel layout
# ---------------------------------------------------------------------------

PLAZA_NAME = "aroli"

INPUT_FOLDER = r"C:\Divyesh\NHIT_dashboard_processing\New Dashboard Insights\Input"

# Excel column name -> logical field
COLUMN_MAPPING = {
    "datetime": "DATE",
    "vehicle_class": "MVC",
    "lane_no": "LANE",
    "mop": "MVC MOP",
}

# Canonical vehicle class -> list of raw Excel strings that map to it
VEHICLE_CLASS_NORMALIZATION: dict[str, list[str]] = {
    "3WHEELER": ["3WHEELER", "3 WHEELER"],
    "BUS-2-AXLE": ["BUS-2-AXLE", "BUS 2 AXLE"],
    "CAR/JEEP": ["CAR/JEEP", "CAR", "JEEP"],
    "LCV": ["LCV"],
    "TRACTOR": ["TRACTOR"],
    "TRUCK 4-6 AXLE": ["TRUCK 4-6 AXLE", "TRUCK 4 6 AXLE"],
    "TRUCK-2-AXLE": ["TRUCK-2-AXLE", "TRUCK 2 AXLE"],
    "TRUCK-3-AXLE": ["TRUCK-3-AXLE", "TRUCK 3 AXLE"],
}

# Canonical MOP bucket -> list of raw Excel strings that map to it
MOP_NORMALIZATION: dict[str, list[str]] = {
    "exempt": ["EXEMPT", "EXEMPTION"],
    "cash": ["CASH"],
    "fastag": ["FASTAG", "FAST TAG", "NETC FASTAG"],
}

LANE_NUMBERS = tuple(f"L{i:02d}" for i in range(1, 13))

# Add lane aliases here when you define plaza-specific lane parsing rules.
LANE_NORMALIZATION: dict[str, list[str]] = {
    lane: [lane] for lane in LANE_NUMBERS
}

VEHICLE_CLASS_COLUMNS = {
    "3WHEELER": "vc_3wheeler",
    "BUS-2-AXLE": "vc_bus_2_axle",
    "CAR/JEEP": "vc_car_jeep",
    "LCV": "vc_lcv",
    "TRACTOR": "vc_tractor",
    "TRUCK 4-6 AXLE": "vc_truck_4_6_axle",
    "TRUCK-2-AXLE": "vc_truck_2_axle",
    "TRUCK-3-AXLE": "vc_truck_3_axle",
}

LANE_COLUMNS = {lane: lane.lower() for lane in LANE_NUMBERS}
MOP_COLUMNS = {"exempt": "mop_exempt", "cash": "mop_cash", "fastag": "mop_fastag"}

EXCEL_EXTENSIONS = {".xlsx", ".xls", ".csv"}

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


def get_analytics_table_name() -> str:
    load_env_file()
    table_name = os.environ.get("Analytical_DB_Table", "").strip()
    if not table_name:
        raise RuntimeError(
            "Missing Analytical_DB_Table in project-root .env "
            "(e.g. Analytical_DB_Table=nhit_analytics)."
        )
    return table_name


def normalize_key(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


def build_lookup(normalization: dict[str, list[str]]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for canonical, aliases in normalization.items():
        lookup[normalize_key(canonical)] = canonical
        for alias in aliases:
            lookup[normalize_key(alias)] = canonical
    return lookup


VEHICLE_CLASS_LOOKUP = build_lookup(VEHICLE_CLASS_NORMALIZATION)
MOP_LOOKUP = build_lookup(MOP_NORMALIZATION)
LANE_LOOKUP = build_lookup(LANE_NORMALIZATION)


def try_parse_datetime(value):
    try:
        return parse_datetime(value)
    except (ValueError, TypeError):
        return None


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


def parse_datetime(value) -> datetime:
    if pd.isna(value):
        raise ValueError("Empty datetime value")
    if isinstance(value, datetime):
        return value
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    text = str(value).strip()
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        raise ValueError(f"Unable to parse datetime: {value!r}")
    return parsed.to_pydatetime()


def hour_bucket_label(dt: datetime) -> str:
    return f"{dt.hour}-{dt.hour + 1}"


def find_header_row(df: pd.DataFrame) -> int | None:
    """Find the header row by matching known header keywords."""
    keywords = [keyword.upper() for keyword in HEADER_KEYWORDS]

    for row_index in range(len(df)):
        row_values = df.iloc[row_index].astype(str).values
        row_text = " ".join(row_values).upper()
        keyword_count = sum(1 for keyword in keywords if keyword in row_text)
        if keyword_count >= 2:
            return row_index

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


def validate_dataframe(df: pd.DataFrame, file_name: str) -> list[str]:
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
        if try_parse_datetime(row[datetime_col]) is None:
            raw_dt = "" if pd.isna(row[datetime_col]) else str(row[datetime_col]).strip()
            invalid_datetimes[raw_dt or "<empty>"] = invalid_datetimes.get(raw_dt or "<empty>", 0) + 1

        if try_normalize_vehicle_class(row[vehicle_col]) is None:
            raw_vc = "" if pd.isna(row[vehicle_col]) else str(row[vehicle_col]).strip()
            unmapped_vehicle_classes[raw_vc or "<empty>"] = (
                unmapped_vehicle_classes.get(raw_vc or "<empty>", 0) + 1
            )

        if try_normalize_lane(row[lane_col]) is None:
            raw_lane = "" if pd.isna(row[lane_col]) else str(row[lane_col]).strip()
            unmapped_lanes[raw_lane or "<empty>"] = unmapped_lanes.get(raw_lane or "<empty>", 0) + 1

        if try_normalize_mop(row[mop_col]) is None:
            raw_mop = "" if pd.isna(row[mop_col]) else str(row[mop_col]).strip()
            unmapped_mops[raw_mop or "<empty>"] = unmapped_mops.get(raw_mop or "<empty>", 0) + 1

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


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    datetime_col = find_column(df, COLUMN_MAPPING["datetime"])
    vehicle_col = find_column(df, COLUMN_MAPPING["vehicle_class"])
    lane_col = find_column(df, COLUMN_MAPPING["lane_no"])
    mop_col = find_column(df, COLUMN_MAPPING["mop"])

    prepared = pd.DataFrame()
    prepared["event_dt"] = df[datetime_col].map(parse_datetime)
    prepared["vehicle_class"] = df[vehicle_col].map(normalize_vehicle_class)
    prepared["lane_no"] = df[lane_col].map(normalize_lane)
    prepared["mop"] = df[mop_col].map(normalize_mop)
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
            row[column_name] = int((group["mop"] == mop_bucket).sum())

        rows.append(row)

    return rows


def create_table_if_not_exists(conn, table_name: str) -> None:
    count_columns = list(VEHICLE_CLASS_COLUMNS.values())
    count_columns.extend(LANE_COLUMNS.values())
    count_columns.extend(MOP_COLUMNS.values())

    column_defs = [
        "id BIGSERIAL PRIMARY KEY",
        "plaza_name TEXT NOT NULL",
        "hour TEXT NOT NULL",
        "date DATE NOT NULL",
    ]
    column_defs.extend(f"{column_name} INTEGER NOT NULL DEFAULT 0" for column_name in count_columns)
    column_defs.append("UNIQUE (plaza_name, date, hour)")

    create_sql = sql.SQL(
        "CREATE TABLE IF NOT EXISTS {table} ({columns});"
    ).format(
        table=sql.Identifier(table_name),
        columns=sql.SQL(", ").join(sql.SQL(part) for part in column_defs),
    )

    with conn.cursor() as cursor:
        cursor.execute(create_sql)
    conn.commit()


def upsert_rows(conn, table_name: str, rows: list[dict]) -> int:
    if not rows:
        return 0

    data_columns = ["plaza_name", "hour", "date"]
    data_columns.extend(VEHICLE_CLASS_COLUMNS.values())
    data_columns.extend(LANE_COLUMNS.values())
    data_columns.extend(MOP_COLUMNS.values())

    update_columns = [
        column
        for column in data_columns
        if column not in {"plaza_name", "hour", "date"}
    ]

    insert_sql = sql.SQL(
        "INSERT INTO {table} ({fields}) VALUES %s "
        "ON CONFLICT (plaza_name, date, hour) DO UPDATE SET {updates}"
    ).format(
        table=sql.Identifier(table_name),
        fields=sql.SQL(", ").join(sql.Identifier(column) for column in data_columns),
        updates=sql.SQL(", ").join(
            sql.SQL("{column} = EXCLUDED.{column}").format(column=sql.Identifier(column))
            for column in update_columns
        ),
    )

    values = [
        tuple(row[column] for column in data_columns)
        for row in rows
    ]

    with conn.cursor() as cursor:
        execute_values(cursor, insert_sql.as_string(conn), values)
    conn.commit()
    return len(rows)


def list_excel_files(folder_path: Path) -> list[Path]:
    return sorted(
        path
        for path in folder_path.iterdir()
        if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS
    )


def process_file(conn, table_name: str, file_path: Path) -> int:
    print(f"\nProcessing file: {file_path.name}")
    try:
        df = read_excel_file(file_path)
    except ValueError as exc:
        print(f"\n  File preparation failed: {exc}")
        raise UnmappedValueError(
            f"File '{file_path.name}' could not be prepared. No rows were inserted."
        ) from exc

    print(f"  Rows read after header detection and normalization: {len(df)}")

    validation_errors = validate_dataframe(df, file_path.name)
    if validation_errors:
        print_validation_errors(validation_errors)
        raise UnmappedValueError(
            f"File '{file_path.name}' failed validation. No rows were inserted."
        )

    print("  Validation passed for entire file. Starting DB insertion...")
    prepared = prepare_dataframe(df)
    hourly_rows = aggregate_hourly_rows(prepared)
    inserted = upsert_rows(conn, table_name, hourly_rows)

    print(f"  Hourly rows upserted: {inserted}")
    return inserted


def process_folder(folder_path: str) -> None:
    folder = Path(folder_path)
    if not folder.is_dir():
        raise FileNotFoundError(f"Input folder not found: {folder_path}")

    table_name = get_analytics_table_name()
    files = list_excel_files(folder)
    if not files:
        print(f"No Excel/CSV files found in: {folder}")
        return

    print(f"Plaza: {PLAZA_NAME}")
    print(f"Target table: {table_name}")
    print(f"Files found: {len(files)}")

    conn = psycopg2.connect(**get_db_connection_kwargs())
    try:
        create_table_if_not_exists(conn, table_name)
        total_rows = 0
        for file_path in files:
            try:
                total_rows += process_file(conn, table_name, file_path)
            except UnmappedValueError:
                print("\nExecution stopped. Fix the abnormal values above and rerun.")
                sys.exit(1)
        print(f"\nCompleted. Total hourly rows upserted: {total_rows}")
    finally:
        conn.close()


def main() -> None:
    folder_path = INPUT_FOLDER
    if len(sys.argv) > 1:
        folder_path = sys.argv[1].strip().strip('"')

    process_folder(folder_path)


if __name__ == "__main__":
    main()
