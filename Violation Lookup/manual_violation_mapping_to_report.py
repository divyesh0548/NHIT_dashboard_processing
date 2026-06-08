import pandas as pd
from pathlib import Path
import sys
import re
import glob


def find_column(df, column_names):
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


def map_violation_amt_to_report(first_file_path, second_file_path, output_file_path=None):
    print("="*80)
    print("MANUAL VIOLATION MAPPING TO REPORT")
    print("="*80)
    
    # Validate input files
    if not Path(first_file_path).exists():
        print(f"Error: First file not found: {first_file_path}")
        return False
    
    if not Path(second_file_path).exists():
        print(f"Error: Second file not found: {second_file_path}")
        return False
    
    print(f"\nFirst file (report): {first_file_path}")
    print(f"Second file (source): {second_file_path}")
    
    try:
        # Read first Excel file
        print("\nReading first Excel file...")
        df_first = pd.read_excel(first_file_path)
        
        print(f"  First file: {len(df_first)} rows, {len(df_first.columns)} columns")
        
        # Check if first file already has a 'Violation Amt' column (case-insensitive, handling variations)
        violation_amt_col_first = find_column(df_first, [
            'Violation Amt',
            'Violation Amt.',
            'Violation_Amt',
            'Violation_Amt.',
            'VIOLATION AMT',
            'VIOLATION AMT.',
            'violation_amt',
            'violation amt'
        ])
        
        # Create or use existing 'Violation Amt' column BEFORE checking second file
        if violation_amt_col_first:
            output_col_name = violation_amt_col_first  # Use the exact column name that exists
            print(f"\n✓ Found existing 'Violation Amt' column: '{output_col_name}'")
            print(f"  Will write values to existing column if matches are found.")
        else:
            # Create new column immediately - use 'Violation Amt' as standard name
            output_col_name = 'Violation Amt'
            print(f"\n✓ Column 'Violation Amt' does not exist. Creating new column: '{output_col_name}'")
            # Initialize the column with None values
            df_first[output_col_name] = None
        
        # Find 'concatenated_output' column in first file
        concat_col_first = find_column(df_first, [
            'concatenated_output',
            'Concatenated_Output',
            'Concatenated Output',
            'CONCATENATED_OUTPUT'
        ])
        
        if not concat_col_first:
            print("\nError: Could not find 'concatenated_output' column in first file.")
            print(f"Available columns in first file: {list(df_first.columns)}")
            return False
        
        print(f"\nFound 'concatenated_output' column in first file: '{concat_col_first}'")
        
        # Now read second Excel file
        print("\nReading second Excel file...")
        df_second = pd.read_excel(second_file_path)
        
        print(f"  Second file: {len(df_second)} rows, {len(df_second.columns)} columns")
        
        # Find 'concatenated_output' column in second file
        concat_col_second = find_column(df_second, [
            'concatenated_output',
            'Concatenated_Output',
            'Concatenated Output',
            'CONCATENATED_OUTPUT'
        ])
        
        if not concat_col_second:
            print("\nError: Could not find 'concatenated_output' column in second file.")
            print(f"Available columns in second file: {list(df_second.columns)}")
            return False
        
        print(f"Found 'concatenated_output' column in second file: '{concat_col_second}'")
        
        # Find 'Violation Amt' column in second file
        violation_amt_col = find_column(df_second, [
            'Violation Amt',
            'Violation Amt.',
            'Violation_Amt',
            'Violation_Amt.',
            'VIOLATION AMT',
            'VIOLATION AMT.',
            'violation_amt',
            'violation amt'
        ])
        
        if not violation_amt_col:
            print("\nError: Could not find 'Violation Amt' column in second file.")
            print(f"Available columns in second file: {list(df_second.columns)}")
            return False
        
        print(f"Found 'Violation Amt' column in second file: '{violation_amt_col}'")
        print(f"  Will update values in '{output_col_name}' column.")
        
        # Create a mapping dictionary from second file: concatenated_output -> Violation Amt
        print("\nCreating mapping dictionary from second file...")
        mapping_dict = {}
        matched_count = 0
        skipped_count = 0
        
        for idx, row in df_second.iterrows():
            concatenated = row[concat_col_second]
            violation_amt = row[violation_amt_col]
            
            # Skip rows with empty concatenated_output
            if pd.isna(concatenated) or str(concatenated).strip() == '':
                skipped_count += 1
                continue
            
            concatenated_str = str(concatenated).strip()
            
            # If violation_amt is empty/NaN, set to None
            if pd.isna(violation_amt):
                violation_value = None
            else:
                try:
                    violation_value = float(violation_amt)
                except (ValueError, TypeError):
                    violation_value = None
            
            # Store in mapping (if multiple entries exist, keep the last one)
            mapping_dict[concatenated_str] = violation_value
            matched_count += 1
        
        print(f"  Processed {matched_count} entries from second file")
        print(f"  Skipped {skipped_count} entries with empty concatenated_output")
        print(f"  Unique mappings created: {len(mapping_dict)}")
        
        # Map violation amounts to first file
        print(f"\nMapping violation amounts to first file...")
        
        # If column doesn't exist, it was already created above, so we don't need to initialize again
        # If column exists, we'll update values in it
        
        matches_found = 0
        no_match_count = 0
        
        for idx, row in df_first.iterrows():
            concatenated = row[concat_col_first]
            
            # Skip rows with empty concatenated_output
            if pd.isna(concatenated) or str(concatenated).strip() == '':
                continue
            
            concatenated_str = str(concatenated).strip()
            
            # Check if match exists in mapping dictionary
            if concatenated_str in mapping_dict:
                df_first.at[idx, output_col_name] = mapping_dict[concatenated_str]
                matches_found += 1
            else:
                no_match_count += 1
        
        print(f"  Matches found: {matches_found}")
        print(f"  No match found: {no_match_count}")
        
        # Count non-null values
        non_null_count = df_first[output_col_name].notna().sum()
        print(f"  Non-null values in new column: {non_null_count}")
        
        # Generate output file path if not provided
        if output_file_path is None:
            first_file_path_obj = Path(first_file_path)
            output_file_path = first_file_path_obj.parent / f"{first_file_path_obj.stem}_mapped{first_file_path_obj.suffix}"
        
        # Save the updated first file
        print(f"\nSaving updated file to: {output_file_path}")
        df_first.to_excel(output_file_path, index=False)
        
        print("\n" + "="*80)
        print("MAPPING COMPLETE")
        print("="*80)
        print(f"✓ Updated file saved: {output_file_path}")
        print(f"  Total rows: {len(df_first)}")
        print(f"  Matches found: {matches_found}")
        print(f"  No matches: {no_match_count}")
        print(f"  Non-null violation amounts: {non_null_count}")
        print("="*80)
        
        return True
        
    except Exception as e:
        print(f"\nError during mapping: {e}")
        import traceback
        traceback.print_exc()
        return False


def find_combined_file(folder_path):
    """
    Find Excel file with 'combined' keyword in filename.
    
    Args:
        folder_path: Path to the folder to search
    
    Returns:
        Path: Path to the matching file, or None if not found
    """
    folder = Path(folder_path)
    
    if not folder.exists() or not folder.is_dir():
        return None
    
    # Search for Excel files with 'combined' in filename (case-insensitive)
    for file_path in folder.glob("*.xlsx"):
        if 'combined' in file_path.name.lower():
            return file_path
    
    # Also try .xls files
    for file_path in folder.glob("*.xls"):
        if 'combined' in file_path.name.lower():
            return file_path
    
    return None


def process_folder(folder_path):
    """
    Process a single folder: find the combined file and violation_lookup_output_1.xlsx,
    then map violation amounts.
    
    Args:
        folder_path: Path to the folder to process
    
    Returns:
        bool: True if successful, False otherwise
    """
    folder = Path(folder_path)
    
    if not folder.exists() or not folder.is_dir():
        print(f"Error: Folder not found or not a directory: {folder_path}")
        return False
    
    print(f"\n{'='*80}")
    print(f"Processing folder: {folder.name}")
    print(f"{'='*80}")
    
    # Find the combined file (any Excel file with 'combined' in filename)
    combined_file = find_combined_file(folder)
    
    if not combined_file:
        print(f"  ✗ Could not find Excel file with 'combined' in filename in folder: {folder.name}")
        return False
    
    print(f"  ✓ Found combined file: {combined_file.name}")
    
    # Find violation_lookup_output_1.xlsx
    violation_file = folder / "violation_lookup_output_1.xlsx"
    
    if not violation_file.exists():
        # Try case-insensitive search
        violation_files = list(folder.glob("violation_lookup_output_1.xlsx"))
        if not violation_files:
            violation_files = list(folder.glob("*violation_lookup_output_1*.xlsx"))
        
        if violation_files:
            violation_file = violation_files[0]
        else:
            print(f"  ✗ Could not find 'violation_lookup_output_1.xlsx' in folder: {folder.name}")
            return False
    
    print(f"  ✓ Found violation file: {violation_file.name}")
    
    # Process the mapping (overwrite the combined file)
    print(f"\n  Processing mapping...")
    success = map_violation_amt_to_report(
        first_file_path=str(combined_file),
        second_file_path=str(violation_file),
        output_file_path=str(combined_file)  # Overwrite the original file
    )
    
    if success:
        print(f"  ✓ Successfully updated: {combined_file.name}")
    else:
        print(f"  ✗ Failed to update: {combined_file.name}")
    
    return success


def process_all_folders(root_folder_path):
    """
    Process all subfolders in the root folder.
    Each subfolder should contain:
    - A file matching pattern "{plaza_name}_combined_remaining_dates.xlsx"
    - A file named "violation_lookup_output_1.xlsx"
    
    Args:
        root_folder_path: Path to the root folder containing subfolders
    
    Returns:
        dict: Statistics about processing (total, successful, failed)
    """
    root_folder = Path(root_folder_path)
    
    if not root_folder.exists() or not root_folder.is_dir():
        print(f"Error: Root folder not found or not a directory: {root_folder_path}")
        return {'total': 0, 'successful': 0, 'failed': 0}
    
    print("="*80)
    print("PROCESSING ALL FOLDERS")
    print("="*80)
    print(f"Root folder: {root_folder_path}")
    
    # Get all subfolders
    subfolders = [f for f in root_folder.iterdir() if f.is_dir()]
    
    if not subfolders:
        print(f"\nNo subfolders found in: {root_folder_path}")
        return {'total': 0, 'successful': 0, 'failed': 0}
    
    print(f"\nFound {len(subfolders)} subfolder(s)")
    print("="*80)
    
    stats = {'total': len(subfolders), 'successful': 0, 'failed': 0}
    
    # Process each subfolder
    for idx, subfolder in enumerate(subfolders, 1):
        print(f"\n[{idx}/{len(subfolders)}] Processing: {subfolder.name}")
        
        try:
            success = process_folder(subfolder)
            if success:
                stats['successful'] += 1
            else:
                stats['failed'] += 1
        except Exception as e:
            print(f"  ✗ Error processing folder {subfolder.name}: {e}")
            stats['failed'] += 1
    
    # Print summary
    print("\n" + "="*80)
    print("PROCESSING SUMMARY")
    print("="*80)
    print(f"Total folders processed: {stats['total']}")
    print(f"  ✓ Successful: {stats['successful']}")
    print(f"  ✗ Failed: {stats['failed']}")
    print("="*80)
    
    return stats


def main():
    """Main function to run the script"""
    if len(sys.argv) < 2:
        print("Usage:")
        print("  Mode 1 - Process folder with subfolders:")
        print("    python manual_violation_mapping_to_report.py <folder_path>")
        print("\n  Mode 2 - Process single file pair:")
        print("    python manual_violation_mapping_to_report.py <first_file.xlsx> <second_file.xlsx> [output_file.xlsx]")
        print("\nMode 1 Example:")
        print("  python manual_violation_mapping_to_report.py C:/data/folders")
        print("\nMode 2 Example:")
        print("  python manual_violation_mapping_to_report.py report.xlsx violations.xlsx")
        print("  python manual_violation_mapping_to_report.py report.xlsx violations.xlsx output.xlsx")
        print("\nMode 1 Description:")
        print("  - Processes all subfolders in the given folder")
        print("  - Each subfolder should contain:")
        print("    * An Excel file with 'combined' in filename")
        print("    * A file named: violation_lookup_output_1.xlsx")
        print("  - Updates the combined file with violation amounts from violation_lookup_output_1.xlsx")
        print("\nMode 2 Description:")
        print("  - First file: Report file that will be updated with violation amounts")
        print("  - Second file: Source file containing violation amounts to map")
        print("  - Both files must have 'concatenated_output' column")
        print("  - Second file must have 'Violation Amt' column")
        print("  - Output file: Optional, defaults to <first_file>_mapped.xlsx")
        return
    
    # Check if first argument is a folder (Mode 1) or files (Mode 2)
    first_arg = Path(sys.argv[1])
    
    if first_arg.is_dir() or (len(sys.argv) == 2):
        # Mode 1: Process folder with subfolders
        folder_path = sys.argv[1]
        stats = process_all_folders(folder_path)
        
        if stats['failed'] > 0:
            sys.exit(1)
    else:
        # Mode 2: Process single file pair
        if len(sys.argv) < 3:
            print("Error: For single file mode, both files must be provided.")
            print("Usage: python manual_violation_mapping_to_report.py <first_file.xlsx> <second_file.xlsx> [output_file.xlsx]")
            sys.exit(1)
        
        first_file = sys.argv[1]
        second_file = sys.argv[2]
        output_file = sys.argv[3] if len(sys.argv) > 3 else None
        
        success = map_violation_amt_to_report(first_file, second_file, output_file)
        
        if not success:
            sys.exit(1)


if __name__ == "__main__":
    main()

