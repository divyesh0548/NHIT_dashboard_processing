from violation_extractor import extract_and_combine_excel_files
from update_exceptions import clear_exceptions, remove_duplicates, update_exceptions_in_db
from Violation_lookup import process_single_excel_file, process_folder, match_and_generate_output
from exception_value_count import clear_exception_values, calculate_exception_values
from update_manual_violation_from_excel import clear_manual_violation, process_excel_and_update_db
from pathlib import Path


def extract_violation_report(plaza_name: str, from_date: str, to_date: str, output_file: str = None):
    print(f"\n{'='*80}")
    print("VIOLATION REPORT EXTRACTION")
    print(f"{'='*80}")
    print(f"Plaza Name: {plaza_name}")
    print(f"Date Range: {from_date} to {to_date}")

    # Generate output filename if not provided
    if not output_file:
        safe_plaza = plaza_name.replace(' ', '_').replace('/', '_')
        output_file = f"{safe_plaza}_combined_{from_date}_to_{to_date}.xlsx"
        print(f"Output File: {output_file}")

    print(f"{'='*80}\n")

    # Extract and combine Excel files using date range (not specific_dates)
    output_path = extract_and_combine_excel_files(
        entity_name=plaza_name,
        from_date=from_date,
        to_date=to_date,
        output_file=output_file,
        specific_dates=None  # Exclude specific dates logic, use date range only
    )

    if output_path:
        print(f"\n{'='*80}")
        print("EXTRACTION COMPLETE")
        print(f"{'='*80}")
        print(f"Output file: {output_path}")
        print(f"{'='*80}")
        return output_path
    else:
        print(f"\n{'='*80}")
        print("EXTRACTION FAILED")
        print(f"{'='*80}")
        return None


def process_violation_lookup(excel_file_path: str, folder_path: str):
    print(f"\n{'='*80}")
    print("VIOLATION LOOKUP PROCESSING")
    print(f"{'='*80}")
    print(f"Excel File: {excel_file_path}")
    print(f"Folder: {folder_path}")
    print(f"{'='*80}")

    # Check if Excel file exists
    if not Path(excel_file_path).exists():
        print(f"\nError: Excel file not found: {excel_file_path}")
        return False

    # Check if folder exists
    if not Path(folder_path).exists():
        print(f"\nError: Folder not found: {folder_path}")
        return False

    # Process single Excel file
    print(f"\n{'='*80}")
    print("STEP 1: Processing single Excel file")
    print(f"{'='*80}")
    single_file_results, df_single = process_single_excel_file(excel_file_path)

    if not single_file_results or df_single is None:
        print("Error: Could not process single Excel file.")
        return False

    # Process all Excel files in the folder
    print(f"\n{'='*80}")
    print("STEP 2: Processing Excel files in folder")
    print(f"{'='*80}")
    folder_results, violation_amts, reader_read_times, journey_types = process_folder(folder_path)

    if not folder_results:
        print("Warning: No results from folder files. Skipping comparison.")
        return False

    # Match, mark violation_matched, and generate output file(s)
    print(f"\n{'='*80}")
    print("STEP 3: Matching & generating output")
    print(f"{'='*80}")
    print("Flow: Compare single file rows (CCH TXN NO) with folder data (Agency Txn Id)")
    print("  -> For ALL matches: Write 'Journey Type', 'Reader Read Time', and 'Violation Amt' to single Excel file")
    print("  -> If Violation Amt is non-zero -> mark 'Yes' in violation_matched + write to output")
    print("  -> Otherwise -> mark 'No' in violation_matched")
    print("  -> Output file(s) contain ONLY matched entries with non-zero Violation Amt")

    success = match_and_generate_output(
        excel_file_path, df_single, folder_path,
        folder_results, violation_amts, reader_read_times, journey_types
    )

    if success:
        print(f"\n{'='*80}")
        print("VIOLATION LOOKUP COMPLETE")
        print(f"{'='*80}")
        print(f"Updated Excel file: {excel_file_path}")
        print(f"{'='*80}")
        return True
    else:
        print(f"\n{'='*80}")
        print("VIOLATION LOOKUP FAILED")
        print(f"{'='*80}")
        return False


if __name__ == "__main__":
    # ============================================================================
    # CONFIGURATION
    # ============================================================================
    plaza_name = "madai"
    from_date = "2026-03-30"
    to_date = "2026-04-13"
    folder_path = "Life_cycle_files_Mar30_to_Apr13/Madai"  # Folder containing Excel files for violation lookup comparison
    # Optional: specify custom output file path
    # output_file = "custom_output.xlsx"
    output_file = None  # Will auto-generate if None
    # ============================================================================

    # Step 1: Extract violation report
    output_path = extract_violation_report(plaza_name, from_date, to_date, output_file)

    if output_path:
        # Step 2: Clear old exceptions values
        clear_exceptions(plaza_name, from_date, to_date)

        # Step 3: Remove duplicates based on 'CCH TXN NO'
        df = remove_duplicates(output_path)

        # Step 4: Update exceptions in DB for each (date, shift) combination
        update_exceptions_in_db(df, plaza_name)

        # Step 5: Process violation lookup (compare with folder files and update original file)
        violation_lookup_success = process_violation_lookup(output_path, folder_path)

        if violation_lookup_success:
            # Step 6: Clear old exception values and calculate new ones
            clear_exception_values(plaza_name, from_date, to_date)
            calculate_exception_values(output_path, plaza_name)

            # Step 7: Clear old manual violation values and update from input file
            print(f"\n{'='*80}")
            print("STEP 7: Updating Manual Violation from Input File")
            print(f"{'='*80}")
            
            # Clear old manual violation values first
            clear_result = clear_manual_violation(plaza_name, from_date, to_date)
            
            if clear_result.get('errors', 0) > 0:
                print("\n✗ Error during manual violation cleanup. Skipping update.")
            else:
                print("\n✓ Cleanup complete. Proceeding to update...\n")
                
                # Process the same input file (output_path) that was generated in step 1
                print(f"Processing input file: {output_path}")
                stats = process_excel_and_update_db(output_path, plaza_name)
                if stats:
                    print(f"  Updated: {stats.get('updated', 0)}, Skipped: {stats.get('skipped', 0)}")

        print(f"\n{'='*80}")
        print("AUTOMATION COMPLETE")
        print(f"{'='*80}")
    else:
        print(f"\n{'='*80}")
        print("AUTOMATION FAILED - Could not extract violation report")
        print(f"{'='*80}")

