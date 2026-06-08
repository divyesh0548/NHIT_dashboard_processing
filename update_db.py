from extraction import process_folder_excel_files, read_processed_excel_and_update_db

# Configuration
FOLDER_PATH = "test"  # Folder containing Excel files to process
# PLAZA_NAME = "dahalapara"  # entity_name in database
OUTPUT_FILE = "combined_output.xlsx"  # Output file name

print("=" * 60)
print("Step 1: Processing all Excel files in folder")
print("=" * 60)

# Step 1: Process all Excel files in the folder and generate combined output
combined_totals = process_folder_excel_files(
    folder_path=FOLDER_PATH,
    output_file_path=OUTPUT_FILE
)

if combined_totals is None:
    print("\nError: Failed to process Excel files. Exiting...")
    exit(1)

print("\n" + "=" * 60)
print("Step 2: Updating database with combined results")
print("=" * 60)

# Step 2: Read the combined output file and update database
# stats = read_processed_excel_and_update_db(
#     input_file_path=OUTPUT_FILE,
#     plaza_name=PLAZA_NAME
# )

# Print final summary
print("\n" + "=" * 60)
print("Final Summary:")
print("=" * 60)
# if stats['updated'] > 0:
#     print(f"  ✓ Updated: {stats['updated']}")
#     print(f"  ⊘ Skipped (already exists): {stats.get('already_exists', 0)}")
#     print(f"  ⊘ Skipped (no valid rows): {stats.get('no_valid_rows', 0)}")
#     print(f"  ⊘ Skipped (other): {stats.get('skipped', 0)}")
# else:
#     print("  No updates were made.")
#     print(f"  Skipped (already exists): {stats.get('already_exists', 0)}")
#     print(f"  Skipped (no valid rows): {stats.get('no_valid_rows', 0)}")
#     print(f"  Skipped (other): {stats.get('skipped', 0)}")
# print("=" * 60)