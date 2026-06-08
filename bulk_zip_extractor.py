import os
import zipfile

def extract_zip_files(folder_path):
    # Check if the provided folder path exists
    if not os.path.isdir(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

    # List all files in the folder
    files_in_folder = os.listdir(folder_path)

    # Filter only .zip files
    zip_files = [file for file in files_in_folder if file.endswith('.zip')]

    # Iterate over each .zip file
    for zip_file in zip_files:
        zip_file_path = os.path.join(folder_path, zip_file)

        # Extract folder name from the .zip file name (without extension)
        extracted_folder_name = os.path.splitext(zip_file)[0]
        extracted_folder_path = os.path.join(folder_path, extracted_folder_name)

        # Check if the folder already exists (skip extraction if it does)
        if os.path.isdir(extracted_folder_path):
            print(f"Folder '{extracted_folder_name}' already exists. Skipping extraction of '{zip_file}'.")
            continue
        
        # If the folder doesn't exist, proceed with extraction
        print(f"Extracting '{zip_file}'...")
        
        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(extracted_folder_path)
            print(f"Extraction of '{zip_file}' completed.")
        except zipfile.BadZipFile:
            print(f"Error: '{zip_file}' is not a valid zip file.")
        except Exception as e:
            print(f"Error extracting '{zip_file}': {e}")

if __name__ == "__main__":
    folder_path = input("Enter the path of the folder containing .zip files: ").strip()
    extract_zip_files(folder_path)