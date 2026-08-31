import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path


ARCHIVE_EXTENSIONS = {".zip", ".rar"}


def unique_destination(dest_dir: Path, filename: str) -> Path:
    dest_path = dest_dir / filename
    if not dest_path.exists():
        return dest_path

    stem = dest_path.stem
    suffix = dest_path.suffix
    counter = 1
    while True:
        candidate = dest_dir / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def move_files_without_folders(source_dir: Path, dest_dir: Path) -> int:
    """Move every file from source_dir into dest_dir, dropping any folder layers."""
    moved = 0
    for root, _dirs, files in os.walk(source_dir):
        for filename in files:
            src = Path(root) / filename
            dest = unique_destination(dest_dir, filename)
            shutil.move(str(src), str(dest))
            moved += 1
    return moved


def extract_zip_to(archive_path: Path, extract_dir: Path) -> None:
    with zipfile.ZipFile(archive_path, "r") as zip_ref:
        zip_ref.extractall(extract_dir)


def find_seven_zip() -> str | None:
    for candidate in (
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        shutil.which("7z"),
        shutil.which("7z.exe"),
    ):
        if candidate and Path(candidate).exists():
            return candidate
    return None


def extract_rar_to(archive_path: Path, extract_dir: Path) -> None:
    try:
        import rarfile

        with rarfile.RarFile(archive_path) as rar_ref:
            rar_ref.extractall(extract_dir)
        return
    except ImportError:
        pass
    except Exception as exc:
        print(f"  rarfile failed for '{archive_path.name}': {exc}")

    extract_dir.mkdir(parents=True, exist_ok=True)

    try:
        result = subprocess.run(
            ["unrar", "x", "-o+", str(archive_path), str(extract_dir) + os.sep],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode == 0:
            return
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass

    seven_zip = find_seven_zip()
    if seven_zip:
        result = subprocess.run(
            [seven_zip, "x", str(archive_path), f"-o{extract_dir}", "-y"],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode == 0:
            return
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "7-Zip extraction failed")

    raise RuntimeError(
        "Could not extract RAR. Install the rarfile package, unrar, or 7-Zip."
    )


def extract_archive_flat(archive_path: Path, dest_dir: Path) -> int:
    suffix = archive_path.suffix.lower()
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        if suffix == ".zip":
            extract_zip_to(archive_path, tmp_dir)
        elif suffix == ".rar":
            extract_rar_to(archive_path, tmp_dir)
        else:
            raise ValueError(f"Unsupported archive type: {archive_path.name}")
        return move_files_without_folders(tmp_dir, dest_dir)


def list_archive_files(plaza_dir: Path) -> list[Path]:
    return sorted(
        [
            path
            for path in plaza_dir.iterdir()
            if path.is_file() and path.suffix.lower() in ARCHIVE_EXTENSIONS
        ]
    )


def extract_archives_in_plaza_folders(
    main_folder: str,
    delete_source_after_extraction: bool = False,
) -> None:
    main_path = Path(main_folder)
    if not main_path.is_dir():
        print(f"Error: The folder '{main_folder}' does not exist.")
        return

    plaza_dirs = sorted([path for path in main_path.iterdir() if path.is_dir()])
    if not plaza_dirs:
        print(f"No plaza folders found in: {main_path}")
        return

    print(f"Main folder: {main_path}")
    print(f"Found {len(plaza_dirs)} folder(s)")
    print(f"Delete source archives after extraction: {delete_source_after_extraction}\n")

    for plaza_dir in plaza_dirs:
        archives = list_archive_files(plaza_dir)
        print("=" * 60)
        print(f"Plaza folder: {plaza_dir.name}")
        if not archives:
            print("  No zip/rar files found. Skipping.")
            continue

        print(f"  Archives found: {len(archives)}")
        for archive_path in archives:
            print(f"  Extracting '{archive_path.name}' into '{plaza_dir.name}'...")
            try:
                moved = extract_archive_flat(archive_path, plaza_dir)
                print(f"    Done. Placed {moved} file(s) directly in the plaza folder.")
                if delete_source_after_extraction:
                    archive_path.unlink()
                    print(f"    Deleted source file: {archive_path.name}")
            except zipfile.BadZipFile:
                print(f"    Error: '{archive_path.name}' is not a valid zip file.")
            except Exception as exc:
                print(f"    Error extracting '{archive_path.name}': {exc}")


if __name__ == "__main__":
    folder_path = input("Enter the path of the main folder: ").strip().strip('"')
    delete_source_after_extraction = True
    extract_archives_in_plaza_folders(
        folder_path,
        delete_source_after_extraction=delete_source_after_extraction,
    )
