#!/usr/bin/env python3
"""
File and folder management for the net-net analysis workflow.

Handles:
- Company folder creation
- Template workbook detection and copying
- Downloaded file organization
"""

import re
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional


# Default base directory (parent of netnet_tools)
DEFAULT_BASE_DIR = Path(__file__).parent.parent


def normalize_company_name(name: str) -> str:
    """
    Convert company name to folder-friendly format.

    Examples:
        "Toso Co Ltd" -> "toso"
        "Trilogiq SA" -> "trilogiq"
        "Nansin Co., Ltd." -> "nansin"
        "Car Mate Mfg Co Ltd" -> "carmate"

    Args:
        name: Full company name

    Returns:
        Normalized lowercase name suitable for folder/file names
    """
    # Remove common suffixes
    suffixes = [
        r'\s+Co\.?,?\s*Ltd\.?',
        r'\s+Inc\.?',
        r'\s+Corp\.?',
        r'\s+Corporation',
        r'\s+SA',
        r'\s+AG',
        r'\s+PLC',
        r'\s+Ltd\.?',
        r'\s+Limited',
        r'\s+Mfg',
        r'\s+Manufacturing',
    ]

    result = name
    for suffix in suffixes:
        result = re.sub(suffix, '', result, flags=re.IGNORECASE)

    # Remove special characters and extra spaces
    result = re.sub(r'[^\w\s-]', '', result)
    result = re.sub(r'\s+', '', result)  # Remove all spaces for compact name

    return result.lower()


def create_company_folder(
    company_name: str,
    base_dir: str | Path = None,
) -> Path:
    """
    Create a subfolder for a company.

    Args:
        company_name: Company name (will be normalized)
        base_dir: Base directory. Defaults to fin_editr root.

    Returns:
        Path to the created folder
    """
    base = Path(base_dir) if base_dir else DEFAULT_BASE_DIR

    folder_name = normalize_company_name(company_name)
    folder_path = base / folder_name

    folder_path.mkdir(parents=True, exist_ok=True)
    print(f"Created folder: {folder_path}")

    return folder_path


def find_template_workbook(base_dir: str | Path = None) -> Optional[Path]:
    """
    Find an existing workbook to use as a template.

    Looks for the most recently modified .xlsx file that contains
    calculation sheets (ncav, profitability, etc.).

    Args:
        base_dir: Base directory to search. Defaults to fin_editr root.

    Returns:
        Path to template workbook, or None if not found
    """
    base = Path(base_dir) if base_dir else DEFAULT_BASE_DIR

    # Find all xlsx files in subdirectories
    candidates = []

    for xlsx_file in base.rglob("*.xlsx"):
        # Skip backup files and raw data files
        if "_backup_" in xlsx_file.name:
            continue
        if "Income Statement" in xlsx_file.name:
            continue
        if "Balance Sheet" in xlsx_file.name:
            continue
        if xlsx_file.name.startswith("~"):  # Skip temp files
            continue

        # Check if it's in a company subfolder (not netnet_tools or data)
        if "netnet_tools" in str(xlsx_file):
            continue
        if xlsx_file.parent == base:  # Skip files in root
            continue

        # This is likely a calculations workbook
        candidates.append(xlsx_file)

    if not candidates:
        print("No template workbook found.")
        return None

    # Sort by modification time (most recent first)
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    template = candidates[0]
    print(f"Using template: {template}")
    return template


def copy_template(
    template_path: Path,
    company_name: str,
    target_dir: Path,
) -> Path:
    """
    Copy and rename a template workbook for a new company.

    Args:
        template_path: Path to the template workbook
        company_name: New company name
        target_dir: Target directory for the new workbook

    Returns:
        Path to the new workbook
    """
    normalized_name = normalize_company_name(company_name)
    new_filename = f"{normalized_name}.xlsx"
    new_path = target_dir / new_filename

    shutil.copy2(template_path, new_path)
    print(f"Copied template to: {new_path}")

    return new_path


def move_downloads_to_folder(
    download_paths: list[Path],
    target_dir: Path,
) -> list[Path]:
    """
    Move downloaded files to the company folder.

    Args:
        download_paths: List of downloaded file paths
        target_dir: Target directory

    Returns:
        List of new file paths
    """
    new_paths = []

    for download_path in download_paths:
        if download_path and download_path.exists():
            new_path = target_dir / download_path.name
            shutil.move(str(download_path), str(new_path))
            print(f"Moved: {download_path.name} -> {target_dir}")
            new_paths.append(new_path)

    return new_paths


def extract_company_name_from_export(file_path: Path) -> Optional[str]:
    """
    Extract company name from an Investing.com export filename.

    Examples:
        "Toso Co Ltd - Income Statement.xlsx" -> "Toso Co Ltd"
        "Nansin Co Ltd - Balance Sheet.xlsx" -> "Nansin Co Ltd"

    Args:
        file_path: Path to the export file

    Returns:
        Company name or None if cannot be extracted
    """
    filename = file_path.stem  # Remove .xlsx

    # Pattern: "Company Name - Statement Type"
    match = re.match(r'^(.+?)\s*-\s*(Income Statement|Balance Sheet)', filename)
    if match:
        return match.group(1).strip()

    return None


def find_export_files(folder: Path) -> tuple[Optional[Path], Optional[Path]]:
    """
    Find Income Statement and Balance Sheet export files in a folder.

    Args:
        folder: Folder to search

    Returns:
        Tuple of (income_statement_path, balance_sheet_path)
    """
    is_path = None
    bs_path = None

    for xlsx_file in folder.glob("*.xlsx"):
        name_lower = xlsx_file.name.lower()
        if "income statement" in name_lower:
            is_path = xlsx_file
        elif "balance sheet" in name_lower:
            bs_path = xlsx_file

    return is_path, bs_path


def get_company_folders(base_dir: str | Path = None) -> list[Path]:
    """
    Get list of existing company folders.

    Args:
        base_dir: Base directory. Defaults to fin_editr root.

    Returns:
        List of company folder paths
    """
    base = Path(base_dir) if base_dir else DEFAULT_BASE_DIR

    folders = []
    for item in base.iterdir():
        if not item.is_dir():
            continue
        # Skip non-company folders
        if item.name in ["netnet_tools", "data", ".claude", ".git", "__pycache__"]:
            continue
        if item.name.startswith("."):
            continue
        folders.append(item)

    return sorted(folders)


def cleanup_temp_downloads(download_dir: Path, max_age_hours: int = 24):
    """
    Clean up old temporary download files.

    Args:
        download_dir: Directory containing downloads
        max_age_hours: Delete files older than this
    """
    if not download_dir.exists():
        return

    cutoff = datetime.now().timestamp() - (max_age_hours * 3600)

    for file_path in download_dir.glob("*"):
        if file_path.stat().st_mtime < cutoff:
            try:
                file_path.unlink()
                print(f"Cleaned up: {file_path.name}")
            except Exception:
                pass


def main():
    """Test file manager functions."""
    print("Testing file_manager functions\n")

    # Test normalize_company_name
    test_names = [
        "Toso Co Ltd",
        "Trilogiq SA",
        "Nansin Co., Ltd.",
        "Car Mate Mfg Co Ltd",
        "Almedio Inc",
    ]

    print("normalize_company_name tests:")
    for name in test_names:
        normalized = normalize_company_name(name)
        print(f"  '{name}' -> '{normalized}'")

    print()

    # Test find_template_workbook
    print("Finding template workbook...")
    template = find_template_workbook()
    if template:
        print(f"  Found: {template}")

    print()

    # List existing company folders
    print("Existing company folders:")
    folders = get_company_folders()
    for folder in folders:
        print(f"  {folder.name}")


if __name__ == "__main__":
    main()
