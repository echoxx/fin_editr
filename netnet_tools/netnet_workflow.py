#!/usr/bin/env python3
"""
Workflow orchestrator for automated net-net stock analysis.

Coordinates the full workflow:
1. Login to Investing.com Pro
2. Search and download financial statements
3. Create company folder and workbook from template
4. Update workbook with financial data
5. Populate Overview tab
6. Optionally run autoscoring
"""

import asyncio
import sys
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

# Local imports
from credentials import get_credentials, credentials_exist
from file_manager import (
    create_company_folder,
    find_template_workbook,
    copy_template,
    move_downloads_to_folder,
    extract_company_name_from_export,
    find_export_files,
    normalize_company_name,
    DEFAULT_BASE_DIR,
)
from overview_populator import populate_overview, clear_overview_manual_fields
from investing_scraper import InvestingComScraper, DownloadResult, PLAYWRIGHT_AVAILABLE


@dataclass
class WorkflowResult:
    """Result of the complete workflow."""
    success: bool = False
    company_name: Optional[str] = None
    folder_path: Optional[Path] = None
    workbook_path: Optional[Path] = None
    income_statement_path: Optional[Path] = None
    balance_sheet_path: Optional[Path] = None
    error: Optional[str] = None
    steps_completed: list[str] = field(default_factory=list)


async def run_full_workflow(
    ticker: str,
    exchange: str = None,
    price: float = None,
    headless: bool = True,
    skip_download: bool = False,
    existing_folder: str = None,
    dry_run: bool = False,
    base_dir: str = None,
) -> WorkflowResult:
    """
    Run the complete analysis workflow for a ticker.

    Args:
        ticker: Stock ticker symbol
        exchange: Exchange code (e.g., "TYO", "NYSE")
        price: Stock price for autoscoring (optional)
        headless: Run browser in headless mode
        skip_download: Use existing files instead of downloading
        existing_folder: Path to folder with existing export files
        dry_run: Show what would be done without making changes
        base_dir: Base directory for company folders

    Returns:
        WorkflowResult with status and paths
    """
    result = WorkflowResult()
    base = Path(base_dir) if base_dir else DEFAULT_BASE_DIR

    try:
        # ================================================================
        # STEP 1: Get financial statements (download or use existing)
        # ================================================================

        if skip_download:
            # Use existing files
            print("\n" + "=" * 60)
            print("STEP 1: Using existing files")
            print("=" * 60)

            if not existing_folder:
                result.error = "Must specify --folder when using --skip-download"
                return result

            folder_path = Path(existing_folder)
            if not folder_path.is_absolute():
                # Try relative to current working directory first
                cwd_path = Path.cwd() / folder_path
                if cwd_path.exists():
                    folder_path = cwd_path.resolve()
                else:
                    # Fall back to relative to base directory
                    folder_path = (base / folder_path).resolve()

            if not folder_path.exists():
                result.error = f"Folder not found: {folder_path}"
                return result

            is_path, bs_path = find_export_files(folder_path)

            if not is_path or not bs_path:
                result.error = f"Could not find Income Statement and Balance Sheet files in {folder_path}"
                return result

            result.income_statement_path = is_path
            result.balance_sheet_path = bs_path

            # Extract company name from file
            company_name = extract_company_name_from_export(is_path)
            if not company_name:
                company_name = extract_company_name_from_export(bs_path)
            if not company_name:
                company_name = folder_path.name

            result.company_name = company_name
            result.folder_path = folder_path
            result.steps_completed.append("Found existing export files")

            print(f"  Income Statement: {is_path.name}")
            print(f"  Balance Sheet: {bs_path.name}")
            print(f"  Company: {company_name}")

        else:
            # Download from Investing.com
            print("\n" + "=" * 60)
            print("STEP 1: Downloading from Investing.com")
            print("=" * 60)

            if not PLAYWRIGHT_AVAILABLE:
                result.error = (
                    "Playwright not installed. Install with:\n"
                    "  pip install playwright\n"
                    "  playwright install chromium"
                )
                return result

            # Check credentials
            if not credentials_exist():
                result.error = (
                    "No credentials configured. Run:\n"
                    "  python netnet_main.py setup-credentials"
                )
                return result

            email, password = get_credentials()

            # Create temp download directory
            temp_download_dir = base / ".downloads"
            temp_download_dir.mkdir(exist_ok=True)

            async with InvestingComScraper(
                download_dir=temp_download_dir,
                headless=headless,
            ) as scraper:
                # Login
                print("\nLogging in to Investing.com...")
                await scraper.login(email, password)
                result.steps_completed.append("Logged in to Investing.com")

                # Download financials
                print(f"\nSearching for: {ticker}" + (f" on {exchange}" if exchange else ""))
                download_result: DownloadResult = await scraper.download_financials(
                    ticker, exchange
                )

                if download_result.error:
                    result.error = download_result.error
                    return result

                if not download_result.income_statement_path or not download_result.balance_sheet_path:
                    result.error = "Failed to download one or both financial statements"
                    return result

                result.income_statement_path = download_result.income_statement_path
                result.balance_sheet_path = download_result.balance_sheet_path
                result.company_name = download_result.company_info.name if download_result.company_info else None
                result.steps_completed.append("Downloaded financial statements")

                print(f"  Downloaded: {download_result.income_statement_path.name}")
                print(f"  Downloaded: {download_result.balance_sheet_path.name}")

            # Extract company name from downloaded files if not already set
            if not result.company_name:
                result.company_name = extract_company_name_from_export(result.income_statement_path)

        if not result.company_name:
            result.error = "Could not determine company name"
            return result

        # ================================================================
        # STEP 2: Create company folder and copy template
        # ================================================================

        print("\n" + "=" * 60)
        print("STEP 2: Setting up company folder")
        print("=" * 60)

        if not skip_download:
            # Create folder for new company
            folder_path = create_company_folder(result.company_name, base)
            result.folder_path = folder_path

            # Move downloaded files to company folder
            moved_paths = move_downloads_to_folder(
                [result.income_statement_path, result.balance_sheet_path],
                folder_path,
            )
            result.income_statement_path = moved_paths[0] if len(moved_paths) > 0 else None
            result.balance_sheet_path = moved_paths[1] if len(moved_paths) > 1 else None

            result.steps_completed.append("Created company folder")
        else:
            folder_path = result.folder_path

        # Find and copy template workbook
        normalized_name = normalize_company_name(result.company_name)
        expected_workbook = folder_path / f"{normalized_name}.xlsx"

        if expected_workbook.exists():
            print(f"  Workbook already exists: {expected_workbook.name}")
            result.workbook_path = expected_workbook
        else:
            template = find_template_workbook(base)
            if not template:
                result.error = "No template workbook found. Create one manually first."
                return result

            if dry_run:
                print(f"  [DRY RUN] Would copy template: {template.name}")
                result.workbook_path = expected_workbook
            else:
                result.workbook_path = copy_template(template, result.company_name, folder_path)
                result.steps_completed.append("Copied template workbook")

        # ================================================================
        # STEP 3: Update workbook with financial data
        # ================================================================

        print("\n" + "=" * 60)
        print("STEP 3: Updating workbook with financial data")
        print("=" * 60)

        if dry_run:
            print("  [DRY RUN] Would update workbook with new data")
        else:
            # Import the updater (avoid circular imports)
            from netnet_updater import update_workbook

            update_result = update_workbook(
                workbook_path=str(result.workbook_path),
                is_export_path=str(result.income_statement_path),
                bs_export_path=str(result.balance_sheet_path),
                replace_all=True,  # Always replace all for new companies
                dry_run=False,
            )

            if not update_result.success and update_result.errors:
                result.error = "; ".join(update_result.errors)
                return result

            print(f"  Sheets renamed: {update_result.sheets_renamed}")
            print(f"  Formula references updated: {update_result.formula_refs_updated}")
            print(f"  IS cells updated: {update_result.is_cells_updated}")
            print(f"  BS cells updated: {update_result.bs_cells_updated}")
            result.steps_completed.append("Updated workbook with financial data")

        # ================================================================
        # STEP 4: Populate Overview tab
        # ================================================================

        print("\n" + "=" * 60)
        print("STEP 4: Populating Overview tab")
        print("=" * 60)

        if dry_run:
            print("  [DRY RUN] Would populate Overview tab")
        else:
            # First clear manual fields (in case template had old data)
            clear_overview_manual_fields(result.workbook_path)

            # Populate with new data
            populated = populate_overview(
                result.workbook_path,
                ticker=ticker,
                exchange=exchange,
            )

            if populated:
                result.steps_completed.append("Populated Overview tab")

        # ================================================================
        # STEP 5: Run autoscoring (optional)
        # ================================================================

        if price is not None:
            print("\n" + "=" * 60)
            print("STEP 5: Running autoscoring")
            print("=" * 60)

            if dry_run:
                print(f"  [DRY RUN] Would run autoscoring with price={price}")
            else:
                try:
                    from netnet_autoscore import score_workbook

                    score_workbook(str(result.workbook_path), price)
                    result.steps_completed.append("Completed autoscoring")
                except ImportError:
                    print("  Warning: Could not import autoscore module")
                except Exception as e:
                    print(f"  Warning: Autoscoring failed: {e}")

        # ================================================================
        # DONE
        # ================================================================

        result.success = True

        print("\n" + "=" * 60)
        print("WORKFLOW COMPLETE")
        print("=" * 60)
        print(f"  Company: {result.company_name}")
        print(f"  Folder: {result.folder_path}")
        print(f"  Workbook: {result.workbook_path}")
        print(f"  Steps completed: {len(result.steps_completed)}")

        return result

    except Exception as e:
        result.error = str(e)
        return result


def run_workflow_sync(
    ticker: str,
    exchange: str = None,
    price: float = None,
    headless: bool = True,
    skip_download: bool = False,
    existing_folder: str = None,
    dry_run: bool = False,
    base_dir: str = None,
) -> WorkflowResult:
    """
    Synchronous wrapper for run_full_workflow.

    Same arguments as run_full_workflow.
    """
    return asyncio.run(run_full_workflow(
        ticker=ticker,
        exchange=exchange,
        price=price,
        headless=headless,
        skip_download=skip_download,
        existing_folder=existing_folder,
        dry_run=dry_run,
        base_dir=base_dir,
    ))


def main():
    """CLI for testing the workflow."""
    import argparse

    parser = argparse.ArgumentParser(description="Run net-net analysis workflow")
    parser.add_argument("ticker", help="Stock ticker symbol")
    parser.add_argument("--exchange", "-e", help="Exchange code (e.g., TYO)")
    parser.add_argument("--price", type=float, help="Stock price for autoscoring")
    parser.add_argument("--no-headless", action="store_true", help="Show browser window")
    parser.add_argument("--skip-download", action="store_true", help="Use existing files")
    parser.add_argument("--folder", help="Folder with existing export files")
    parser.add_argument("--dry-run", action="store_true", help="Show what would be done")

    args = parser.parse_args()

    result = run_workflow_sync(
        ticker=args.ticker,
        exchange=args.exchange,
        price=args.price,
        headless=not args.no_headless,
        skip_download=args.skip_download,
        existing_folder=args.folder,
        dry_run=args.dry_run,
    )

    if result.success:
        print("\nWorkflow completed successfully!")
        return 0
    else:
        print(f"\nWorkflow failed: {result.error}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
