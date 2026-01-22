#!/usr/bin/env python3
"""
Net-Net Stock Analysis CLI

Main entry point for automated net-net stock analysis workflow.

Commands:
    analyze         Analyze a stock (download financials, create workbook, run scoring)
    setup-credentials   Configure Investing.com Pro login credentials
    verify-credentials  Test that stored credentials work
    list            List existing company analyses

Examples:
    # One-time credential setup
    python netnet_main.py setup-credentials

    # Analyze a new stock
    python netnet_main.py analyze 7859 --exchange TYO

    # Analyze with price for autoscoring
    python netnet_main.py analyze 7859 --exchange TYO --price 879

    # Use existing downloaded files
    python netnet_main.py analyze 7859 --skip-download --folder ./Toso

    # Show what would be done without making changes
    python netnet_main.py analyze 7859 --exchange TYO --dry-run

    # Show browser window for debugging
    python netnet_main.py analyze 7859 --exchange TYO --no-headless
"""

import argparse
import sys
from pathlib import Path


def cmd_analyze(args):
    """Run the full analysis workflow."""
    from netnet_workflow import run_workflow_sync

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
        print("\n" + "=" * 60)
        print("SUCCESS")
        print("=" * 60)
        print(f"Company: {result.company_name}")
        print(f"Workbook: {result.workbook_path}")
        print("\nNext steps:")
        print("  1. Open the workbook and review the data")
        print("  2. Fill in manual fields in Overview tab (Industry, Price, etc.)")
        print("  3. Run autoscoring with: python netnet_autoscore.py <workbook> --price <price>")
        return 0
    else:
        print("\n" + "=" * 60)
        print("FAILED")
        print("=" * 60)
        print(f"Error: {result.error}")
        if result.steps_completed:
            print(f"\nCompleted steps before failure:")
            for step in result.steps_completed:
                print(f"  - {step}")
        return 1


def cmd_setup_credentials(args):
    """Set up Investing.com credentials."""
    from credentials import setup_credentials
    success = setup_credentials()
    return 0 if success else 1


def cmd_verify_credentials(args):
    """Verify stored credentials."""
    from credentials import verify_credentials
    success = verify_credentials()
    return 0 if success else 1


def cmd_list(args):
    """List existing company analyses."""
    from file_manager import get_company_folders, find_export_files, DEFAULT_BASE_DIR

    folders = get_company_folders()

    if not folders:
        print("No company analyses found.")
        return 0

    print(f"\nExisting analyses in {DEFAULT_BASE_DIR}:\n")
    print(f"{'Folder':<20} {'Workbook':<25} {'Has IS':<8} {'Has BS':<8}")
    print("-" * 65)

    for folder in folders:
        # Check for workbook
        workbooks = list(folder.glob("*.xlsx"))
        workbooks = [w for w in workbooks if "_backup_" not in w.name and not w.name.startswith("~")]

        # Find main workbook (not export files)
        main_workbook = None
        for wb in workbooks:
            if "Income Statement" not in wb.name and "Balance Sheet" not in wb.name:
                main_workbook = wb.name
                break

        # Check for export files
        is_path, bs_path = find_export_files(folder)

        print(f"{folder.name:<20} {main_workbook or '-':<25} {'Yes' if is_path else 'No':<8} {'Yes' if bs_path else 'No':<8}")

    return 0


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Net-Net Stock Analysis CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # analyze command
    analyze_parser = subparsers.add_parser(
        "analyze",
        help="Analyze a stock ticker",
        description="Download financials, create workbook, and run analysis",
    )
    analyze_parser.add_argument("ticker", help="Stock ticker symbol (e.g., 7859, AAPL)")
    analyze_parser.add_argument(
        "--exchange", "-e",
        help="Exchange code (e.g., TYO, NYSE, NASDAQ). Helps disambiguate tickers.",
    )
    analyze_parser.add_argument(
        "--price", "-p",
        type=float,
        help="Current stock price for autoscoring",
    )
    analyze_parser.add_argument(
        "--skip-download",
        action="store_true",
        help="Use existing files instead of downloading",
    )
    analyze_parser.add_argument(
        "--folder", "-f",
        help="Path to folder with existing export files (use with --skip-download)",
    )
    analyze_parser.add_argument(
        "--no-headless",
        action="store_true",
        help="Show browser window (useful for debugging)",
    )
    analyze_parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be done without making changes",
    )
    analyze_parser.set_defaults(func=cmd_analyze)

    # setup-credentials command
    setup_parser = subparsers.add_parser(
        "setup-credentials",
        help="Configure Investing.com Pro login credentials",
    )
    setup_parser.set_defaults(func=cmd_setup_credentials)

    # verify-credentials command
    verify_parser = subparsers.add_parser(
        "verify-credentials",
        help="Test that stored credentials can be retrieved",
    )
    verify_parser.set_defaults(func=cmd_verify_credentials)

    # list command
    list_parser = subparsers.add_parser(
        "list",
        help="List existing company analyses",
    )
    list_parser.set_defaults(func=cmd_list)

    # Parse arguments
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    # Run the command
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
