"""CLI entry point for the FBDI comparison engine."""

import argparse
import logging
import sys
from pathlib import Path

from fbdi.compare import compare_all
from fbdi.utils import match_fbdi_files


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="fbdi",
        description="Oracle FBDI template comparison engine",
    )
    subparsers = parser.add_subparsers(dest="command")

    compare_parser = subparsers.add_parser(
        "compare",
        help="Compare FBDI templates between two release versions",
    )
    compare_parser.add_argument(
        "--old", required=True, type=Path,
        help="Path to directory containing old FBDI templates",
    )
    compare_parser.add_argument(
        "--new", required=True, type=Path,
        help="Path to directory containing new FBDI templates",
    )
    compare_parser.add_argument(
        "--output", type=Path, default=Path("Comparison_Report.xlsx"),
        help="Output file path (default: Comparison_Report.xlsx)",
    )
    compare_parser.add_argument(
        "--all-rows", action="store_true",
        help="Include unchanged rows in output (default: changes only)",
    )
    compare_parser.add_argument(
        "--verbose", action="store_true",
        help="Set logging to DEBUG (shows header detection scores)",
    )

    args = parser.parse_args(argv)

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "compare":
        _run_compare(args)


def _run_compare(args: argparse.Namespace) -> None:
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(name)s: %(message)s",
    )

    old_dir = args.old
    new_dir = args.new

    if not old_dir.is_dir():
        print(f"Error: old directory not found: {old_dir}")
        sys.exit(1)
    if not new_dir.is_dir():
        print(f"Error: new directory not found: {new_dir}")
        sys.exit(1)

    # Print summary of matched files before comparison
    matched, old_only, new_only = match_fbdi_files(old_dir, new_dir)

    print(f"Matched file pairs: {len(matched)}")
    if old_only:
        print(f"Old-only files ({len(old_only)}):")
        for f in old_only:
            print(f"  - {f.name}")
    if new_only:
        print(f"New-only files ({len(new_only)}):")
        for f in new_only:
            print(f"  - {f.name}")

    print(f"\nComparing {len(matched)} file pairs...")

    output_path, skipped_files = compare_all(
        old_dir,
        new_dir,
        args.output,
        changes_only=not args.all_rows,
    )

    # Count changes in output
    from openpyxl import load_workbook
    wb = load_workbook(output_path, read_only=True)
    ws = wb.active
    change_count = max((ws.max_row or 1) - 1, 0)
    wb.close()

    print(f"\nChanges found: {change_count}")
    print(f"Output written to: {output_path}")

    if skipped_files:
        from fbdi.config import MAX_FILE_SIZE_BYTES
        limit_mb = MAX_FILE_SIZE_BYTES // (1024 * 1024)
        print(f"\n{'=' * 60}")
        print(f"WARNING: {len(skipped_files)} file(s) were skipped (>{limit_mb}MB) and excluded from this report.")
        print("These files require manual review:")
        for s in skipped_files:
            print(f"  - {s['name']} ({s['size_mb']:.1f}MB)")
        print(f"{'=' * 60}")
