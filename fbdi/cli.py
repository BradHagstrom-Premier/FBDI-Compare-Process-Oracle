"""CLI entry point for the FBDI comparison engine."""

import argparse
import logging
import sys
from pathlib import Path

from fbdi.compare import compare_all
from fbdi.utils import match_fbdi_files


def _resolve_dir(path: Path) -> Path:
    """Resolve a release label to its baselines originals directory.

    If path is already a directory, return it unchanged.
    Otherwise, try baselines/<path>/originals/ as a convenience shorthand.
    Falls through to the original path if no match (caller handles the error).
    """
    if path.is_dir():
        return path
    candidate = Path("baselines") / str(path) / "originals"
    if candidate.is_dir():
        return candidate
    return path


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

    diagnose_parser = subparsers.add_parser(
        "diagnose",
        help="Diagnose header detection outcomes for FBDI templates",
    )
    diagnose_parser.add_argument(
        "--release", type=str, default=None,
        help="Release label (e.g. 26a) — looks in baselines/<release>/",
    )
    diagnose_parser.add_argument(
        "--old", type=Path, default=None,
        help="Path to old release directory",
    )
    diagnose_parser.add_argument(
        "--new", type=Path, default=None,
        help="Path to new release directory",
    )
    diagnose_parser.add_argument(
        "--output", type=Path, default=None,
        help="Output file path (default: Diagnostic_Report_<label>.xlsx)",
    )
    diagnose_parser.add_argument(
        "--verbose", action="store_true",
        help="Set logging to DEBUG",
    )

    args = parser.parse_args(argv)

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "compare":
        _run_compare(args)
    elif args.command == "diagnose":
        _run_diagnose(args)


def _run_compare(args: argparse.Namespace) -> None:
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(name)s: %(message)s",
    )

    old_dir = _resolve_dir(args.old)
    new_dir = _resolve_dir(args.new)

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


def _run_diagnose(args: argparse.Namespace) -> None:
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(name)s: %(message)s",
    )

    from fbdi.diagnose import diagnose_file, write_diagnostic_report

    # Resolve directories
    dirs: list[Path] = []
    label_parts: list[str] = []

    if args.release:
        release_dir = Path("baselines") / args.release
        if not release_dir.is_dir():
            print(f"Error: release directory not found: {release_dir}")
            sys.exit(1)
        dirs.append(release_dir)
        label_parts.append(args.release.upper())
    elif args.old or args.new:
        if not args.old or not args.new:
            print("Error: --old and --new must be used together")
            sys.exit(1)
        for d, flag in [(args.old, "--old"), (args.new, "--new")]:
            if not d.is_dir():
                print(f"Error: directory not found ({flag}): {d}")
                sys.exit(1)
        dirs.extend([args.old, args.new])
        label_parts.extend([args.old.name.upper(), args.new.name.upper()])
    else:
        print("Error: provide --release or --old/--new")
        sys.exit(1)

    # Determine output path
    label = "_".join(label_parts)
    output_path = args.output or Path(f"Diagnostic_Report_{label}.xlsx")

    # Scan files
    all_rows = []
    for directory in dirs:
        xlsm_files = sorted(directory.glob("*.xlsm"))
        print(f"Scanning {len(xlsm_files)} files in {directory} ...")
        for file_path in xlsm_files:
            rows = diagnose_file(file_path)
            all_rows.extend(rows)

    detected = sum(1 for r in all_rows if r.detection_result == "DETECTED")
    no_header = sum(1 for r in all_rows if r.detection_result == "NO_HEADER")
    skipped_tab = sum(1 for r in all_rows if r.detection_result == "SKIPPED_TAB")
    file_too_large = sum(1 for r in all_rows if r.detection_result == "FILE_TOO_LARGE")
    file_error = sum(1 for r in all_rows if r.detection_result == "FILE_ERROR")

    write_diagnostic_report(all_rows, output_path)

    print(f"\nDiagnostic complete: {len(all_rows)} tab entries")
    print(f"  DETECTED:       {detected}")
    print(f"  NO_HEADER:      {no_header}")
    print(f"  SKIPPED_TAB:    {skipped_tab}")
    print(f"  FILE_TOO_LARGE: {file_too_large}")
    print(f"  FILE_ERROR:     {file_error}")
    print(f"Output written to: {output_path}")
