"""Stage 8 post-run verification for fbdi-compare-release.

- Runs `python -m fbdi diagnose --release <ver>` and parses the Diagnostic
  xlsx output for NO_HEADER regressions.
- Reads FBDI_Master_Catalog.xlsx Issues tab, filters by release, and flags
  catalog Issues-tab regression if:
      new_count > 2 * prior_count   OR   new_count - prior_count > 50
- Reconciles baselines/<ver>/originals/ against baseline_files.txt (the
  release's own section, else the prior release). A file the reference expects
  but the folder lacks is a HARD failure — this is the safety net for the
  silent-baseline-drift class of bug (a template that failed to download and was
  never noticed).

Exit codes:
    0 = clean
    1 = soft regression detected (diagnose / catalog Issues) — non-blocking
    2 = baseline inventory hard-fail (originals folder short vs reference) — blocking
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections import Counter
from contextlib import closing
from pathlib import Path

from openpyxl import load_workbook

# verify_download lives beside this script; reuse its inventory helpers rather
# than duplicating the reconciliation logic.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import verify_download  # noqa: E402

ISSUE_MULTIPLIER_THRESHOLD = 2.0
ISSUE_ABSOLUTE_THRESHOLD = 50


def check_catalog_issues(catalog_path: Path, release: str) -> dict:
    """Read Issues tab, group by release, compare release against most-recent prior."""
    release = release.upper()
    with closing(load_workbook(catalog_path, read_only=True, data_only=True)) as wb:
        if "Issues" not in wb.sheetnames:
            return {
                "release_issue_count": 0,
                "prior_release": None,
                "prior_issue_count": 0,
                "regression": False,
                "threshold": {"multiplier": ISSUE_MULTIPLIER_THRESHOLD,
                              "absolute": ISSUE_ABSOLUTE_THRESHOLD},
            }
        ws = wb["Issues"]
        counter: Counter[str] = Counter()
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0 or not row or not row[0]:
                continue
            counter[str(row[0]).upper()] += 1

    release_count = counter.get(release, 0)
    priors = sorted(r for r in counter if r < release)
    prior = priors[-1] if priors else None
    prior_count = counter.get(prior, 0) if prior else 0

    regression = False
    if prior and prior_count > 0:
        if release_count > ISSUE_MULTIPLIER_THRESHOLD * prior_count:
            regression = True
        if release_count - prior_count > ISSUE_ABSOLUTE_THRESHOLD:
            regression = True

    return {
        "release_issue_count": release_count,
        "prior_release": prior,
        "prior_issue_count": prior_count,
        "regression": regression,
        "threshold": {"multiplier": ISSUE_MULTIPLIER_THRESHOLD,
                      "absolute": ISSUE_ABSOLUTE_THRESHOLD},
    }


def run_diagnose(release: str, repo_root: Path) -> dict:
    """Invoke `python -m fbdi diagnose --release <release>` and parse its output xlsx."""
    diagnostic_path = repo_root / f"Diagnostic_Report_{release.upper()}.xlsx"
    if diagnostic_path.is_file():
        diagnostic_path.unlink()

    proc = subprocess.run(
        [sys.executable, "-m", "fbdi", "diagnose", "--release", release],
        cwd=repo_root,
        capture_output=True, text=True,
    )
    if proc.returncode != 0 or not diagnostic_path.is_file():
        return {
            "skipped": False,
            "no_header_count": None,
            "file_error_count": None,
            "regression": False,
            "error": f"diagnose invocation failed: {proc.stderr[:500]}",
        }

    no_header = 0
    file_error = 0
    with closing(load_workbook(diagnostic_path, read_only=True, data_only=True)) as wb:
        ws = wb.active
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0 or not row:
                continue
            result = row[2]  # "Detection Result" column
            if result == "NO_HEADER":
                no_header += 1
            elif result == "FILE_ERROR":
                file_error += 1

    return {
        "skipped": False,
        "no_header_count": no_header,
        "file_error_count": file_error,
        "regression": no_header > 0,
        "diagnostic_path": str(diagnostic_path),
    }


def check_baseline_inventory(
    new_release: str,
    originals_root: Path,
    inventory_path: Path,
) -> dict:
    """Reconcile baselines/<new_release>/originals/ against baseline_files.txt.

    Prefers the release's own committed section as the reference; falls back to
    the most-recent prior release when it has none. Asymmetric: only a shortfall
    (a file the reference expects but the folder lacks, excluding MANUAL_FILES)
    is a hard failure; net additions are surfaced but never block. Reuses the
    reconciliation helpers in verify_download so there is a single source of
    truth for inventory logic.
    """
    new_release = new_release.upper()
    originals = originals_root / new_release / "originals"
    current = verify_download.list_downloaded(originals)
    inventory_text = (
        inventory_path.read_text(encoding="utf-8") if inventory_path.is_file() else ""
    )
    inventory = verify_download.parse_inventory(inventory_text)

    if new_release in inventory:
        diff = verify_download.diff_against_inventory(
            new_release, current, inventory, verify_download.MANUAL_FILES,
        )
        dropped = diff["missing"]
        added = diff["extras"]
        reference = "own-section"
    else:
        prior_diff = verify_download.diff_against_prior(
            current, inventory, new_release, verify_download.MANUAL_FILES,
        )
        dropped = prior_diff["dropped"]
        added = prior_diff["added"]
        prior_rel = prior_diff["prior_release"]
        reference = f"prior:{prior_rel}" if prior_rel else "none"

    return {
        "release": new_release,
        "reference": reference,
        "current_count": len(current),
        "dropped": dropped,
        "added": added,
        "missing_by_module": verify_download.group_missing_by_module(dropped),
        "hard_fail": bool(dropped),
        "inventory_present": bool(inventory),
    }


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Stage 8 post-run verification")
    parser.add_argument("--release", required=True)
    parser.add_argument(
        "--catalog", type=Path, default=Path("FBDI_Master_Catalog.xlsx"),
    )
    parser.add_argument(
        "--repo-root", type=Path, default=Path.cwd(),
        help="Repo root for fbdi diagnose invocation",
    )
    parser.add_argument(
        "--skip-diagnose", action="store_true",
        help="Skip the diagnose subprocess (for unit tests / quick runs)",
    )
    parser.add_argument(
        "--originals-root", type=Path, default=Path("baselines"),
        help="Root of the per-release baseline folders (default: ./baselines)",
    )
    parser.add_argument(
        "--inventory", type=Path, default=Path("baseline_files.txt"),
        help="Path to baseline_files.txt (default: ./baseline_files.txt)",
    )
    parser.add_argument(
        "--skip-inventory", action="store_true",
        help="Skip the baseline inventory reconciliation check",
    )
    args = parser.parse_args(argv)

    release = args.release.upper()

    if args.skip_diagnose:
        diag = {
            "skipped": True,
            "no_header_count": None,
            "file_error_count": None,
            "regression": False,
        }
    else:
        diag = run_diagnose(release, args.repo_root)

    cat = check_catalog_issues(args.catalog, release)

    if args.skip_inventory:
        inv = {"skipped": True, "hard_fail": False}
    else:
        inv = check_baseline_inventory(release, args.originals_root, args.inventory)

    # Soft regressions (diagnose / catalog Issues) stay non-blocking — exit 1,
    # surfaced by the skill as a warning. A baseline inventory shortfall is a
    # HARD failure — exit 2, which the skill treats as fatal. Exit 2 takes
    # precedence so a shortfall is never masked by a concurrent soft regression.
    soft_regression = bool(diag.get("regression")) or bool(cat.get("regression"))
    inventory_hard_fail = bool(inv.get("hard_fail"))
    payload = {
        "release": release,
        "diagnose": diag,
        "catalog_issues": cat,
        "baseline_inventory": inv,
        "overall_regression": soft_regression,
        "inventory_hard_fail": inventory_hard_fail,
    }
    print(json.dumps(payload, indent=2))
    if inventory_hard_fail:
        return 2
    return 1 if soft_regression else 0


if __name__ == "__main__":
    sys.exit(main())
