"""Stage 3 download verification for fbdi-compare-release.

Diffs baselines/<ver>/originals/ against the <ver> section of
baseline_files.txt. Handles first-run bootstrap (no <ver> section yet) and
commits an updated inventory on demand.

Exit codes:
    0 = clean (missing == 0, extras == 0)
    1 = missing > 0  (triggers retry / §5 #5 prompt)
    2 = extras only  (triggers §5 #6 prompt)
    3 = first-run bootstrap required (no <ver> section in inventory)
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

MANUAL_FILES = ["RapidImplementationForCashManagement.xlsm"]
FIRST_RUN_DELTA_THRESHOLD = 0.15  # 15%, per spec §5 #6

_SECTION_RE = re.compile(r"^(\d{2}[A-D])\s+ORIGINALS\s*\(\d+\s+files?\)\s*$", re.IGNORECASE)


def parse_inventory(text: str) -> dict[str, list[str]]:
    """Parse baseline_files.txt into {release: [filenames...]}.

    Recognizes sections of the form:
        ============================
        26A ORIGINALS (212 files)
        ============================
        <filename>.xlsm
        ...

    Lines not ending in .xlsm are ignored inside sections. Sections end at
    the next `===` delimiter or EOF. A 'DIFFERENCES' section header is not
    an ORIGINALS section and its content is discarded.
    """
    result: dict[str, list[str]] = {}
    current_release: str | None = None
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _SECTION_RE.match(line)
        if m:
            current_release = m.group(1).upper()
            result.setdefault(current_release, [])
            i += 1
            continue
        if line.startswith("==="):
            # Delimiter line — doesn't change state on its own; next non-delim
            # line decides. Sections are terminated by the next SECTION_RE match
            # or a non-.xlsm header block.
            i += 1
            continue
        if current_release is not None and line.lower().endswith(".xlsm"):
            result[current_release].append(line)
        elif current_release is not None and line and not line.lower().endswith(".xlsm"):
            # A non-blank non-.xlsm line inside a section could be a new
            # free-text block (e.g. "DIFFERENCES"). End the current section.
            if line.upper() == "DIFFERENCES" or re.search(r"[A-Za-z]", line) and ":" in line:
                current_release = None
        i += 1
    # Sort each section for deterministic diffs
    for k in result:
        result[k] = sorted(result[k])
    return result


def diff_against_inventory(
    release: str,
    downloaded_names: list[str],
    inventory: dict[str, list[str]],
    manual_files: list[str],
) -> dict:
    """Return {"missing": [...], "extras": [...]}.

    missing = inventory[release] - downloaded - manual_files
    extras  = downloaded - inventory[release]
    """
    expected = set(inventory.get(release.upper(), []))
    actual = set(downloaded_names)
    manual = set(manual_files)

    missing = sorted((expected - actual) - manual)
    extras = sorted(actual - expected)
    return {"missing": missing, "extras": extras}


def list_downloaded(originals_dir: Path) -> list[str]:
    if not originals_dir.is_dir():
        return []
    return sorted(
        p.name for p in originals_dir.iterdir()
        if p.suffix.lower() == ".xlsm" and not p.name.startswith("~$")
    )


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Stage 3 download verification")
    parser.add_argument("--release", required=True, help="Release label, e.g. 26B")
    parser.add_argument(
        "--inventory", type=Path, default=Path("baseline_files.txt"),
        help="Path to baseline_files.txt (default: ./baseline_files.txt)",
    )
    parser.add_argument(
        "--originals", type=Path, default=None,
        help="Path to baselines/<release>/originals/ (default: derived from --release)",
    )
    args = parser.parse_args(argv)

    release = args.release.upper()
    originals = args.originals or (Path("baselines") / release / "originals")
    downloaded = list_downloaded(originals)
    inventory_text = args.inventory.read_text(encoding="utf-8") if args.inventory.is_file() else ""
    inventory = parse_inventory(inventory_text)

    # First-run: no section for this release
    if release not in inventory:
        payload = {
            "release": release,
            "first_run": True,
            "downloaded_count": len(downloaded),
            "downloaded": downloaded,
        }
        print(json.dumps(payload, indent=2))
        return 3

    diff = diff_against_inventory(release, downloaded, inventory, MANUAL_FILES)
    payload = {
        "release": release,
        "first_run": False,
        "downloaded_count": len(downloaded),
        "expected_count": len(inventory[release]),
        "missing": diff["missing"],
        "extras": diff["extras"],
    }
    print(json.dumps(payload, indent=2))

    if diff["missing"]:
        return 1
    if diff["extras"]:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
