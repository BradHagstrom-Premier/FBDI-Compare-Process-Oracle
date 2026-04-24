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


MODULE_PREFIXES = {
    "project-management": [
        "Import", "Project", "Resource", "Idea", "Lease", "Revenue",
        "FinancialProject", "ExpenseLease",
    ],
    "financials": [
        "Payables", "Receivables", "FixedAsset", "Cash", "General", "Journal",
        "Account", "Chartof", "Daily", "AutoInvoice", "Cross", "Intercompany",
        "Gl", "Netting", "Tax", "Budget", "Attachment", "Xla", "ZX_",
        "Configurator", "Create", "IbyLegacy", "FiscalDocument",
        "ImportStandaloneFiscal", "InboundFiscal", "UploadCredit", "UploadCustomers",
        "BillingData", "RapidImplementation",
    ],
    "procurement": [
        "PO", "Requisition", "Supplier", "ChangeOrder", "Poi", "PONN",
        "Sch", "ImportDocumentActions", "ProductProposal",
    ],
    "supply-chain-and-manufacturing": [
        "Scp", "Work", "Cse", "Maintenance", "Mnt", "Inventory", "Item",
        "Order", "Egp", "Sus", "Vcs", "Ship", "Source", "Production",
        "Perform", "Process", "CycleCount", "Dos", "InterfacedPick",
        "Receiving", "Requirement", "StandardCost", "CostLists",
        "DiscountList", "PriceList", "CustomerImport",
    ],
}


def _module_for_filename(name: str) -> str:
    """Match filename to Oracle module using the longest prefix across all
    modules. Longest-first ordering matters because several modules share a
    common short prefix (e.g. project-management's "Import" would otherwise
    swallow financials-specific "ImportStandaloneFiscal*" files).
    """
    candidates: list[tuple[int, str, str]] = [
        (len(prefix), module, prefix)
        for module, prefixes in MODULE_PREFIXES.items()
        for prefix in prefixes
        if name.startswith(prefix)
    ]
    if not candidates:
        return "other"
    # Longest prefix wins; tie-break by module order is unimportant
    # since same-length collisions across modules don't occur in this set.
    candidates.sort(key=lambda t: -t[0])
    return candidates[0][1]


def group_missing_by_module(missing: list[str]) -> dict[str, list[str]]:
    """Group missing filenames by best-guess Oracle docs module.

    Heuristic prefix-based match. Returns {module: sorted_names}.
    Empty input returns {}.
    """
    if not missing:
        return {}
    groups: dict[str, list[str]] = {}
    for name in missing:
        module = _module_for_filename(name)
        groups.setdefault(module, []).append(name)
    return {k: sorted(v) for k, v in groups.items()}


def most_recent_release(inventory: dict[str, list[str]]) -> str | None:
    """Return the ASCII-max release key from the inventory, or None."""
    if not inventory:
        return None
    return max(inventory.keys())


def compute_first_run_delta(
    downloaded_count: int,
    inventory: dict[str, list[str]],
) -> dict:
    """For the first-run bootstrap case, compare download count to the most
    recent prior release. Returns {prior_release, prior_count, delta_pct,
    over_threshold}. delta_pct is relative ((new-prior)/prior); always
    non-negative (we care about absolute deviation)."""
    prior = most_recent_release(inventory)
    if prior is None or not inventory[prior]:
        return {
            "prior_release": None,
            "prior_count": 0,
            "delta_pct": 0.0,
            "over_threshold": False,
        }
    prior_count = len(inventory[prior])
    delta = abs(downloaded_count - prior_count) / prior_count
    return {
        "prior_release": prior,
        "prior_count": prior_count,
        "delta_pct": delta,
        "over_threshold": delta > FIRST_RUN_DELTA_THRESHOLD,
    }


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


def _format_section(release: str, filenames: list[str]) -> str:
    sorted_names = sorted(filenames)
    banner = "=" * 28
    lines = [
        banner,
        f"{release} ORIGINALS ({len(sorted_names)} {'file' if len(sorted_names) == 1 else 'files'})",
        banner,
        *sorted_names,
        "",
    ]
    return "\n".join(lines) + "\n"


def _strip_differences_block(text: str) -> str:
    """Remove any existing DIFFERENCES block (we'll regenerate it)."""
    return re.sub(
        r"={20,}\s*\nDIFFERENCES\s*\n={20,}\s*\n(?:.*\n?)*\Z",
        "",
        text,
        flags=re.IGNORECASE | re.MULTILINE,
    ).rstrip() + "\n"


def _render_differences(inventory: dict[str, list[str]]) -> str:
    """Regenerate the DIFFERENCES footer from the current inventory.

    For each pair of adjacent releases (by ASCII sort), emit
    'Only in <NEW>: <comma-list>'. Simple and mirrors the existing file.
    """
    releases = sorted(inventory.keys())
    if len(releases) < 2:
        return ""
    banner = "=" * 28
    lines = [banner, "DIFFERENCES", banner]
    for i in range(1, len(releases)):
        prev, cur = releases[i - 1], releases[i]
        only_in_cur = sorted(set(inventory[cur]) - set(inventory[prev]))
        only_in_prev = sorted(set(inventory[prev]) - set(inventory[cur]))
        if only_in_cur:
            lines.append(f"Only in {cur}: {', '.join(only_in_cur)}")
        if only_in_prev:
            lines.append(f"Only in {prev}: {', '.join(only_in_prev)}")
    return "\n".join(lines) + "\n"


def commit_inventory(
    inventory_text: str, release: str, filenames: list[str],
) -> str:
    """Return a new inventory text with release's section inserted or replaced,
    and the DIFFERENCES footer regenerated."""
    release = release.upper()
    inventory = parse_inventory(inventory_text)
    if inventory_text and not inventory_text.endswith("\n"):
        inventory_text = inventory_text + "\n"
    inventory[release] = sorted(filenames)

    # Strip old DIFFERENCES
    text = _strip_differences_block(inventory_text)

    # Replace existing section for `release` if present
    section_pattern = re.compile(
        r"={20,}\s*\n" + re.escape(release) + r"\s+ORIGINALS\s*\([^)]*\)\s*\n={20,}\s*\n"
        r"(?:[^\n]*\n)*?"
        r"(?=(?:={20,}\s*\n(?:\d{2}[A-D]\s+ORIGINALS|DIFFERENCES))|\Z)",
        re.IGNORECASE,
    )
    new_section = _format_section(release, inventory[release])
    if section_pattern.search(text):
        text = section_pattern.sub(new_section, text, count=1)
    else:
        # Append after last ORIGINALS section (or at end if none)
        text = text.rstrip() + "\n\n" + new_section

    # Re-append DIFFERENCES footer
    diffs = _render_differences(inventory)
    if diffs:
        text = text.rstrip() + "\n\n" + diffs
    return text


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
    parser.add_argument(
        "--commit-inventory", action="store_true",
        help="Rewrite inventory to match the downloaded files for --release",
    )
    args = parser.parse_args(argv)

    release = args.release.upper()
    originals = args.originals or (Path("baselines") / release / "originals")
    downloaded = list_downloaded(originals)
    inventory_text = args.inventory.read_text(encoding="utf-8") if args.inventory.is_file() else ""
    inventory = parse_inventory(inventory_text)

    # Short-circuit: --commit-inventory rewrites baseline_files.txt and exits 0.
    if args.commit_inventory:
        new_text = commit_inventory(inventory_text, release, downloaded)
        args.inventory.write_text(new_text, encoding="utf-8")
        payload = {
            "release": release,
            "committed": True,
            "count": len(downloaded),
            "inventory_path": str(args.inventory),
        }
        print(json.dumps(payload, indent=2))
        return 0

    # First-run: no section for this release
    if release not in inventory:
        delta = compute_first_run_delta(len(downloaded), inventory)
        payload = {
            "release": release,
            "first_run": True,
            "downloaded_count": len(downloaded),
            "downloaded": downloaded,
            **delta,
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
        "missing_by_module": group_missing_by_module(diff["missing"]),
    }
    print(json.dumps(payload, indent=2))

    if diff["missing"]:
        return 1
    if diff["extras"]:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
