"""Shared utilities for the FBDI comparison engine."""

from pathlib import Path


def col_index_to_letter(index: int) -> str:
    """Convert 1-based column index to Excel column letter. 1->A, 27->AA, etc."""
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def match_fbdi_files(
    old_dir: Path, new_dir: Path
) -> tuple[list[tuple[Path, Path]], list[Path], list[Path]]:
    """Match FBDI files between old and new directories by filename stem.

    Matches are case-insensitive and support both .xlsm and .xlsx extensions.
    Returns: (matched_pairs, old_only, new_only) sorted by stem.
    """
    extensions = {".xlsm", ".xlsx"}

    old_files = {
        f.stem.lower(): f
        for f in old_dir.iterdir()
        if f.suffix.lower() in extensions
    }
    new_files = {
        f.stem.lower(): f
        for f in new_dir.iterdir()
        if f.suffix.lower() in extensions
    }

    old_stems = set(old_files.keys())
    new_stems = set(new_files.keys())

    matched = sorted(
        [(old_files[s], new_files[s]) for s in old_stems & new_stems],
        key=lambda pair: pair[0].stem.lower(),
    )
    old_only = sorted(
        [old_files[s] for s in old_stems - new_stems],
        key=lambda p: p.stem.lower(),
    )
    new_only = sorted(
        [new_files[s] for s in new_stems - old_stems],
        key=lambda p: p.stem.lower(),
    )

    return matched, old_only, new_only
