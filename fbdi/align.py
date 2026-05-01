"""LCS-style alignment of FBDI tab rows across two releases.

Pure function — no I/O. Takes two lists of AlignedField (one per release)
and returns a typed Change list classifying every difference: ADDED,
REMOVED, MODIFIED, RENAMED, SHIFTED, MULTI.

The algorithm matches fields by identity (technical name first, label
fallback) using longest common subsequence, then classifies each matched
pair across three independent axes (label, metadata, position). Unmatched
rows on either side become ADDED or REMOVED.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class AlignedField:
    """One field at one position within one release."""
    position: int
    label: str | None
    technical: str | None
    data_type: str | None
    length: int | None
    required: bool | None


@dataclass(frozen=True)
class Change:
    """One classified difference between two releases."""
    change_type: str                      # ADDED | REMOVED | MODIFIED | RENAMED | SHIFTED | MULTI
    old_position: int | None
    new_position: int | None
    old_field: AlignedField | None
    new_field: AlignedField | None
    axes: tuple[str, ...] = ()            # subset of ("label", "metadata", "position")
    sub_kinds: tuple[str, ...] = ()       # subset of ("type", "length", "required") when metadata changed


def align_tabs(old: list[AlignedField], new: list[AlignedField]) -> list[Change]:
    """Align two release row lists and return classified changes."""
    if not old and not new:
        return []
    if not old:
        return [
            Change(change_type="ADDED", old_position=None, new_position=f.position,
                   old_field=None, new_field=f)
            for f in new
        ]
    if not new:
        return [
            Change(change_type="REMOVED", old_position=f.position, new_position=None,
                   old_field=f, new_field=None)
            for f in old
        ]
    # Matched-pair classification not implemented yet.
    raise NotImplementedError("matched-pair classification — Task 3")
