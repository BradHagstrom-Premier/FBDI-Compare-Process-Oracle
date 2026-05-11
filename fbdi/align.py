"""LCS-style alignment of FBDI tab rows across two releases.

Pure function — no I/O. Takes two lists of AlignedField (one per release)
and returns a typed Change list classifying every difference: ADDED,
REMOVED, MODIFIED, RENAMED, SHIFTED, MULTI.

The algorithm matches fields by identity (technical name first, label
fallback) using longest common subsequence, then classifies each matched
pair across three independent axes (label, metadata, position). The
metadata axis decomposes further into sub-kinds (type, length, scale,
required). Unmatched rows on either side become ADDED or REMOVED.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class AlignedField:
    """One field at one position within one release."""
    position: int
    label: str | None
    technical: str | None
    data_type: str | None
    length: int | None
    scale: int | None
    required: bool | None
    data_type_raw: str | None = None


@dataclass(frozen=True)
class Change:
    """One classified difference between two releases."""
    change_type: str                      # ADDED | REMOVED | MODIFIED | RENAMED | SHIFTED | MULTI
    old_position: int | None
    new_position: int | None
    old_field: AlignedField | None
    new_field: AlignedField | None
    axes: tuple[str, ...] = ()            # subset of ("label", "metadata", "position")
    sub_kinds: tuple[str, ...] = ()       # subset of ("type", "length", "scale", "required") when metadata changed


def _identity_key(f: AlignedField) -> tuple[str, str]:
    """Stable identity for matching across releases.

    Prefers technical name (canonical, position-independent). Falls back to
    label when technical is missing (thin tabs). Tag distinguishes the
    space so a label "ITEM_NAME" never matches a technical "ITEM_NAME".

    Note: fields with both technical=None and label="" all share the same
    identity key ("label", "") and are matched positionally by the LCS.
    Verified on the 26A→26B catalog: zero suspicious blank-label change
    rows, so this collision is benign on real Oracle FBDI data.
    """
    if f.technical:
        return ("tech", f.technical)
    return ("label", f.label or "")


def _lcs_match(old: list[AlignedField], new: list[AlignedField]) -> list[tuple[int, int]]:
    """Longest common subsequence over identity keys. Returns matched index pairs.

    Standard O(m*n) DP. Indices are 0-based positions in the input lists.
    """
    m, n = len(old), len(new)
    if m == 0 or n == 0:
        return []
    old_keys = [_identity_key(f) for f in old]
    new_keys = [_identity_key(f) for f in new]
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(m):
        for j in range(n):
            if old_keys[i] == new_keys[j]:
                dp[i + 1][j + 1] = dp[i][j] + 1
            else:
                dp[i + 1][j + 1] = max(dp[i][j + 1], dp[i + 1][j])
    # Backtrack to recover the matched pairs.
    pairs: list[tuple[int, int]] = []
    i, j = m, n
    while i > 0 and j > 0:
        if old_keys[i - 1] == new_keys[j - 1]:
            pairs.append((i - 1, j - 1))
            i -= 1
            j -= 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1
    pairs.reverse()
    return pairs


def _classify_pair(old_f: AlignedField, new_f: AlignedField) -> Change | None:
    """Classify a matched pair across three axes; None if unchanged."""
    label_changed = (old_f.label or "") != (new_f.label or "")
    metadata_kinds: list[str] = []
    if (old_f.data_type or "") != (new_f.data_type or ""):
        metadata_kinds.append("type")
    if old_f.length != new_f.length:
        metadata_kinds.append("length")
    if old_f.scale != new_f.scale:
        metadata_kinds.append("scale")
    if old_f.required != new_f.required:
        metadata_kinds.append("required")
    metadata_changed = bool(metadata_kinds)
    position_changed = old_f.position != new_f.position

    axes = []
    if label_changed:
        axes.append("label")
    if metadata_changed:
        axes.append("metadata")
    if position_changed:
        axes.append("position")
    if not axes:
        return None

    if len(axes) == 1:
        change_type = {"label": "RENAMED", "metadata": "MODIFIED", "position": "SHIFTED"}[axes[0]]
    else:
        change_type = "MULTI"

    return Change(
        change_type=change_type,
        old_position=old_f.position,
        new_position=new_f.position,
        old_field=old_f,
        new_field=new_f,
        axes=tuple(axes),
        sub_kinds=tuple(metadata_kinds),
    )


def align_tabs(old: list[AlignedField], new: list[AlignedField]) -> list[Change]:
    """Align two release row lists and return classified changes."""
    if not old and not new:
        return []

    matched = _lcs_match(old, new)
    matched_old = {i for i, _ in matched}
    matched_new = {j for _, j in matched}

    changes: list[Change] = []
    # REMOVED: old fields with no match
    for i, f in enumerate(old):
        if i not in matched_old:
            changes.append(Change(
                change_type="REMOVED", old_position=f.position, new_position=None,
                old_field=f, new_field=None,
            ))
    # ADDED: new fields with no match
    for j, f in enumerate(new):
        if j not in matched_new:
            changes.append(Change(
                change_type="ADDED", old_position=None, new_position=f.position,
                old_field=None, new_field=f,
            ))
    # Classified pair changes
    for i, j in matched:
        c = _classify_pair(old[i], new[j])
        if c is not None:
            changes.append(c)

    # Stable sort: by new_position (None last), then old_position.
    changes.sort(key=lambda c: (c.new_position is None, c.new_position or 0,
                                c.old_position is None, c.old_position or 0))
    return changes
