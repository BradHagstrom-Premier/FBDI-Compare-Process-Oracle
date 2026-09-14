# Repo Cleanse — Phase 1 Implementation Plan (PR #1)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove the Applaud subsystem (superseded by ApplaudMCP), decouple the report from Applaud into a consultant-facing release-change report, de-clutter the tree, and rewrite the docs — leaving a focused core+catalog+report repo with a green test suite.

**Architecture:** The four layers are cleanly separable — core imports nothing from Applaud, and no Applaud module imports the core. We remove the Applaud CLI commands first (they consume `report.load_mapping`), then rewrite `report.py` to drop all Applaud coupling and source module grouping from `file_modules.json`, then delete the now-orphaned Applaud modules/tools/tests, then clean artifacts/docs/skill/CLAUDE.md.

**Tech Stack:** Python 3.14+, openpyxl, jinja2, weasyprint (opt-in only), pytest.

**Spec:** `docs/superpowers/specs/2026-09-14-repo-cleanse-and-report-redesign-design.md`

## Global Constraints

- **Python:** 3.14+. On this Windows machine use the `py` launcher: `py -m pytest tests/`, `py -m fbdi ...` (bare `python` is the Store stub).
- **Tests green after every task:** `py -m pytest tests/` must pass at each commit. Baseline before Phase 1 = **458 tests**.
- **Workflow:** all work on branch `chore/repo-cleanse` (already created). Never push to `master`, never self-merge. Push + open PR only after the whole plan is done and verified.
- **Commit attribution:** end every commit message with `Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`.
- **`catalog_normalize` stays** — it produces the catalog's `column_label` (catalog.py:320,432), so it is a general label normalizer now, not Applaud-only. Do not delete it; only its docstring framing is Applaud-specific (left for the docs task).
- **No Applaud residue:** after Task 3, `grep -ri applaud fbdi/ tools/ tests/` returns nothing.

---

### Task 1: Remove the Applaud CLI subcommands

Removes the four Applaud subcommands and their handlers. These are the only in-CLI consumers of `report.load_mapping` and the Applaud modules, so removing them first lets Task 2 drop `load_mapping` cleanly.

**Files:**
- Modify: `fbdi/cli.py` — delete the `populate-module`, `audit-applaud`, `correspondence-derive`, `correspondence-confirm` subparsers (lines ~109-125, ~153-210), their dispatch branches (lines ~224-233), and their handler functions `_run_populate_module`, `_run_audit_applaud`, `_run_correspondence_derive`, `_run_correspondence_confirm` (lines ~406-461, ~491-662).
- Modify: `tests/test_cli.py` — delete `test_populate_module_subcommand_invocation` (lines 78-122) and `test_populate_module_missing_json_exits_2` (lines 125-137).

**Interfaces:**
- Produces: `fbdi/cli.py` retaining only the `compare`, `diagnose`, `catalog`, `report` subcommands and handlers `_run_compare`, `_run_diagnose`, `_run_catalog`, `_run_report`, plus `_resolve_dir` and `main`. `_run_report` is still the OLD signature here (Task 2 rewrites it).

- [ ] **Step 1: Remove the four subparsers and dispatch branches**

In `fbdi/cli.py`, delete the `populate_parser`, `audit_applaud_parser`, `corr_derive_parser`, `corr_confirm_parser` blocks. In the dispatch chain (currently lines 218-233), reduce to:

```python
    if args.command == "compare":
        _run_compare(args)
    elif args.command == "diagnose":
        _run_diagnose(args)
    elif args.command == "catalog":
        _run_catalog(args)
    elif args.command == "report":
        _run_report(args)
```

- [ ] **Step 2: Delete the four handler functions**

Delete `_run_populate_module`, `_run_audit_applaud`, `_run_correspondence_derive`, `_run_correspondence_confirm` entirely. Leave `_run_compare`, `_run_diagnose`, `_run_catalog`, `_run_report` intact.

- [ ] **Step 3: Delete the populate-module CLI tests**

Remove the two `test_populate_module_*` functions from `tests/test_cli.py`.

- [ ] **Step 4: Verify import + help + suite**

Run:
```bash
py -c "import fbdi.cli"
py -m fbdi --help
py -m pytest tests/test_cli.py -q
```
Expected: import clean; help lists only `compare, diagnose, catalog, report`; test_cli passes. (`_run_report` still works — report.py untouched.)

- [ ] **Step 5: Commit**

```bash
git add fbdi/cli.py tests/test_cli.py
git commit -m "$(cat <<'EOF'
chore(cli): remove Applaud subcommands (populate-module, audit-applaud, correspondence-*)

Superseded by ApplaudMCP. Removes the only in-CLI consumers of
report.load_mapping and the Applaud modules, unblocking the report decouple.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 2: Decouple + repurpose the report (Applaud-free release-change report)

Rewrites `report.py`, its template, and its tests. Scope becomes *all changed tabs* grouped by module from `file_modules.json`; drops Applaud field names / types / 30-char flag / mapping / pending-base; PDF becomes opt-in; outputs renamed to `FBDI_Change_Report_*`.

**Files:**
- Rewrite: `fbdi/report.py`
- Rewrite: `fbdi/templates/report.html.j2` (minimal plain template — Phase 2 redesigns it)
- Rewrite: `tests/test_report.py`
- Modify: `fbdi/cli.py` — `_run_report` + the `report` subparser
- Modify: `tests/test_cli.py` — `TestReportSubcommand`

**Interfaces:**
- Consumes: `fbdi.align.{AlignedField, Change, align_tabs}` (unchanged).
- Produces:
  - `build_report_context(catalog_old, catalog_new, module_of: dict[str,str], old_release: str, new_release: str, generated_date: str|None=None) -> ReportContext`
  - `load_catalog_release(catalog_path: Path, release: str) -> dict[tuple[str,str], list[AlignedField]]` (unchanged behavior)
  - `load_file_modules(release: str) -> dict[str, str]`
  - `generate_report(catalog_path: Path, old_release: str, new_release: str, out_dir: Path, pdf: bool=False) -> tuple[Path, Path|None]`
  - Dataclasses: `ChangeRow`, `FileSection` (fields: `file, tab, module, changes_by_type, shift_summary, shift_is_uniform`), `ScopeTotals`, `ReportContext` (no `pending_base`).

- [ ] **Step 1: Write the new report.py**

Replace the entire contents of `fbdi/report.py` with:

```python
"""FBDI Release Change Report generator.

Reads the FBDI Master Catalog (per-release sheets), aligns each (file, tab)
across two releases, groups the changes by module, and emits a consultant-
facing HTML report ("here's what changed from <OLD> to <NEW>"), with an
opt-in PDF via weasyprint.

Module grouping comes from baselines/<release>/file_modules.json (written by
the downloader); the NEW release wins over OLD.

Public surface:
- build_report_context(...) — pure view-model construction (testable)
- load_catalog_release(...) — read one release sheet, grouped by (file, tab)
- load_file_modules(...)    — read a release's {file: module} map
- generate_report(...)      — load -> build -> render -> write
"""

from __future__ import annotations

import json
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from fbdi.align import AlignedField, Change, align_tabs

UNCLASSIFIED = "Unclassified"


@dataclass
class ChangeRow:
    """One row in a per-tab change-type table (view-model)."""
    change_type: str
    field_name: str            # FBDI technical name (or label when no technical)
    label: str
    oracle_type_str: str       # e.g. "VARCHAR2(30)" — empty when not applicable
    old_position: int | None
    new_position: int | None
    required: bool | None
    axes: tuple[str, ...]
    sub_kinds: tuple[str, ...]
    old_label: str | None = None
    new_label: str | None = None
    old_oracle_type_str: str | None = None
    new_oracle_type_str: str | None = None
    old_required: bool | None = None
    new_required: bool | None = None


@dataclass
class FileSection:
    """One per-tab section, grouped under a module."""
    file: str
    tab: str
    module: str
    changes_by_type: dict[str, list[ChangeRow]] = field(default_factory=dict)
    shift_summary: str | None = None
    shift_is_uniform: bool = False


@dataclass
class ScopeTotals:
    """Aggregate change counts across all sections (for the cover)."""
    tabs: int = 0
    add: int = 0
    rem: int = 0
    mod: int = 0
    shift: int = 0


@dataclass
class ReportContext:
    """Top-level view-model passed to the Jinja2 template."""
    old_release: str
    new_release: str
    generated_date: str
    file_sections: list[FileSection]
    totals: ScopeTotals = field(default_factory=ScopeTotals)


def build_report_context(
    catalog_old: dict[tuple[str, str], list[AlignedField]],
    catalog_new: dict[tuple[str, str], list[AlignedField]],
    module_of: dict[str, str],
    old_release: str,
    new_release: str,
    generated_date: str | None = None,
) -> ReportContext:
    """Build the report context from grouped catalog data + a file->module map.

    Scope: every (file, tab) present in either release that has at least one
    detected change. No mapping filter. module_of maps file_name -> module;
    files absent from it group under UNCLASSIFIED.
    """
    from datetime import date as _date
    if generated_date is None:
        generated_date = _date.today().isoformat()

    file_sections: list[FileSection] = []
    all_keys = set(catalog_old.keys()) | set(catalog_new.keys())

    for key in sorted(all_keys):
        file_name, tab = key
        changes = align_tabs(catalog_old.get(key, []), catalog_new.get(key, []))
        if not changes:
            continue
        section = FileSection(
            file=file_name, tab=tab,
            module=module_of.get(file_name, UNCLASSIFIED),
        )
        section.changes_by_type = _bucket_changes(changes)
        shifted = section.changes_by_type.get("SHIFTED", [])
        section.shift_summary = _build_shift_summary(shifted)
        section.shift_is_uniform = _is_uniform_shift(shifted)
        file_sections.append(section)

    # Sort by (module, file, tab) — also drives the template's groupby('module').
    file_sections.sort(key=lambda s: (s.module or "", s.file, s.tab))

    totals = ScopeTotals(
        tabs=len(file_sections),
        add=sum(len(s.changes_by_type.get("ADDED", [])) for s in file_sections),
        rem=sum(len(s.changes_by_type.get("REMOVED", [])) for s in file_sections),
        mod=sum(
            len(s.changes_by_type.get("MODIFIED", []))
            + len(s.changes_by_type.get("MULTI", []))
            for s in file_sections
        ),
        shift=sum(len(s.changes_by_type.get("SHIFTED", [])) for s in file_sections),
    )
    return ReportContext(
        old_release=old_release, new_release=new_release,
        generated_date=generated_date, file_sections=file_sections, totals=totals,
    )


def _oracle_type_str(f: AlignedField | None) -> str:
    """Oracle-style type string; prefers data_type_raw (keeps CHAR unit)."""
    if f is None or not f.data_type:
        return ""
    if f.data_type_raw:
        return f.data_type_raw
    if f.length is not None and f.scale is not None:
        return f"{f.data_type}({f.length},{f.scale})"
    if f.length is not None:
        return f"{f.data_type}({f.length})"
    return f.data_type


def _bucket_changes(changes: list[Change]) -> dict[str, list[ChangeRow]]:
    """Group classified changes into per-type buckets of ChangeRow view-models."""
    buckets: dict[str, list[ChangeRow]] = defaultdict(list)
    for c in changes:
        primary = c.new_field if c.new_field is not None else c.old_field
        row = ChangeRow(
            change_type=c.change_type,
            field_name=(primary.technical or primary.label or ""),
            label=primary.label or "",
            oracle_type_str=_oracle_type_str(primary),
            old_position=c.old_position,
            new_position=c.new_position,
            required=primary.required,
            axes=c.axes,
            sub_kinds=c.sub_kinds,
            old_label=c.old_field.label if c.old_field else None,
            new_label=c.new_field.label if c.new_field else None,
            old_oracle_type_str=_oracle_type_str(c.old_field) if c.old_field else None,
            new_oracle_type_str=_oracle_type_str(c.new_field) if c.new_field else None,
            old_required=c.old_field.required if c.old_field else None,
            new_required=c.new_field.required if c.new_field else None,
        )
        buckets[c.change_type].append(row)
    return dict(buckets)


def _build_shift_summary(shifted_rows: list[ChangeRow]) -> str | None:
    if not shifted_rows:
        return None
    old_positions = sorted(r.old_position for r in shifted_rows if r.old_position is not None)
    new_positions = sorted(r.new_position for r in shifted_rows if r.new_position is not None)
    if not old_positions or not new_positions:
        return None
    n = len(shifted_rows)
    return (
        f"{n} field{'s' if n != 1 else ''} shifted from positions "
        f"{old_positions[0]}-{old_positions[-1]} to {new_positions[0]}-{new_positions[-1]}."
    )


def _is_uniform_shift(shifted_rows: list[ChangeRow]) -> bool:
    if len(shifted_rows) < 2:
        return False
    deltas = {
        r.new_position - r.old_position
        for r in shifted_rows
        if r.old_position is not None and r.new_position is not None
    }
    if len(deltas) != 1:
        return False
    old_positions = sorted(r.old_position for r in shifted_rows if r.old_position is not None)
    return all(
        old_positions[i + 1] - old_positions[i] == 1
        for i in range(len(old_positions) - 1)
    )


def load_catalog_release(catalog_path: Path, release: str) -> dict[tuple[str, str], list[AlignedField]]:
    """Read one release sheet from the master catalog and group by (file, tab).

    Catalog schema: release | file_name | tab_name | position | column_label |
    column_technical | data_type | length | scale | data_type_raw | required
    """
    wb = load_workbook(catalog_path, read_only=True, data_only=True)
    if release not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Release sheet '{release}' not found in {catalog_path}")
    ws = wb[release]

    grouped: dict[tuple[str, str], list[AlignedField]] = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        _rel, file_name, tab_name, position, label, technical, data_type, length, scale, data_type_raw, required = row
        if file_name is None or tab_name is None:
            continue
        grouped[(file_name, tab_name)].append(AlignedField(
            position=int(position),
            label=label,
            technical=(technical or None),
            data_type=(data_type or None),
            length=(int(length) if length is not None and length != "" else None),
            scale=(int(scale) if scale is not None and scale != "" else None),
            required=_parse_required(required),
            data_type_raw=(str(data_type_raw).strip() if data_type_raw is not None and data_type_raw != "" else None),
        ))

    wb.close()
    for k in grouped:
        grouped[k].sort(key=lambda f: f.position)
    return dict(grouped)


def _parse_required(v) -> bool | None:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    s = str(v).strip().upper()
    if s == "TRUE":
        return True
    if s == "FALSE":
        return False
    return None


def load_file_modules(release: str) -> dict[str, str]:
    """Read baselines/<release>/file_modules.json -> {file_name: module}.

    Returns {} if the file is absent (report groups those files under
    UNCLASSIFIED). Release label is lowercased to match the on-disk dir.
    """
    path = Path("baselines") / release.lower() / "file_modules.json"
    if not path.is_file():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# --- PDF (opt-in) GTK registration on Windows -------------------------------

_GTK_WINDOWS_BIN_CANDIDATES = (
    r"C:\msys64\mingw64\bin",
    r"C:\Program Files\GTK3-Runtime Win64\bin",
    r"C:\Program Files\Gtk-Runtime\bin",
    r"C:\Program Files (x86)\GTK3-Runtime Win64\bin",
)


def _register_windows_gtk_dlls() -> None:
    """On Windows, make a known GTK install dir loadable by weasyprint. No-op elsewhere."""
    import os
    import sys
    if sys.platform != "win32":
        return
    for candidate in _GTK_WINDOWS_BIN_CANDIDATES:
        if Path(candidate, "libgobject-2.0-0.dll").is_file():
            os.add_dll_directory(candidate)
            if candidate not in os.environ.get("PATH", ""):
                os.environ["PATH"] = candidate + os.pathsep + os.environ.get("PATH", "")
            return


def generate_report(
    catalog_path: Path,
    old_release: str,
    new_release: str,
    out_dir: Path,
    pdf: bool = False,
) -> tuple[Path, Path | None]:
    """Load -> build -> render -> write. Returns (html_path, pdf_path|None).

    HTML is always written. PDF is written only when pdf=True (weasyprint +
    GTK imported lazily so the common HTML path has no heavy dependency).
    """
    import jinja2

    catalog_old = load_catalog_release(catalog_path, old_release)
    catalog_new = load_catalog_release(catalog_path, new_release)
    # NEW wins over OLD for module classification.
    module_of = {**load_file_modules(old_release), **load_file_modules(new_release)}

    ctx = build_report_context(
        catalog_old=catalog_old, catalog_new=catalog_new,
        module_of=module_of, old_release=old_release, new_release=new_release,
    )

    template_dir = Path(__file__).parent / "templates"
    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(template_dir),
        autoescape=jinja2.select_autoescape(["html", "j2"]),
        trim_blocks=True, lstrip_blocks=True,
    )
    tpl = env.get_template("report.html.j2")

    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"FBDI_Change_Report_{old_release}_{new_release}"
    html_path = out_dir / f"{base}.html"
    html_path.write_text(tpl.render(ctx=ctx, print_mode=False), encoding="utf-8")

    pdf_path: Path | None = None
    if pdf:
        _register_windows_gtk_dlls()
        import weasyprint
        pdf_path = out_dir / f"{base}.pdf"
        pdf_html = tpl.render(ctx=ctx, print_mode=True)
        weasyprint.HTML(string=pdf_html, base_url=str(template_dir)).write_pdf(str(pdf_path))

    return html_path, pdf_path
```

- [ ] **Step 2: Write the minimal plain template**

Replace the entire contents of `fbdi/templates/report.html.j2` with (Phase 2 redesigns this with the Definian system):

```jinja
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>FBDI Release Change Report — {{ ctx.old_release }} → {{ ctx.new_release }}</title>
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 2rem; color: #1a1a1a; }
  h1 { font-size: 1.6rem; color: #0D2C71; }
  .totals { margin: 1rem 0; }
  .totals span { display: inline-block; margin-right: 1.5rem; font-weight: 600; }
  h2.module { margin-top: 2rem; border-bottom: 2px solid #0D2C71; padding-bottom: 0.25rem; }
  h3.tab { margin-top: 1.25rem; }
  table { border-collapse: collapse; width: 100%; margin: 0.5rem 0 1.25rem; }
  th, td { border: 1px solid #ccc; padding: 4px 8px; text-align: left; font-size: 0.85rem; }
  th { background: #f0f2f7; }
  caption { text-align: left; font-weight: 600; margin-bottom: 0.25rem; }
  .center { text-align: center; }
</style>
</head>
<body>
<h1>FBDI Release Change Report — {{ ctx.old_release }} &rarr; {{ ctx.new_release }}</h1>
<p>Generated {{ ctx.generated_date }}</p>
<div class="totals">
  <span>Tabs changed: {{ ctx.totals.tabs }}</span>
  <span>Added: {{ ctx.totals.add }}</span>
  <span>Removed: {{ ctx.totals.rem }}</span>
  <span>Modified: {{ ctx.totals.mod }}</span>
  <span>Shifted: {{ ctx.totals.shift }}</span>
</div>
{% for module, sections in ctx.file_sections | groupby("module") %}
<h2 class="module">{{ module }}</h2>
  {% for s in sections %}
<h3 class="tab">{{ s.file }} &middot; {{ s.tab }}</h3>
    {% for change_type in ["ADDED", "REMOVED", "MODIFIED", "RENAMED", "MULTI", "SHIFTED"] %}
      {% set rows = s.changes_by_type.get(change_type, []) %}
      {% if rows %}
        {% if change_type == "SHIFTED" and s.shift_is_uniform %}
<p><em>{{ s.shift_summary }}</em> (position shift only — informational, no DB impact).</p>
        {% else %}
<table>
  <caption>{{ change_type | capitalize }} ({{ rows | length }})</caption>
  <thead>
    <tr>
      <th scope="col">Field</th><th scope="col">Label</th><th scope="col">Oracle Type</th>
      <th scope="col">Old Pos</th><th scope="col">New Pos</th><th scope="col">Required</th>
    </tr>
  </thead>
  <tbody>
    {% for r in rows %}
    <tr>
      <td>{{ r.field_name }}</td>
      <td>{{ r.label }}</td>
      <td>{% if change_type in ["MODIFIED", "RENAMED", "MULTI"] %}{{ r.old_oracle_type_str or "—" }} &rarr; {{ r.new_oracle_type_str or "—" }}{% else %}{{ r.oracle_type_str or "—" }}{% endif %}</td>
      <td class="center">{{ r.old_position if r.old_position is not none else "—" }}</td>
      <td class="center">{{ r.new_position if r.new_position is not none else "—" }}</td>
      <td class="center">{% if r.required is none %}—{% else %}{{ r.required | string | upper }}{% endif %}</td>
    </tr>
    {% endfor %}
  </tbody>
</table>
        {% endif %}
      {% endif %}
    {% endfor %}
  {% endfor %}
{% endfor %}
{% if not ctx.file_sections %}
<p>No field-level changes detected between {{ ctx.old_release }} and {{ ctx.new_release }}.</p>
{% endif %}
</body>
</html>
```

- [ ] **Step 3: Write the new test_report.py**

Replace the entire contents of `tests/test_report.py` with:

```python
"""Tests for fbdi.report — Applaud-free release-change view-model + loaders."""

import json
from pathlib import Path

import jinja2
from openpyxl import Workbook

from fbdi.report import (
    FileSection,
    ReportContext,
    build_report_context,
    load_catalog_release,
    load_file_modules,
    _oracle_type_str,
)
from fbdi.align import AlignedField


def _render_report(ctx, print_mode=False):
    template_dir = Path(__file__).parent.parent / "fbdi" / "templates"
    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(template_dir),
        autoescape=jinja2.select_autoescape(["html", "j2"]),
        trim_blocks=True, lstrip_blocks=True,
    )
    return env.get_template("report.html.j2").render(ctx=ctx, print_mode=print_mode)


def _aligned(position, label, technical, data_type=None, length=None, required=None, scale=None):
    return AlignedField(position=position, label=label, technical=technical,
                        data_type=data_type, length=length, scale=scale, required=required)


class TestScope:
    def test_all_changed_tabs_included_without_mapping(self):
        catalog_old = {("F1", "T"): [_aligned(1, "A", "A_F")]}
        catalog_new = {("F1", "T"): [_aligned(1, "A", "A_F"), _aligned(2, "B", "B_F")]}
        ctx = build_report_context(catalog_old, catalog_new, {"F1": "Financials"}, "26A", "26B")
        assert len(ctx.file_sections) == 1
        assert ctx.file_sections[0].module == "Financials"

    def test_unchanged_tab_excluded(self):
        rows = [_aligned(1, "A", "A_F")]
        ctx = build_report_context({("F", "T"): rows}, {("F", "T"): rows}, {}, "26A", "26B")
        assert ctx.file_sections == []

    def test_unknown_file_groups_unclassified(self):
        catalog_old = {("Mystery", "T"): []}
        catalog_new = {("Mystery", "T"): [_aligned(1, "A", "A_F")]}
        ctx = build_report_context(catalog_old, catalog_new, {}, "26A", "26B")
        assert ctx.file_sections[0].module == "Unclassified"


class TestFieldName:
    def test_uses_technical_when_present(self):
        ctx = build_report_context(
            {("F", "T"): []},
            {("F", "T"): [_aligned(1, "Some Label", "FIELD_NAME", "VARCHAR2", 30, True)]},
            {}, "26A", "26B",
        )
        added = ctx.file_sections[0].changes_by_type["ADDED"]
        assert added[0].field_name == "FIELD_NAME"

    def test_falls_back_to_label_when_technical_none(self):
        ctx = build_report_context(
            {("F", "T"): []},
            {("F", "T"): [_aligned(1, "Landed Cost Enabled", None, None, None, True)]},
            {}, "26A", "26B",
        )
        added = ctx.file_sections[0].changes_by_type["ADDED"]
        assert added[0].field_name == "Landed Cost Enabled"


class TestTotals:
    def test_totals_sum_adds_across_sections(self):
        catalog_old = {("F1", "T1"): [], ("F2", "T2"): []}
        catalog_new = {
            ("F1", "T1"): [_aligned(1, "A", "A_F"), _aligned(2, "B", "B_F")],
            ("F2", "T2"): [_aligned(1, "C", "C_F")],
        }
        ctx = build_report_context(catalog_old, catalog_new, {}, "26A", "26B")
        assert ctx.totals.tabs == 2
        assert ctx.totals.add == 3


class TestOracleTypeStr:
    def test_char_unit_preserved(self):
        f = AlignedField(position=1, label="L", technical="T", data_type="VARCHAR2",
                         length=30, scale=None, required=None, data_type_raw="VARCHAR2(30 CHAR)")
        assert _oracle_type_str(f) == "VARCHAR2(30 CHAR)"

    def test_reconstructed_when_raw_absent(self):
        f = AlignedField(position=1, label="L", technical="T", data_type="VARCHAR2",
                         length=30, scale=None, required=None)
        assert _oracle_type_str(f) == "VARCHAR2(30)"


class TestLoaders:
    def test_load_catalog_release_groups_by_file_and_tab(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "26B"
        ws.append(["release", "file_name", "tab_name", "position", "column_label",
                   "column_technical", "data_type", "length", "scale", "data_type_raw", "required"])
        ws.append(["26B", "F1", "T1", 1, "Lab", "TECH", "VARCHAR2", 30, None, "VARCHAR2(30)", "TRUE"])
        ws.append(["26B", "F1", "T1", 2, "Lab2", "TECH2", "NUMBER", 18, None, "NUMBER(18)", "FALSE"])
        path = tmp_path / "cat.xlsx"
        wb.save(path)
        result = load_catalog_release(path, "26B")
        assert len(result[("F1", "T1")]) == 2
        assert result[("F1", "T1")][0].required is True

    def test_load_file_modules_reads_json(self, tmp_path, monkeypatch):
        (tmp_path / "baselines" / "26b").mkdir(parents=True)
        (tmp_path / "baselines" / "26b" / "file_modules.json").write_text(
            json.dumps({"AutoInvoiceImportTemplate.xlsm": "Financials"}))
        monkeypatch.chdir(tmp_path)
        assert load_file_modules("26B") == {"AutoInvoiceImportTemplate.xlsm": "Financials"}

    def test_load_file_modules_missing_returns_empty(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        assert load_file_modules("99Z") == {}


class TestRender:
    def test_renders_title_and_field_without_applaud(self):
        catalog_old = {("F", "T"): []}
        catalog_new = {("F", "T"): [_aligned(1, "Label", "MYFIELD", "VARCHAR2", 30, required=None)]}
        ctx = build_report_context(catalog_old, catalog_new, {"F": "Financials"}, "26A", "26B")
        html = _render_report(ctx)
        assert "FBDI Release Change Report" in html
        assert "MYFIELD" in html
        assert "applaud" not in html.lower()
        assert ">—<" in html  # required=None renders a dash cell

    def test_empty_context_renders_no_changes_message(self):
        ctx = build_report_context({}, {}, {}, "26A", "26B")
        html = _render_report(ctx)
        assert "No field-level changes detected" in html
```

- [ ] **Step 4: Run report tests to verify they pass**

Run: `py -m pytest tests/test_report.py -q`
Expected: PASS (all classes green).

- [ ] **Step 5: Update the `report` CLI command**

In `fbdi/cli.py`, replace the `report_parser` block so it drops `--mapping`, updates help, and adds `--pdf`:

```python
    report_parser = subparsers.add_parser(
        "report",
        help="Generate the FBDI Release Change Report (HTML; --pdf for PDF) from the catalog",
    )
    report_parser.add_argument("--old", required=True, type=str, help="Older release label (e.g. 26B)")
    report_parser.add_argument("--new", required=True, type=str, help="Newer release label (e.g. 26C)")
    report_parser.add_argument("--out-dir", type=Path, default=Path("."), help="Output directory (default: ./)")
    report_parser.add_argument("--catalog", type=Path, default=Path("FBDI_Master_Catalog.xlsx"),
                               help="Path to the master catalog (default: ./FBDI_Master_Catalog.xlsx)")
    report_parser.add_argument("--pdf", action="store_true",
                               help="Also render a PDF (requires weasyprint + GTK; see CLAUDE.md)")
```

And replace `_run_report` with:

```python
def _run_report(args: argparse.Namespace) -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(name)s: %(message)s")

    if not args.catalog.is_file():
        print(f"Error: catalog file not found: {args.catalog}")
        sys.exit(1)

    from fbdi.report import generate_report

    html_path, pdf_path = generate_report(
        catalog_path=args.catalog,
        old_release=args.old.upper(),
        new_release=args.new.upper(),
        out_dir=args.out_dir,
        pdf=args.pdf,
    )
    print(f"HTML: {html_path}")
    if pdf_path is not None:
        print(f"PDF : {pdf_path}")
```

- [ ] **Step 6: Update TestReportSubcommand**

Replace `TestReportSubcommand` in `tests/test_cli.py` with:

```python
class TestReportSubcommand:
    def test_report_subcommand_html_only_by_default(self, monkeypatch, tmp_path):
        from fbdi import cli
        called = {}

        def fake_generate(catalog_path, old_release, new_release, out_dir, pdf=False):
            called.update(dict(catalog_path=catalog_path, old_release=old_release,
                               new_release=new_release, out_dir=out_dir, pdf=pdf))
            return tmp_path / "x.html", None

        (tmp_path / "cat.xlsx").write_bytes(b"stub")
        monkeypatch.setattr("fbdi.report.generate_report", fake_generate)
        cli.main(["report", "--old", "26B", "--new", "26C",
                  "--out-dir", str(tmp_path), "--catalog", str(tmp_path / "cat.xlsx")])
        assert called["old_release"] == "26B"
        assert called["new_release"] == "26C"
        assert called["pdf"] is False

    def test_report_subcommand_pdf_flag(self, monkeypatch, tmp_path):
        from fbdi import cli
        called = {}

        def fake_generate(catalog_path, old_release, new_release, out_dir, pdf=False):
            called["pdf"] = pdf
            return tmp_path / "x.html", tmp_path / "x.pdf"

        (tmp_path / "cat.xlsx").write_bytes(b"stub")
        monkeypatch.setattr("fbdi.report.generate_report", fake_generate)
        cli.main(["report", "--old", "26B", "--new", "26C",
                  "--out-dir", str(tmp_path), "--catalog", str(tmp_path / "cat.xlsx"), "--pdf"])
        assert called["pdf"] is True
```

- [ ] **Step 7: Run the full suite**

Run: `py -m pytest tests/test_report.py tests/test_cli.py -q`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add fbdi/report.py fbdi/templates/report.html.j2 fbdi/cli.py tests/test_report.py tests/test_cli.py
git commit -m "$(cat <<'EOF'
feat(report): repurpose to Applaud-free FBDI release-change report

Scope is now all changed tabs grouped by module from file_modules.json
(no Applaud mapping). Drops Applaud field names/types/30-char flag/
pending-base. PDF is opt-in (--pdf); outputs renamed FBDI_Change_Report_*.
Template reduced to a plain working layout (Definian redesign is Phase 2).

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 3: Delete the Applaud modules, their tests, and the config helper

Now that nothing imports them, remove the Applaud engine wholesale.

**Files:**
- Delete: `fbdi/audit.py`, `fbdi/audit_applaud.py`, `fbdi/applaud_appmap.py`, `fbdi/applaud_snapshot.py`, `fbdi/correspondence.py`, `fbdi/build_mapping.py`, `fbdi/populate_module.py`, `fbdi/applaud_type.py`
- Delete: `tests/test_audit.py`, `tests/test_audit_applaud.py`, `tests/test_correspondence.py`, `tests/test_applaud_appmap.py`, `tests/test_applaud_snapshot.py`, `tests/test_populate_module.py`, `tests/test_module_classifier.py`, `tests/test_applaud_type.py`
- Modify: `fbdi/config.py` — remove `applaud_snapshot_path()` and its comment block (lines ~48-60).

**Interfaces:**
- Produces: `fbdi/config.py` retaining `MAX_FILE_SIZE_BYTES`, `MIN_CELLS`, `SKIP_TABS`, `REPORT_HEADERS`, `CATALOG_TIMEOUT` only.

- [ ] **Step 1: git rm the modules and tests**

```bash
git rm fbdi/audit.py fbdi/audit_applaud.py fbdi/applaud_appmap.py fbdi/applaud_snapshot.py \
       fbdi/correspondence.py fbdi/build_mapping.py fbdi/populate_module.py fbdi/applaud_type.py
git rm tests/test_audit.py tests/test_audit_applaud.py tests/test_correspondence.py \
       tests/test_applaud_appmap.py tests/test_applaud_snapshot.py tests/test_populate_module.py \
       tests/test_module_classifier.py tests/test_applaud_type.py
```

- [ ] **Step 2: Remove `applaud_snapshot_path` from config.py**

Delete the `applaud_snapshot_path` function and the two-line comment above it. Keep everything else in `fbdi/config.py`.

- [ ] **Step 3: Verify no residue + import + suite**

Run:
```bash
grep -ri applaud fbdi/ tools/ tests/
py -c "import fbdi.cli, fbdi.catalog, fbdi.report, fbdi.config"
py -m pytest tests/ -q
```
Expected: grep prints nothing; imports clean; full suite PASSES (test count dropped by the 8 deleted files + the populate-module CLI tests).

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "$(cat <<'EOF'
chore(applaud): remove the FBDI->Applaud engine (superseded by ApplaudMCP)

Deletes 8 modules + 8 test modules + config.applaud_snapshot_path. Nothing
in the core/catalog/report path imports these. catalog_normalize is kept
(it produces the catalog's column_label).

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 4: Delete the Applaud tools

**Files:**
- Delete: `tools/_gen_appmap_full.py`, `tools/assemble_applaud_snapshot.py`, `tools/extract_applaud_snapshot.mjs`

- [ ] **Step 1: git rm the tools**

```bash
git rm tools/_gen_appmap_full.py tools/assemble_applaud_snapshot.py tools/extract_applaud_snapshot.mjs
```

- [ ] **Step 2: Verify**

Run: `ls tools/` — expect only `download_and_clear.py` (+ `__pycache__`). Run `py -m pytest tests/ -q` — PASS.

- [ ] **Step 3: Commit**

```bash
git add -A
git commit -m "$(cat <<'EOF'
chore(applaud): remove Applaud snapshot/app-map tools

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 5: De-clutter the working tree and .gitignore

**Files:**
- Untrack: `FBDI_Master_Catalog.xlsx` (regenerable)
- Delete (tracked): `FBDI_to_ApplaudTables_Mapping.xlsx`, `FBDI_to_Applaud_AppMap.xlsx`
- Delete (on-disk clutter): `Archive/`, `FBDI_Compliance_Report_26A_26B.html`, `FBDI_Compliance_Report_26A_26B.pdf`, `Diagnostic_Report_26B.xlsx`, `Comparison_Report_26A_26B.xlsx`, `Applaud_Compliance_Report_*.xlsx`, `Applaud_FieldMap_Review_*.xlsx`
- Modify: `.gitignore`

- [ ] **Step 1: Untrack the regenerable catalog and remove the Applaud workbooks**

```bash
git rm --cached FBDI_Master_Catalog.xlsx
git rm FBDI_to_ApplaudTables_Mapping.xlsx FBDI_to_Applaud_AppMap.xlsx
```
(`FBDI_Master_Catalog.xlsx` stays on disk — the report task uses it as a smoke-test input.)

- [ ] **Step 2: Delete on-disk clutter**

```bash
rm -rf Archive/
rm -f FBDI_Compliance_Report_26A_26B.html FBDI_Compliance_Report_26A_26B.pdf \
      Diagnostic_Report_26B.xlsx Comparison_Report_26A_26B.xlsx \
      Applaud_Compliance_Report_*.xlsx Applaud_FieldMap_Review_*.xlsx
```

- [ ] **Step 3: Update .gitignore**

In `.gitignore`: add a line `FBDI_Master_Catalog.xlsx`; change `FBDI_Compliance_Report*.html` → `FBDI_Change_Report*.html` and `FBDI_Compliance_Report*.pdf` → `FBDI_Change_Report*.pdf`; delete the Applaud lines (`applaud_table_coverage.xlsx`, the two comment blocks about FBDI_to_Applaud tracked files, `Applaud_Compliance_Report_*.xlsx`, `Applaud_FieldMap_Review_*.xlsx`). Keep `baselines/`, `Archive/`, `Comparison_Report*.xlsx`, `Diagnostic_Report*.xlsx`, `handoff_*.md`, and the `.claude/skills/impeccable/` / `.superpowers/` lines.

- [ ] **Step 4: Verify**

Run: `git status` — the three workbooks staged for deletion/untrack, `.gitignore` modified, and no stray report/Applaud artifacts listed as untracked. Run `py -m pytest tests/ -q` — PASS.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "$(cat <<'EOF'
chore: de-clutter tree; untrack regenerable catalog; drop Applaud workbooks

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 6: Archive the stale Applaud + old-report docs

Preserve blame via `git mv` into `docs/archive/` (repo convention).

**Files:** move the plan/spec/loose docs listed in the spec's Appendix B into `docs/archive/`.

- [ ] **Step 1: git mv the plans and specs**

```bash
cd docs/superpowers
git mv plans/2026-04-21-applaud-mapping-audit.md ../archive/
git mv plans/2026-05-01-fbdi-compliance-report.md ../archive/
git mv plans/2026-05-04-compliance-report-skill-stage9.md ../archive/
git mv plans/2026-05-04-fbdi-run-headless-pipeline.md ../archive/
git mv plans/2026-05-04-fbdi-run-headless-pipeline.png ../archive/
git mv plans/2026-06-02-applaud-compliance-audit.md ../archive/
git mv plans/2026-06-09-applaud-audit-first-run.md ../archive/
git mv plans/2026-06-11-applaud-field-correspondence.md ../archive/
git mv specs/2026-04-21-applaud-mapping-audit-design-revised.md ../archive/
git mv specs/2026-05-01-fbdi-compliance-report-design.md ../archive/
git mv specs/2026-05-04-compliance-report-skill-stage9-design.md ../archive/
git mv specs/2026-05-04-fbdi-run-headless-pipeline-design.md ../archive/
git mv specs/2026-06-02-applaud-compliance-audit-design.md ../archive/
git mv specs/2026-06-09-applaud-audit-first-run-design.md ../archive/
git mv specs/2026-06-10-applaud-field-correspondence-design.md ../archive/
git mv applaud-audit-first-run-notes.md ../archive/
git mv AUDIT_RESULTS_applaud-compliance-audit.md ../archive/
git mv AUDIT_RESULTS_field-correspondence.md ../archive/
git mv AUDIT_RESULTS_field-correspondence_plan_pass1.md ../archive/
git mv AUDIT_RESULTS_field-correspondence_plan_pass2_CLEARED.md ../archive/
git mv AUDIT_RESULTS_plan_pass2.md ../archive/
git mv AUDIT_RESULTS_plan_pass3.md ../archive/
git mv references/applaud-snapshot-extraction.md ../archive/
cd ../..
```
(If `handoff_phase2_rerun.md` is present it is gitignored — `rm -f docs/superpowers/handoff_phase2_rerun.md`.)

- [ ] **Step 2: Verify**

Run: `grep -rl applaud docs/superpowers/` — expect only this Phase-1 plan and the cleanse spec (which mention Applaud in the context of removing it). No orphaned Applaud-only design docs remain outside `docs/archive/`.

- [ ] **Step 3: Commit**

```bash
git add -A
git commit -m "$(cat <<'EOF'
docs(archive): move stale Applaud + old-report specs/plans to docs/archive

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 7: Trim the fbdi-compare-release skill

Remove the populate-module stage and refit the report stage to the new command/output; drop the mapping check from `verify_rerun.py`.

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/SKILL.md`
- Modify: `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py`
- Modify: `tests/test_verify_rerun.py`
- Possibly modify: `.claude/skills/fbdi-compare-release/scripts/summarize_report.py` (Stage 6.5 module-field splice)

- [ ] **Step 1: Read the skill top-to-bottom**

Read `.claude/skills/fbdi-compare-release/SKILL.md` fully so the edits below land in context.

- [ ] **Step 2: Remove Stage 6.5 and refit the description**

- Delete the entire **Stage 6.5 — Populate Module column** section (starts at the `## Stage 6.5` heading, ~line 253, through the end of that stage before the next `## Stage`).
- In the frontmatter `description`, remove the report-only trigger clause that references `FBDI_to_ApplaudTables_Mapping.xlsx`; rewrite the report trigger to: `Also triggers on report-only phrases like 'generate the release change report', 'generate the report for 26B 26C', 'regenerate the PDF' — for these, verify FBDI_Master_Catalog.xlsx exists at repo root, then jump directly to the report stage.`
- Remove Stage 6.5 references in the Stage 7/8 summary text (the `Module column update (Stage 6.5)` block and the `summarize_report.py` note about `populated`/`blank`/`overwritten`).

- [ ] **Step 3: Refit the report stage (was Stage 9 — Compliance Report)**

- Retitle to **Release Change Report**.
- Remove the standalone-preflight requirement for `FBDI_to_ApplaudTables_Mapping.xlsx` (catalog only).
- Change the command to: `py -m fbdi report --old <OLD> --new <NEW>` (add `--pdf` when a PDF is requested).
- Update output names to `FBDI_Change_Report_<OLD>_<NEW>.html` / `.pdf`.
- Remove the "mapping wasn't found" troubleshooting bullet.

- [ ] **Step 4: Drop the mapping check from verify_rerun.py**

In `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py`, remove the `--mapping` argument (lines ~135-136) and any logic that opens/validates the mapping workbook or its "FBDI Mapping" sheet. In `tests/test_verify_rerun.py`, delete the mapping-specific test(s) (e.g. `test_mapping_sheet_missing_distinct_regression`) and any mapping fixtures.

- [ ] **Step 5: Verify**

Run:
```bash
grep -rin "populate\|applaud\|compliance\|_to_Applaud\|--mapping" .claude/skills/fbdi-compare-release/ --include=*.md --include=*.py
py -m pytest tests/test_skill_scripts.py tests/test_verify_rerun.py tests/test_verify_run.py -q
```
Expected: grep returns nothing (except incidental prose); the skill-script tests PASS.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "$(cat <<'EOF'
chore(skill): trim fbdi-compare-release — drop Stage 6.5, refit report stage

Removes the populate-module stage and the mapping check in verify_rerun;
report stage now runs `fbdi report --old --new [--pdf]` -> FBDI_Change_Report_*.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 8: Rewrite CLAUDE.md and the top-level docs

Bring the project docs in line with the cleansed reality.

**Files:**
- Rewrite: `CLAUDE.md`
- Modify: `README.md`, `DESIGN.md`, `PRODUCT.md`
- Modify: `docs/operator-guide.md`, `docs/developer-guide.md`
- Modify: `fbdi/catalog_normalize.py` (docstring only — drop "Applaud MDB" framing)

- [ ] **Step 1: Rewrite CLAUDE.md via the improver skill**

Invoke the `claude-md-management:claude-md-improver` skill. The rewrite must: drop the Applaud "Current Frontier" and all Applaud modules from the pipeline list; drop the Applaud-specific hazards; drop the `populate-module`/`audit-applaud`/`correspondence-*` Quick Start lines; document `py -m fbdi report --old --new [--pdf]` producing `FBDI_Change_Report_*`; correct the test count to the post-cleanse number (run `py -m pytest tests/ -q` and use the actual total); **keep** the still-live hazards (RapidImplementationForCashManagement manual download, phantom columns, corrupt XML, weasyprint/GTK for `--pdf`, the Inter-font note — the Inter note is removed in Phase 2, not here).

- [ ] **Step 2: Update README / DESIGN / PRODUCT**

Remove Applaud framing and the FBDI→Applaud mapping description; describe the pipeline as download → clear → compare → catalog → report.

- [ ] **Step 3: Update the operator and developer guides**

`docs/operator-guide.md`: remove the populate-module step; document the `report --pdf` flow with the new output name. `docs/developer-guide.md`: remove the Applaud modules from the module tour; keep core/catalog/report.

- [ ] **Step 4: Fix catalog_normalize docstring**

In `fbdi/catalog_normalize.py`, change the module/function docstring so it describes general FBDI label normalization (used to produce the catalog's `column_label`) rather than "Applaud MDB compatibility."

- [ ] **Step 5: Verify**

Run:
```bash
grep -rin applaud CLAUDE.md README.md DESIGN.md PRODUCT.md docs/operator-guide.md docs/developer-guide.md fbdi/
py -m pytest tests/ -q
```
Expected: grep returns nothing; suite PASSES.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "$(cat <<'EOF'
docs: rewrite CLAUDE.md + guides for the cleansed core+catalog+report repo

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
EOF
)"
```

---

### Task 9: Final verification and PR

- [ ] **Step 1: Full suite + import + help**

Run:
```bash
py -m pytest tests/ -q
py -m fbdi --help
py -c "import fbdi.cli, fbdi.catalog, fbdi.report, fbdi.diagnose, fbdi.clear, fbdi.compare"
```
Expected: all green; help lists only `compare, diagnose, catalog, report`.

- [ ] **Step 2: Smoke-test the real report**

Run: `py -m fbdi report --old 26A --new 26B` (uses the on-disk `FBDI_Master_Catalog.xlsx`).
Expected: writes `FBDI_Change_Report_26A_26B.html`; open it and confirm module-grouped change tables, no "Applaud" text, and a totals line. (The HTML is gitignored — do not commit it.)

- [ ] **Step 3: Residue sweep**

Run: `grep -rIni applaud . --exclude-dir=.git --exclude-dir=docs/archive --exclude-dir=baselines`
Expected: only this plan + the cleanse spec (which discuss removing Applaud). No live code/doc references.

- [ ] **Step 4: Push and open PR**

```bash
git push -u origin chore/repo-cleanse
gh pr create --base master --title "Repo cleanse: remove Applaud layer + repurpose report" --body "$(cat <<'EOF'
Phase 1 of the repo cleanse (spec: docs/superpowers/specs/2026-09-14-repo-cleanse-and-report-redesign-design.md).

Removes the FBDI->Applaud subsystem (superseded by ApplaudMCP): 8 modules,
3 tools, 8 test modules, 4 CLI subcommands, tracked mapping workbooks, and
the Applaud docs (archived). Repurposes the report into an Applaud-free,
module-grouped FBDI release-change report (HTML default, --pdf opt-in,
output FBDI_Change_Report_*). Rewrites CLAUDE.md + guides + trims the
fbdi-compare-release skill. Full suite green.

Phase 2 (Definian design-system redesign of the report) follows in a
separate PR.

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

- [ ] **Step 5: Wait for review** — do not self-merge (per project workflow).

---

## Self-Review

**Spec coverage:**
- §5.1 remove Applaud layer → Tasks 1, 3, 4. ✓ (`catalog_normalize` kept per its conditional — documented in Global Constraints.)
- §5.2 decouple report → Task 2. ✓ (scope=all changed tabs, module from file_modules.json, --pdf, rename).
- §5.3 de-clutter → Task 5. ✓
- §5.4 docs archive → Task 6; guides → Task 8. ✓
- §5.5 trim skill → Task 7. ✓
- §5.6 CLAUDE.md + top-level docs → Task 8. ✓
- §5.7 acceptance criteria → Task 9. ✓

**Placeholder scan:** No TBD/TODO; every code step has full content; edit steps name exact files/anchors and give exact replacement text.

**Type consistency:** `build_report_context(catalog_old, catalog_new, module_of, old_release, new_release, generated_date=None)`, `generate_report(catalog_path, old_release, new_release, out_dir, pdf=False)`, `load_file_modules(release)`, and the `FileSection`/`ChangeRow` fields (`field_name`, `module`, no `applaud_*`) are used identically in report.py, the template, the tests, and the cli. The cli fake in TestReportSubcommand matches the new `generate_report` signature (`pdf` kwarg). ✓

**Note on divergence from spec default:** the spec's §5.1 gave a default of *removing* `catalog_normalize`; on inspection it produces the catalog's `column_label`, so the spec's own conditional ("if nothing non-Applaud needs it") resolves to **keep**. Recorded in Global Constraints and Task 3.
