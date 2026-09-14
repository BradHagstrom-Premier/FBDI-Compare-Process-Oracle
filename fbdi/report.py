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
    # NEW wins over OLD for module classification. Catalog file_name values are
    # extension-less (e.g. "AutoInvoiceImportTemplate") while file_modules.json
    # keys carry the ".xlsm" extension — reconcile on the stem so the lookup hits.
    raw_modules = {**load_file_modules(old_release), **load_file_modules(new_release)}
    module_of = {Path(k).stem: v for k, v in raw_modules.items()}

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
