"""Build the patched vbaProject.bin for RapidImplementationForCashManagement.xlsm.

One-time helper, re-run only when Oracle changes the template's macro (the
refresh then fails with a BlankPatchError naming this script). The quarterly
refresh itself needs no Excel: fbdi.blank_patches swaps the committed asset in.

What it does:
  1. Copies the Oracle original to a temp folder (bytes only, no zone marker).
  2. Opens it in a private, hidden Excel instance (macros enabled so the
     ActiveX button loads; refused if the template has any auto-run VBA).
  3. In the "Bank Account" sheet module, replaces the Save As ``With fd ...
     End With`` block (which hard-codes ``.FilterIndex = 5 'for xml`` — "CSV
     UTF-8" in current Excel) with one that picks the XML filter by extension,
     opens in the workbook's folder, and forces a .xml extension.
  4. Saves, extracts xl/vbaProject.bin into fbdi/assets/, verifies it with
     olevba, and prints the original's sha256 for
     fbdi.blank_patches.RAPID_IMPL_ORIGINAL_VBA_SHA256.

Requires Windows + Excel + pywin32, and Excel's "Trust access to the VBA
project object model" (File > Options > Trust Center > Trust Center Settings >
Macro Settings; registry HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Security
AccessVBOM=1). ActiveX must also be allowed to load (Trust Center > ActiveX
Settings: any option except "Disable all controls without notification").
With ActiveX disabled Excel never instantiates CommandButton1 and silently
drops ``Attribute VB_Control = "CommandButton1, ..."`` from the sheet module on
save — the click handler would no longer be bound to the button. The script
checks that the button loads and that no binding was lost, and refuses to
write the asset otherwise.

Usage:
    py tools/build_vba_patch.py baselines/26C/originals/RapidImplementationForCashManagement.xlsm
"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from fbdi.blank_patches import PATCHED_VBA_ASSET, RAPID_IMPL_SHEET, VBA_PART  # noqa: E402

# Macros must be *enabled* while Excel holds the file: with them force-disabled
# Excel doesn't load the ActiveX button, and on save it drops the module's
# ``Attribute VB_Control = "CommandButton1, ..."`` binding. Enabling is safe only
# because the template has no auto-run code — assert_no_auto_exec() checks that
# before the file is opened.
MSO_AUTOMATION_SECURITY_LOW = 1

# Replacement for the original ``With fd ... End With`` block, plus the .xml
# extension enforcement that follows it. Tested by hand on a live engagement.
NEW_BLOCK = r"""    With fd     'Configure dialog box
        'Definian 2026-11-05: open in this file's folder; pick the XML type by name, not position
        '(position 5 is CSV UTF-8 in current Excel)
        .InitialFileName = ThisWorkbook.Path & "\CE_RAPID_IMPLEMENTATION.xml"
        Dim k As Long
        On Error Resume Next
        For k = 1 To .Filters.Count
            If InStr(1, .Filters(k).Extensions, "*.xml", vbTextCompare) > 0 Then .FilterIndex = k: Exit For
        Next k
        On Error GoTo 0
        .Show
        If (.SelectedItems.Count = 0) Then
            gsFilePath = "False"
        Else
            gsFilePath = .SelectedItems.Item(1)
        End If
    End With

    'Definian 2026-11-05: the content is always XML, so force the .xml extension
    If gsFilePath <> "False" Then
        Dim nm As String, p As Long
        p = InStrRev(gsFilePath, "\")
        nm = Mid(gsFilePath, p + 1)
        If InStr(nm, ".") > 0 Then nm = Left(nm, InStr(nm, ".") - 1)
        gsFilePath = Left(gsFilePath, p) & nm & ".xml"
    End If""".splitlines()


def find_save_as_block(lines: list[str]) -> tuple[int, int]:
    """0-based (start, end) inclusive line indexes of the original With fd block."""
    starts = [i for i, l in enumerate(lines) if re.match(r"\s*With fd\b", l)]
    if len(starts) != 1:
        raise SystemExit(f"expected exactly one 'With fd' block, found {len(starts)}")
    start = starts[0]
    end = next(
        (i for i in range(start, len(lines)) if re.match(r"\s*End With\b", lines[i])), None
    )
    if end is None:
        raise SystemExit("no 'End With' after 'With fd'")
    block = "\n".join(lines[start:end + 1])
    if not re.search(r"\.FilterIndex\s*=\s*5\b", block):
        raise SystemExit(
            "the With fd block has no '.FilterIndex = 5' — the macro changed; "
            "review it by hand before patching:\n" + block
        )
    return start, end


def vba_modules(vba_bin: bytes) -> dict[str, str]:
    """{module filename: source} extracted with olevba."""
    from oletools.olevba import VBA_Parser

    parser = VBA_Parser("vbaProject.bin", data=vba_bin)
    try:
        return {vba_name: code.replace("\r\n", "\n") for _, _, vba_name, code in parser.extract_macros()}
    finally:
        parser.close()


def assert_no_auto_exec(vba_bin: bytes) -> None:
    """Refuse to open the template with macros enabled if anything auto-runs.

    olevba also reports ActiveX ``*_Click`` handlers as AutoExec; those need a
    user click, which never happens in a hidden automation session, so they
    are allowed. Anything else (Workbook_Open, Auto_Open, *_Activate,
    *_GotFocus, *_Layout, ...) is refused.
    """
    from oletools.olevba import VBA_Parser

    parser = VBA_Parser("vbaProject.bin", data=vba_bin)
    try:
        auto = [
            kw for kind, kw, _ in parser.analyze_macros()
            if kind == "AutoExec" and not kw.endswith("_Click")
        ]
    finally:
        parser.close()
    if auto:
        raise SystemExit(
            f"template has auto-run VBA ({', '.join(auto)}); refusing to open it with "
            "macros enabled. Review the macro and patch by hand."
        )


def _vb_controls(modules: dict[str, str]) -> set[str]:
    return {
        f"{name}: {line.strip()}"
        for name, code in modules.items()
        for line in code.splitlines()
        if line.startswith("Attribute VB_Control")
    }


def build(original: Path, out: Path) -> str:
    import pywintypes
    import win32com.client

    data = original.read_bytes()
    original_vba = zipfile.ZipFile(original).read(VBA_PART)
    original_sha = hashlib.sha256(original_vba).hexdigest()
    assert_no_auto_exec(original_vba)

    tmpdir = Path(tempfile.mkdtemp(prefix="vba_patch_"))
    work = tmpdir / original.name
    work.write_bytes(data)  # bytes only: no Mark-of-the-Web / Protected View

    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        wb = excel.Workbooks.Open(str(work))
        try:
            sheet = wb.Worksheets(RAPID_IMPL_SHEET)
            for i in range(1, sheet.OLEObjects().Count + 1):
                ole = sheet.OLEObjects().Item(i)
                try:
                    ole.Object  # raises if Excel refused to load the control
                except pywintypes.com_error:
                    raise SystemExit(
                        f"Excel did not load ActiveX control {ole.Name!r}. Set Trust "
                        "Center > ActiveX Settings to anything except 'Disable all "
                        "controls without notification' and re-run."
                    )
            code_name = sheet.CodeName
            try:
                module = wb.VBProject.VBComponents(code_name).CodeModule
            except pywintypes.com_error as e:
                raise SystemExit(
                    "Excel refused VBA project access. Enable File > Options > Trust "
                    "Center > Trust Center Settings > Macro Settings > 'Trust access "
                    f"to the VBA project object model' and re-run. ({e})"
                )
            lines = module.Lines(1, module.CountOfLines).splitlines()
            start, end = find_save_as_block(lines)
            module.DeleteLines(start + 1, end - start + 1)
            module.InsertLines(start + 1, "\r\n".join(NEW_BLOCK))
            wb.Save()
        finally:
            wb.Close(False)
    finally:
        excel.Quit()

    patched = zipfile.ZipFile(work).read(VBA_PART)
    verify(original_vba, patched)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(patched)
    return original_sha


def verify(original_vba: bytes, patched_vba: bytes) -> None:
    """The fix is present, and nothing else lost: same modules, same control bindings."""
    before, after = vba_modules(original_vba), vba_modules(patched_vba)
    if set(before) != set(after):
        raise SystemExit(f"module set changed: {sorted(before)} -> {sorted(after)}")
    lost = _vb_controls(before) - _vb_controls(after)
    if lost:
        raise SystemExit(f"Excel dropped ActiveX bindings on save: {sorted(lost)}")
    code = "\n".join(after.values())
    if re.search(r"FilterIndex\s*=\s*5", code) or ".Filters(k).Extensions" not in code:
        raise SystemExit("verification failed: patched VBA does not contain the fix")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("original", type=Path, help="Oracle-shipped RapidImplementationForCashManagement.xlsm")
    ap.add_argument("--out", type=Path, default=PATCHED_VBA_ASSET)
    args = ap.parse_args()

    sha = build(args.original.resolve(), args.out)
    print(f"Wrote {args.out} ({args.out.stat().st_size:,} bytes)")
    print(f"Original {VBA_PART} sha256: {sha}")
    print("Set fbdi.blank_patches.RAPID_IMPL_ORIGINAL_VBA_SHA256 to that value if it changed.")


if __name__ == "__main__":
    main()
