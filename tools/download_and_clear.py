"""
Download Oracle FBDI templates for a given release version and create
blank (cleared) copies with headers preserved.

Discovery is static HTML, no browser: every Oracle docs book publishes a
`toc.htm` that lists all of its topic pages, and each topic page links its
template as an absolute `.../fbdi-<ver>/fbdi/xlsm/<File>.xlsm` href. (The old
Selenium scraper clicked through the JET `<oj-tree-view>` and silently missed
pages, e.g. 26D RequirementImportTemplate.)

The run is loud about shortfalls: a TOC/topic page that can't be fetched, a
download that fails, or a downloaded file that is gone from disk afterwards
(SentinelOne quarantines some templates seconds after they're written) is
named, and the script exits 1. The skill's Stage 4.5 reconciliation gate is
still the authoritative inventory check.

Usage:
    py tools/download_and_clear.py 26D
    py tools/download_and_clear.py 26D --discover-only   # list what Oracle publishes; write nothing
    py tools/download_and_clear.py 26D --skip-clear      # download only
    py tools/download_and_clear.py 26D --clear-only      # clear only (already downloaded)
    py tools/download_and_clear.py 26D --use-vba-macro   # use Clear_FBDIs VBA instead of Python

Exit codes: 0 = clean; 1 = discovery errors, failed downloads, or files
missing from originals/ after the run.
"""

import argparse
import json
import multiprocessing
import os
import re
import shutil
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from urllib.parse import unquote, urljoin, urlparse

# Add project root to sys.path so subprocesses can import fbdi
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ---------------------------------------------------------------------------
# Oracle docs books -- swap the version token (e.g. 26a, 26b). Order matters:
# when a template is linked from more than one book, the FIRST book wins its
# module in file_modules.json (e.g. ChangeOrderImportTemplate -> SCM).
# ---------------------------------------------------------------------------
BOOK_TOC_TEMPLATES = [
    "https://docs.oracle.com/en/cloud/saas/project-management/{ver}/oefpp/toc.htm",
    "https://docs.oracle.com/en/cloud/saas/financials/{ver}/oefbf/toc.htm",
    "https://docs.oracle.com/en/cloud/saas/procurement/{ver}/oefbp/toc.htm",
    "https://docs.oracle.com/en/cloud/saas/supply-chain-and-manufacturing/{ver}/oefsc/toc.htm",
]

# Where Oracle hosts every template (the same href the docs pages link to).
# Used to print a recovery URL for a file that went missing.
WEBFOLDER_URL = "https://www.oracle.com/webfolder/technetwork/docs/fbdi-{ver}/fbdi/xlsm/{name}"

# Seconds to wait after the last download before checking the files are still
# on disk. SentinelOne deleted 26D's quarantined templates within seconds.
SETTLE_SECONDS = 15

# Files that cannot be auto-downloaded from Oracle docs pages.
# They are distributed only through Oracle Fusion's Functional Setup Manager UI.
# After running this script, place them manually in baselines/<VER>/originals/.
#
# RapidImplementationForCashManagement.xlsm:
#   Setup and Maintenance → click the hamburger menu (top-right) → Search
#   → search "Create Banks, Branches, and Accounts in Spreadsheet"
#   → click the task — it downloads the template directly
MANUAL_FILES = [
    "RapidImplementationForCashManagement.xlsm",
]


# Maps URL slug → canonical module name. Keep keys aligned with
# BOOK_TOC_TEMPLATES; values are the module taxonomy fbdi/report.py groups by
# (`&` not "and" for SCM).
URL_TO_MODULE = {
    "project-management": "Project Management",
    "financials": "Financials",
    "procurement": "Procurement",
    "supply-chain-and-manufacturing": "Supply Chain & Manufacturing",
}


def module_from_base_url(url: str) -> str:
    """Extract the Oracle module name from a docs-book URL.

    Example: 'https://docs.oracle.com/en/cloud/saas/financials/26b/oefbf/toc.htm'
             → 'Financials'

    Raises ValueError for URLs that don't match any known module slug.
    The `/saas/<slug>/` guard avoids false positives where the slug
    happens to appear elsewhere in the URL (e.g., as a query param).
    """
    for slug, module in URL_TO_MODULE.items():
        if f"/saas/{slug}/" in url:
            return module
    raise ValueError(f"Unknown Oracle module URL: {url}")


def write_module_map(file_modules: dict, version: str, baselines_root: str) -> str:
    """Write {filename: module} to baselines/<VER>/file_modules.json.

    Adds the hardcoded entry for the FSM-distributed file
    RapidImplementationForCashManagement.xlsm (manually placed by user, so
    never discovered). Returns the absolute path written.
    """
    file_modules = dict(file_modules)
    file_modules.setdefault("RapidImplementationForCashManagement.xlsm", "Financials")
    out_path = os.path.join(baselines_root, version.upper(), "file_modules.json")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(file_modules, f, indent=2, sort_keys=True)
    return out_path


# ---------------------------------------------------------------------------
# Discovery (pure functions + an injected `fetch`, so tests need no network)
# ---------------------------------------------------------------------------
_HREF_RE = re.compile(r"""href\s*=\s*["']([^"']+)["']""", re.IGNORECASE)


def parse_toc(html: str, toc_url: str) -> list[str]:
    """Topic-page URLs listed in a book's toc.htm, in order, deduplicated.

    Keeps only .htm/.html pages inside the book's own directory; drops
    fragments, the book's index/toc pages, stylesheets and external links.
    """
    book_dir = toc_url.rsplit("/", 1)[0] + "/"
    out: list[str] = []
    seen: set[str] = set()
    for href in _HREF_RE.findall(html):
        url = urljoin(toc_url, href.strip()).split("#", 1)[0]
        if not url.startswith(book_dir):
            continue
        page = url[len(book_dir):]
        if not page.lower().endswith((".htm", ".html")):
            continue
        if page.lower() in ("index.html", "index.htm", "toc.htm", "toc.html"):
            continue
        if url not in seen:
            seen.add(url)
            out.append(url)
    return out


def extract_xlsm_links(html: str, page_url: str) -> list[str]:
    """Absolute URLs of every .xlsm href on a topic page, in order, deduplicated."""
    out: list[str] = []
    for href in _HREF_RE.findall(html):
        url = urljoin(page_url, href.strip()).split("#", 1)[0]
        if urlparse(url).path.lower().endswith(".xlsm") and url not in out:
            out.append(url)
    return out


def xlsm_filename(url: str) -> str:
    return unquote(os.path.basename(urlparse(url).path))


@dataclass
class Discovery:
    files: dict = field(default_factory=dict)    # filename -> (url, module)
    errors: list = field(default_factory=list)   # human-readable, one per failed fetch
    pages: int = 0                               # topic pages scanned


def discover(version: str, fetch, workers: int = 8) -> Discovery:
    """Find every template the release's docs books link to.

    `fetch(url) -> str` returns page HTML or raises. A TOC or topic page that
    can't be fetched is recorded in Discovery.errors -- never silently skipped.
    """
    ver = version.lower()
    result = Discovery()

    def _safe(url):
        try:
            return url, fetch(url), None
        except Exception as e:  # noqa: BLE001 - reported, not swallowed
            return url, None, f"{type(e).__name__}: {e}"

    for template in BOOK_TOC_TEMPLATES:
        toc_url = template.format(ver=ver)
        module = module_from_base_url(toc_url)
        _, toc_html, err = _safe(toc_url)
        if err:
            result.errors.append(f"TOC {toc_url}: {err}")
            continue
        topics = parse_toc(toc_html, toc_url)
        if not topics:
            result.errors.append(f"TOC {toc_url}: no topic pages found")
            continue
        with ThreadPoolExecutor(max_workers=workers) as pool:
            for page_url, html, err in pool.map(_safe, topics):  # map keeps TOC order
                result.pages += 1
                if err:
                    result.errors.append(f"page {page_url}: {err}")
                    continue
                for link in extract_xlsm_links(html, page_url):
                    result.files.setdefault(xlsm_filename(link), (link, module))
    return result


def find_missing(expected_names, folder: str) -> list[str]:
    """Expected filenames that are absent (or empty) in `folder`."""
    missing = []
    for name in sorted(expected_names):
        path = os.path.join(folder, name)
        if not os.path.isfile(path) or os.path.getsize(path) == 0:
            missing.append(name)
    return missing


def make_session(retries: int = 5, backoff_factor: float = 1) -> requests.Session:
    retry = Retry(
        total=retries,
        read=retries,
        connect=retries,
        backoff_factor=backoff_factor,
        status_forcelist=(429, 500, 502, 503, 504),
    )
    adapter = HTTPAdapter(max_retries=retry, pool_maxsize=16)
    session = requests.Session()
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def make_fetch(session: requests.Session, timeout=(5, 30)):
    def fetch(url: str) -> str:
        response = session.get(url, timeout=timeout)
        response.raise_for_status()
        return response.text
    return fetch


def download_all(files: dict, dest: str, session: requests.Session, timeout=(5, 60)) -> list[tuple[str, str]]:
    """Download each discovered (url, module) into `dest`. Returns [(name, error)].

    Writes to `<name>.part` and renames on completion, so an interrupted
    download never leaves a truncated .xlsm behind.
    """
    failures = []
    for i, (name, (url, _module)) in enumerate(sorted(files.items()), 1):
        print(f"  [{i}/{len(files)}] {name}")
        final = os.path.join(dest, name)
        part = final + ".part"
        try:
            with session.get(url, stream=True, timeout=timeout) as response:
                response.raise_for_status()
                with open(part, "wb") as f:
                    for chunk in response.iter_content(chunk_size=65536):
                        f.write(chunk)
            os.replace(part, final)
        except (requests.exceptions.RequestException, OSError) as e:
            failures.append((name, f"{type(e).__name__}: {e}"))
            print(f"      FAILED: {e}")
            if os.path.exists(part):
                try:
                    os.unlink(part)
                except OSError:
                    pass
    return failures


def report_missing(missing: list[str], version: str, folder: str) -> None:
    print(f"\n  *** {len(missing)} downloaded file(s) are NOT in {folder}: ***")
    for name in missing:
        print(f"      {name}")
        print(f"        recover from: {WEBFOLDER_URL.format(ver=version.lower(), name=name)}")
    print("  A file that downloaded and then vanished is most likely an AV quarantine")
    print("  (SentinelOne: tray app -> Quarantined Files; only IT can release it).")
    print("  See CLAUDE.md Known Hazards.")


def run_discover_only(version: str, originals_path: str) -> int:
    """Print what Oracle publishes for `version` vs. what's in originals/. Writes nothing."""
    print(f"Discovering templates for {version} (static toc.htm scan)...")
    disc = discover(version, make_fetch(make_session()))
    by_module: dict[str, int] = {}
    for _url, module in disc.files.values():
        by_module[module] = by_module.get(module, 0) + 1
    print(f"  {disc.pages} topic pages scanned, {len(disc.files)} templates found")
    for module, n in sorted(by_module.items()):
        print(f"    {module}: {n}")
    for err in disc.errors:
        print(f"  ERROR {err}")

    not_discovered: list[str] = []
    if os.path.isdir(originals_path):
        on_disk = {f for f in os.listdir(originals_path) if f.endswith(".xlsm") and not f.startswith("~$")}
        not_discovered = sorted(on_disk - set(disc.files) - set(MANUAL_FILES))
        new = sorted(set(disc.files) - on_disk)
        print(f"\n  vs {originals_path} ({len(on_disk)} files):")
        print(f"    on disk but not discovered ({len(not_discovered)}): {not_discovered}")
        print(f"    discovered but not on disk ({len(new)}): {new}")
    return 1 if (disc.errors or not_discovered) else 0


def create_folder(path, clear=True):
    """Create a folder. If clear=True, wipe contents first."""
    if os.path.exists(path):
        if clear:
            for name in os.listdir(path):
                fp = os.path.join(path, name)
                try:
                    if os.path.isfile(fp) or os.path.islink(fp):
                        os.unlink(fp)
                    elif os.path.isdir(fp):
                        shutil.rmtree(fp)
                except Exception as e:
                    print(f"  Failed to delete {fp}: {e}")
    else:
        os.makedirs(path)


def clean_temp_files(folder):
    """Remove Excel temp/lock files (~$*.xls*) that break batch processing."""
    removed = 0
    for f in os.listdir(folder):
        if f.startswith("~$") and ".xls" in f:
            try:
                os.unlink(os.path.join(folder, f))
                removed += 1
            except PermissionError:
                print(f"  Warning: could not delete {f} (locked by Excel -- close Excel and retry)")

    if removed:
        print(f"  Cleaned {removed} temp file(s)")
    return removed


def _clear_single_file(src, dst):
    """Worker: clear one FBDI file. Runs in a subprocess for timeout support."""
    from fbdi.clear import clear_workbook
    clear_workbook(src, dst)


def clear_files_python(originals_path, blanks_path, timeout=120):
    """
    Clear FBDI templates using smart header detection.

    For each file in originals_path:
      - Detect the header row per sheet using detect_header_row()
      - Clear all cell values below the header row
      - Save to blanks_path with the same filename (no BLANK_ prefix)

    Each file is processed in a subprocess with a timeout so that huge
    files don't block the entire batch.
    """
    files = sorted(f for f in os.listdir(originals_path)
                   if f.endswith((".xlsm", ".xlsx")) and not f.startswith("~$"))

    total = len(files)
    cleared = 0
    skipped = 0
    failed = []
    slow_files = []
    timed_out = []

    for i, filename in enumerate(files, 1):
        src = os.path.join(originals_path, filename)
        dst = os.path.join(blanks_path, filename)

        if os.path.exists(dst):
            skipped += 1
            cleared += 1
            continue

        size_kb = os.path.getsize(src) // 1024
        print(f"  [{i}/{total}] {filename} ({size_kb:,}KB) ... ", end="", flush=True)

        start = time.time()
        proc = multiprocessing.Process(target=_clear_single_file, args=(src, dst))
        proc.start()
        proc.join(timeout=timeout)

        if proc.is_alive():
            proc.terminate()
            proc.join(5)
            elapsed = time.time() - start
            timed_out.append((filename, size_kb))
            print(f"TIMEOUT after {elapsed:.0f}s — skipped")
            if os.path.exists(dst):
                os.unlink(dst)
        elif proc.exitcode != 0:
            elapsed = time.time() - start
            failed.append((filename, f"subprocess exit code {proc.exitcode}"))
            print(f"FAILED ({elapsed:.1f}s)")
            if os.path.exists(dst):
                os.unlink(dst)
        else:
            elapsed = time.time() - start
            cleared += 1
            if elapsed > 30:
                slow_files.append((filename, elapsed, size_kb))
            print(f"done ({elapsed:.1f}s)")

    # --- Summary ---
    print(f"\n  Results: {cleared}/{total} cleared", end="")
    if skipped:
        print(f" ({skipped} already existed)", end="")
    print()

    if timed_out:
        print(f"\n  *** TIMED OUT ({len(timed_out)} files, >{timeout}s each) — clear these manually: ***")
        for name, kb in timed_out:
            print(f"      {name} ({kb:,}KB)")

    if slow_files:
        print(f"\n  Slow files (>{30}s but completed):")
        for name, secs, kb in slow_files:
            print(f"      {name} ({kb:,}KB) — {secs:.0f}s")

    if failed:
        print(f"\n  Failed ({len(failed)}):")
        for name, err in failed:
            print(f"      {name}: {err}")

    return cleared, failed + [(n, f"timeout >{timeout}s") for n, _ in timed_out]


def run_clear_macro(originals_path, blanks_path, macro_workbook):
    """
    Run the Clear_FBDIs VBA macro in batch mode via xlwings.
    Fallback option -- use --use-vba-macro flag to enable.

    Sets "File Path Configuration" sheet cells:
      B2 = source folder (trailing backslash)
      C2 = "" (empty = batch mode, processes all .xls* files)
      B3 = destination folder (trailing backslash)
    """
    import xlwings as xw

    src = os.path.abspath(originals_path).rstrip("\\") + "\\"
    dst = os.path.abspath(blanks_path).rstrip("\\") + "\\"
    macro_path = os.path.abspath(macro_workbook)

    print(f"\nRunning Clear_FBDIs macro (batch mode)")
    print(f"  Source:      {src}")
    print(f"  Destination: {dst}")

    with xw.App(visible=False) as app:
        wb = app.books.open(macro_path)
        ws = wb.sheets["File Path Configuration"]

        ws.range("B2").value = src
        ws.range("C2").value = ""
        ws.range("B3").value = dst

        wb.macro("Sheet1.CommandButton1_Click")()

        wb.save()
        wb.close()

    print("  Clear macro complete.")


def main():
    parser = argparse.ArgumentParser(
        description="Download Oracle FBDI templates and create blank copies"
    )
    parser.add_argument(
        "version",
        help="Oracle release version (e.g. 26a, 26b)",
    )
    parser.add_argument(
        "--discover-only",
        action="store_true",
        help="List the templates Oracle's docs link for this release and diff "
             "against baselines/<VER>/originals/; download/write nothing",
    )
    parser.add_argument(
        "--skip-clear",
        action="store_true",
        help="Skip the clearing step (download only)",
    )
    parser.add_argument(
        "--clear-only",
        action="store_true",
        help="Skip downloading, just clear existing files in originals/",
    )
    parser.add_argument(
        "--use-vba-macro",
        action="store_true",
        help="Use Clear_FBDIs VBA macro instead of Python clearing (requires Excel)",
    )
    parser.add_argument(
        "--clear-macro",
        default="Clear_FBDIs - 20210412.xlsm",
        help="Path to Clear_FBDIs workbook (only with --use-vba-macro)",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=120,
        help="Per-file timeout in seconds for Python clearing (default: 120)",
    )
    parser.add_argument(
        "--baselines-root",
        default=None,
        help="Directory holding <VER>/originals and <VER>/blanks "
             "(default: <repo>/baselines)",
    )
    args = parser.parse_args()

    version = args.version.upper()

    # Resolve paths relative to the repo root (parent of tools/)
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    baselines_root = os.path.abspath(args.baselines_root or os.path.join(repo_root, "baselines"))
    originals_path = os.path.join(baselines_root, version, "originals")
    blanks_path = os.path.join(baselines_root, version, "blanks")

    if args.discover_only:
        sys.exit(run_discover_only(version, originals_path))

    print(f"=== FBDI Download & Clear: {version} ===")
    print(f"  Originals: {originals_path}")
    print(f"  Blanks:    {blanks_path}")

    problems: list[str] = []
    expected: list[str] = []  # files this run downloaded; re-checked at the end

    # --- Download ---
    if not args.clear_only:
        create_folder(originals_path)
        if not args.skip_clear:
            create_folder(blanks_path)

        print(f"\nDiscovering templates for {version} (static toc.htm scan)...")
        session = make_session()
        disc = discover(version, make_fetch(session))
        print(f"  {disc.pages} topic pages scanned, {len(disc.files)} templates found")
        for err in disc.errors:
            print(f"  ERROR {err}")
        problems += [f"discovery: {e}" for e in disc.errors]
        if not disc.files:
            print("ERROR: discovery found no templates -- nothing downloaded.")
            sys.exit(1)

        # The module map comes from discovery, so every template Oracle links
        # gets a module even if its download fails or the file is quarantined.
        modules_path = write_module_map(
            {name: module for name, (_url, module) in disc.files.items()},
            version, baselines_root,
        )
        print(f"Wrote module map for {len(disc.files)} files to {modules_path}")

        print(f"\nDownloading {len(disc.files)} templates...")
        failures = download_all(disc.files, originals_path, session)
        problems += [f"download {name}: {err}" for name, err in failures]

        expected = sorted(set(disc.files) - {name for name, _ in failures})
        print(f"\nWaiting {SETTLE_SECONDS}s, then checking every download is still on disk...")
        time.sleep(SETTLE_SECONDS)
        missing = find_missing(expected, originals_path)
        if missing:
            report_missing(missing, version, originals_path)
    else:
        os.makedirs(blanks_path, exist_ok=True)

    # Count downloads
    clean_temp_files(originals_path)
    downloaded = [f for f in os.listdir(originals_path)
                  if f.endswith((".xlsm", ".xlsx")) and not f.startswith("~$")]
    print(f"\n{len(downloaded)} .xlsm files in {originals_path}")

    if not downloaded:
        print("No files found -- nothing to clear.")
        sys.exit(1)

    # --- Clear ---
    if args.skip_clear:
        print("Skipping clear step (--skip-clear).")
    elif args.use_vba_macro:
        macro_path = os.path.join(repo_root, args.clear_macro)
        if not os.path.exists(macro_path):
            print(f"ERROR: Clear macro not found at {macro_path}")
            sys.exit(1)
        run_clear_macro(originals_path, blanks_path, macro_path)
        cleared = [f for f in os.listdir(blanks_path) if f.endswith((".xlsm", ".xlsx"))]
        print(f"Created {len(cleared)} blank copies in {blanks_path}")
    else:
        print(f"\nClearing templates (Python, timeout={args.timeout}s per file)...")
        clear_files_python(originals_path, blanks_path, timeout=args.timeout)

    # Warn about files that must be placed manually (not auto-downloadable)
    missing_manual = [f for f in MANUAL_FILES if f not in os.listdir(originals_path)]
    if missing_manual:
        print(f"\n  *** MANUAL FILES REQUIRED — place these in {originals_path}: ***")
        for name in missing_manual:
            print(f"      {name}")
        print("  See MANUAL_FILES comment in this script for download instructions.")

    # Final re-check: quarantine can land minutes later, after the clear.
    if expected:
        missing = find_missing(expected, originals_path)
        if missing:
            report_missing(missing, version, originals_path)
            problems += [f"missing after download: {name}" for name in missing]

    if problems:
        print(f"\n=== {version}: {len(problems)} problem(s) -- exit 1 ===")
        for p in problems:
            print(f"  - {p}")
        sys.exit(1)
    print(f"\n=== Done: {version} ===")


if __name__ == "__main__":
    main()
