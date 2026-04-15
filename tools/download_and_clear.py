"""
Download Oracle FBDI templates for a given release version and create
blank (cleared) copies with headers preserved.

Usage:
    python tools/download_and_clear.py 26a
    python tools/download_and_clear.py 26b
    python tools/download_and_clear.py 26a --skip-clear       # download only
    python tools/download_and_clear.py 26a --clear-only        # clear only (already downloaded)
    python tools/download_and_clear.py 26a --use-vba-macro     # use Clear_FBDIs VBA instead of Python
"""

import argparse
import glob
import multiprocessing
import os
import shutil
import sys
import time

import requests
from requests.adapters import HTTPAdapter
from selenium import webdriver
from selenium.common.exceptions import (
    ElementNotInteractableException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from urllib3.util.retry import Retry
from webdriver_manager.chrome import ChromeDriverManager


# ---------------------------------------------------------------------------
# Oracle doc URL patterns -- swap the version token (e.g. 26a, 26b)
# ---------------------------------------------------------------------------
MODULE_URL_TEMPLATES = [
    "https://docs.oracle.com/en/cloud/saas/project-management/{ver}/oefpp/index.html",
    "https://docs.oracle.com/en/cloud/saas/financials/{ver}/oefbf/index.html",
    "https://docs.oracle.com/en/cloud/saas/procurement/{ver}/oefbp/index.html",
    "https://docs.oracle.com/en/cloud/saas/supply-chain-and-manufacturing/{ver}/oefsc/index.html",
]


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


def download_with_retry(session, url, timeout=(5, 15), retries=5, backoff_factor=1):
    retry = Retry(
        total=retries,
        read=retries,
        connect=retries,
        backoff_factor=backoff_factor,
        status_forcelist=(500, 502, 504),
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    response = session.get(url, stream=True, timeout=timeout)
    response.raise_for_status()
    return response


def download_files(driver, download_path, version):
    """Scrape Oracle docs for the given version and download all .xlsm files."""
    base_urls = [t.format(ver=version) for t in MODULE_URL_TEMPLATES]
    session = requests.Session()
    seen_filenames = set(os.listdir(download_path))  # skip already-downloaded files

    for base_url in base_urls:
        print(f"\nNavigating to {base_url}")
        driver.get(base_url)

        # Oracle JET pages can be slow to render -- wait with retry
        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "navigationDrawer"))
            )
        except TimeoutException:
            print(f"  navigationDrawer not found after 60s, refreshing...")
            driver.refresh()
            time.sleep(5)
            try:
                WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.ID, "navigationDrawer"))
                )
            except TimeoutException:
                print(f"  SKIP: Could not load navigation for {base_url}")
                continue

        section_items = driver.find_elements(By.CSS_SELECTOR, "#navigationDrawer li")

        for section in section_items:
            try:
                # Expand collapsible sections
                try:
                    expand_icon = section.find_element(
                        By.CSS_SELECTOR, ".oj-clickable-icon-nocontext"
                    )
                    driver.execute_script("arguments[0].click();", expand_icon)
                    time.sleep(1)
                except Exception:
                    pass

                links = section.find_elements(By.CSS_SELECTOR, "a")
                for link in links:
                    link_text = link.text
                    try:
                        driver.execute_script("arguments[0].click();", link)
                        time.sleep(1)

                        try:
                            WebDriverWait(driver, 10).until(
                                EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "a[href$='.xlsm']")
                                )
                            )
                            download_links = driver.find_elements(
                                By.CSS_SELECTOR, "a[href$='.xlsm']"
                            )
                            for dl in download_links:
                                download_url = dl.get_attribute("href")
                                if download_url and download_url.endswith(".xlsm"):
                                    local_filename = download_url.split("/")[-1]

                                    # Skip duplicates (same file linked from multiple sections)
                                    if local_filename in seen_filenames:
                                        continue
                                    seen_filenames.add(local_filename)

                                    print(f"  Downloading: {local_filename}")
                                    try:
                                        for cookie in driver.get_cookies():
                                            session.cookies.set(
                                                cookie["name"], cookie["value"]
                                            )
                                        response = download_with_retry(
                                            session, download_url
                                        )
                                        with open(
                                            os.path.join(download_path, local_filename),
                                            "wb",
                                        ) as f:
                                            for chunk in response.iter_content(
                                                chunk_size=8192
                                            ):
                                                f.write(chunk)
                                        time.sleep(1)
                                    except requests.exceptions.RequestException as e:
                                        print(
                                            f"  Failed to download {download_url}: {e}"
                                        )
                        except TimeoutException:
                            pass  # No .xlsm links in this section
                    except Exception as e:
                        print(f"  Failed to process link '{link_text}': {e}")
            except Exception as e:
                print(f"  Error in section: {e}")

        print(f"Completed: {base_url}")


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
        "--skip-clear",
        action="store_true",
        help="Skip the clearing step (download only)",
    )
    parser.add_argument(
        "--clear-only",
        action="store_true",
        help="Skip downloading, just clear existing files in Originals/",
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
    args = parser.parse_args()

    version = args.version.upper()
    version_lower = args.version.lower()

    # Resolve paths relative to the repo root (parent of tools/)
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    originals_path = os.path.join(repo_root, "baselines", version, "originals")
    blanks_path = os.path.join(repo_root, "baselines", version, "blanks")

    print(f"=== FBDI Download & Clear: {version} ===")
    print(f"  Originals: {originals_path}")
    print(f"  Blanks:    {blanks_path}")

    # --- Download ---
    if not args.clear_only:
        create_folder(originals_path)
        if not args.skip_clear:
            create_folder(blanks_path)

        print(f"\nStarting Selenium for version {version_lower}...")
        chrome_options = Options()
        chrome_options.add_argument("--ignore-ssl-errors=yes")
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--window-size=1920,1080")
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_options,
        )

        try:
            download_files(driver, originals_path, version_lower)
        finally:
            driver.quit()
    else:
        os.makedirs(blanks_path, exist_ok=True)

    # Count downloads
    clean_temp_files(originals_path)
    downloaded = [f for f in os.listdir(originals_path)
                  if f.endswith((".xlsm", ".xlsx")) and not f.startswith("~$")]
    print(f"\n{len(downloaded)} .xlsm files in {originals_path}")

    if not downloaded:
        print("No files found -- nothing to clear.")
        return

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
        cleared_count, failed = clear_files_python(originals_path, blanks_path, timeout=args.timeout)

    print(f"\n=== Done: {version} ===")


if __name__ == "__main__":
    main()
