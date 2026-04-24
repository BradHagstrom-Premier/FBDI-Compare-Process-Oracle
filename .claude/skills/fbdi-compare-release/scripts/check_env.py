"""Stage 1 preflight for fbdi-compare-release.

Checks: OS, Python >= 3.14, required deps importable, Chrome installed,
baselines/ exists, baseline_files.txt present.

JSON stdout. Exit codes: 0=ok, 1=fatal, 2=deps-only-missing.
"""

from __future__ import annotations

import importlib
import json
import os
import platform
import shutil
import sys
from pathlib import Path

REQUIRED_DEPS = ["openpyxl", "selenium", "webdriver_manager", "requests", "pytest"]
MIN_PYTHON = (3, 14)


def check_python_version(current=None) -> dict:
    current = current or sys.version_info[:3]
    ok = tuple(current[:2]) >= MIN_PYTHON
    detail = f"{current[0]}.{current[1]}.{current[2]} (need >= 3.14)"
    return {"name": "python_version", "ok": ok, "detail": detail}


def check_os() -> dict:
    system = platform.system()
    if system == "Windows":
        return {"name": "os", "ok": True, "detail": platform.platform()}
    return {
        "name": "os",
        "ok": True,  # non-fatal warning
        "detail": f"{system} (Windows is the supported platform; proceeding anyway)",
    }


def check_deps(required=None) -> dict:
    required = required or REQUIRED_DEPS
    missing = []
    for name in required:
        try:
            importlib.import_module(name)
        except ImportError:
            missing.append(name)
    if missing:
        return {
            "name": "deps",
            "ok": False,
            "detail": f"missing: {', '.join(missing)}",
            "missing": missing,
        }
    return {"name": "deps", "ok": True, "detail": "all required deps importable"}


def check_chrome() -> dict:
    # Windows default locations
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return {"name": "chrome", "ok": True, "detail": path}
    # Fall back to PATH lookup (covers Mac/Linux dev machines)
    for exe in ("chrome", "google-chrome", "chromium"):
        found = shutil.which(exe)
        if found:
            return {"name": "chrome", "ok": True, "detail": found}
    return {
        "name": "chrome",
        "ok": False,
        "detail": "Google Chrome not found. Install from https://www.google.com/chrome/",
    }


def check_baselines_dir(root: Path) -> dict:
    baselines = root / "baselines"
    if baselines.is_dir():
        return {"name": "baselines_dir", "ok": True, "detail": "baselines/ exists"}
    baselines.mkdir(parents=True, exist_ok=True)
    return {"name": "baselines_dir", "ok": True, "detail": "created baselines/"}


def check_baseline_files_txt(root: Path) -> dict:
    path = root / "baseline_files.txt"
    if path.is_file():
        return {"name": "baseline_files_txt", "ok": True, "detail": "present"}
    return {
        "name": "baseline_files_txt",
        "ok": False,  # non-fatal — caller decides
        "detail": "baseline_files.txt not found; download verification will be limited",
    }


def main(argv=None) -> int:
    root = Path.cwd()
    checks = [
        check_python_version(),
        check_os(),
        check_deps(),
        check_chrome(),
        check_baselines_dir(root),
        check_baseline_files_txt(root),
    ]

    fatal = []
    missing_deps: list[str] = []
    for c in checks:
        if not c["ok"]:
            if c["name"] == "deps":
                missing_deps = c.get("missing", [])
            elif c["name"] == "baseline_files_txt":
                pass  # non-fatal warning
            else:
                fatal.append(c["name"])

    payload = {
        "ok": not fatal and not missing_deps,
        "checks": checks,
        "missing_deps": missing_deps,
        "fatal": fatal,
    }
    print(json.dumps(payload, indent=2))

    if fatal:
        return 1
    if missing_deps:
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
