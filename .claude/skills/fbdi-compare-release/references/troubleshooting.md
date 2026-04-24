# FBDI Compare-Release — Troubleshooting

Load this only when a stage fails or the user reports an anomaly. The
symptoms below are the ones documented in `CLAUDE.md` and
`docs/scraper-gap-findings-2026-04-23.md`; anything outside this list is
an unknown — escalate to Brad rather than guessing.

---

## Stage 3 — Download ran but produced zero files / short counts

**Symptom:** `scripts/verify_download.py` reports many missing files from a
single Oracle module URL (e.g., all procurement files absent), and
`download_and_clear.py` logs show `Navigating to ...` immediately followed
by `Completed: ...` with no "Downloading:" lines between them.

**Cause:** Silent module-page failure in Oracle's JET-rendered
navigation drawer. The `navigationDrawer` element loads, but section
expansion yields no children before the scraper harvests links. Observed
on 2026-04-23 (first run on Brad's Windows machine, 26B); resolved on
retry.

**Fix:** Retry `download_and_clear.py <ver>` once (per spec §5 #5, auto).
If the retry still shorts, `MODULE_URL_TEMPLATES` in
`tools/download_and_clear.py:44-49` may need updating — Oracle may have
restructured the docs URL pattern for that module. Do not attempt to
debug from within the skill; defer to Brad.

**Retry cap:** 3 total download attempts per run. After the 3rd
verification failure, option (a) "retry again" is no longer offered — only
abort or proceed-with-gaps remain.

---

## Stage 3 — `RapidImplementationForCashManagement.xlsm` is missing

**Cause:** This is an Oracle Fusion FSM (Functional Setup Manager)
template, not a standard FBDI template hosted on docs pages. The scraper
cannot fetch it — it must be obtained manually from Oracle Fusion.

**Walk-through (paste verbatim into the user prompt):**

1. Log into Oracle Fusion.
2. Navigate: **Setup and Maintenance** → click the **hamburger menu**
   (top-right) → **Search**.
3. Search for: `Create Banks, Branches, and Accounts in Spreadsheet`.
4. Click the task — the template downloads directly.
5. Place it in `baselines/<new_release>/originals/`.

**Fast alternative:** copy from the prior release's `originals/` folder.
Oracle rarely updates this template, and `sha256sum` confirms the 26A and
26B copies were bit-identical on 2026-04-23.

---

## Stage 4 — Per-file clear timed out

**Symptom:** `download_and_clear.py --clear-only` prints:
```
*** TIMED OUT (N files, >120s each) — clear these manually: ***
    PayablesCollectionDocuments.xlsm (9,xxxKB)
```

**Cause:** Large xlsm files (typically >8MB) can exceed the 120s per-file
subprocess timeout in `tools/download_and_clear.py:244-281`. Known
recurring offender on 2026-04-23: `PayablesCollectionDocuments.xlsm`
(~9MB, timed out in both 26A and 26B).

**Impact:** **Not a blocker for Stage 5.** The comparison engine reads
`originals/` (not `blanks/`). The timeout only affects whether the
`blanks/<ver>/` folder has a cleared copy of that file for downstream
client use.

**Fix (optional):** Clear the file manually — either in Excel (select all
data rows below the header, delete, save) or via the legacy VBA macro
`reference/Clear_FBDIs - 20210412.xlsm`.

**Flag in final summary:** SKILL.md Stage 7 surfaces the timeout filename
list in the terminal summary so the user knows what needs manual
clearing.

---

## Stage 5 — Per-pair compare failure

**Symptom:** Compare output includes a file pair where one side's xlsm
cannot be opened (corrupt xml, phantom `max_column=16384`, etc.).

**Cause:** Oracle occasionally ships xlsm files with corrupt metadata.
`compare.py` uses subprocess-per-pair isolation with a 120s timeout; a
failure here does not abort the run. Historically ~11 FILE_ERROR files
exist in 26B.

**Fix:** None available from within the skill. Failures are collected and
surfaced at end of Stage 5 per spec §5 #4. If >5 failures, the skill
pauses for user input (retry / skip / abort).

---

## Stage 6 — Catalog Issues tab jumped

**Symptom:** `verify_run.py` reports `catalog_issues.regression: true`
— the new release has either >2× the prior release's Issues count, or
>50 more absolute.

**Likely causes:**
- Oracle changed how data-type strings are encoded in a new release
  (e.g., new temporal format mask not covered by `fbdi/type_parser.py`).
- A large new tab was added with no header and produces `NO_HEADER` rows.

**Fix:** Inspect the Issues tab for the new release — if all new entries
are `TYPE_PARSE_WARNING` with a common pattern, extend
`fbdi/type_parser.py` (see the 2026-04-20 fix: temporal format masks,
trailing-period typos). This is `fbdi/`-package work, not skill work —
defer to Brad.

**Does not block:** the regression is flagged in the summary; the user
decides whether to defer.

---

## Selenium / Chrome

**`chromedriver` version mismatch:** `webdriver-manager` auto-installs a
matching chromedriver for the installed Chrome. If it fails, upgrade
Chrome to the latest stable and re-run.

**Chrome not found:** `check_env.py` reports `"chrome": {"ok": false}`.
Install Chrome from https://www.google.com/chrome/.
