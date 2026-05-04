# Operator Guide: Quarterly FBDI Refresh

> If you're looking to develop on this codebase rather than run the pipeline, start with [`developer-guide.md`](developer-guide.md).

## What this is and who it's for

Every quarter, Oracle ships a new release of its FBDI (File-Based Data Import) template files. This pipeline compares the new release against the previous one, field by field, across all templates. It produces three deliverables:

1. A comparison report showing exactly what changed (`Comparison_Report_<OLD>_<NEW>.xlsx`).
2. An updated master catalog of every field in every template for the new release (`FBDI_Master_Catalog.xlsx`).
3. An HTML and PDF compliance report that's safe to hand to clients or auditors (`FBDI_Compliance_Report_<OLD>_<NEW>.html` and `.pdf`).

Definian team members who maintain Oracle integrations use these outputs to keep those integrations aligned with Oracle's changes. This guide is for whoever is running that quarterly refresh. No Python knowledge required.

---

## Before your first run

**Environment.** You need Python 3.14 or higher and Google Chrome installed on your machine. The pipeline runs on Windows. PDF rendering needs MSYS2 mingw64 GTK installed; the install steps are in CLAUDE.md under Known Hazards. Before your first run, open a terminal and run:

```
pip install -r requirements.txt
```

That installs the Python packages the pipeline depends on.

**Where files land.** Downloads go into a folder called `baselines/<release>/originals/`, like `baselines/26B/originals/`. These folders aren't tracked by Git (they're gitignored), so they won't interfere with source control. If you use the Claude Code skill to run the pipeline, it creates these folders for you. If you run the CLI commands by hand, you'll need to create the folder yourself before starting.

**Windows sleep warning.** The download step uses a browser automation tool that needs Chrome to stay open for 15–20 minutes. If your laptop goes to sleep or locks the screen mid-run, Chrome loses its connection and the download fails. Before starting, disable Windows sleep: Settings, System, Power, Screen and sleep, then set to "Never" for the duration. Or plan to stay at your desk. Running overnight with sleep disabled also works fine.

**Time budget.** A full run takes 35–50 minutes end to end. Downloads account for most of that. The compare and catalog steps are each under 5 minutes. The compliance report takes another 5 to 15 seconds.

**Two ways to run it.** The recommended path, especially for your first run, is to open Claude Code and say something like "Oracle released 26C, run the FBDI refresh." That invokes the `fbdi-compare-release` skill, which walks you through each stage with human-in-the-loop checkpoints at every decision point. The alternative is to run the CLI commands for each stage directly in a terminal. That's faster once you know what you're doing, but you lose the checkpoint prompts and the automatic error handling.

---

## The 9 stages, in order

(Plus a Stage 6.5 between catalog and summary that handles the FBDI-to-Applaud mapping update.)

### Stage 1: Environment preflight

**What it does.** Checks that your machine has everything the pipeline needs before doing any real work. It verifies Python version, Chrome presence, and that the required packages are installed.

**What you see on screen.** A short pass/fail report. Green means proceed. If something fails, the script tells you exactly what's wrong.

**Expected wall time.** Under 30 seconds.

**If it stalls.** It shouldn't. If the script hangs, kill it and check that Python is on your PATH.

---

### Stage 2: Resolve OLD and NEW releases

**What it does.** Figures out which release you're comparing from and which you're comparing to. If you told the skill explicitly (for example, "compare 26A to 26B"), it uses those. Otherwise it looks at what's already in `baselines/` and infers the previous release, then infers the new release as the next quarter. The naming convention after 26D is 27A.

**What you see on screen.** A confirmation of the two release versions it's going to use, before anything destructive happens.

**Expected wall time.** A few seconds.

**If it stalls.** It shouldn't. If something looks wrong with the version detection, tell the skill explicitly which versions to use.

---

### Stage 3: Download + verify

**What it does.** Downloads all FBDI template files from Oracle's documentation site for the new release (and the old release too, if it's not already on disk). It uses browser automation to scrape and download each file, then verifies the download against a known inventory. If files are missing after the first attempt, it retries once automatically.

**What you see on screen.** Download progress as files are fetched, then a verification summary showing how many files were expected, how many were found, and any gaps. Filenames scroll by for 15–20 minutes.

**Expected wall time.** 15–20 minutes per release being downloaded.

**If it stalls.** If the browser freezes, kill the process and re-invoke the skill. It will re-download. Note: re-running Stage 3 wipes `originals/` first, so don't restart it partway through if you're trying to preserve a partial download.

---

### Stage 4: Smart-clear

**What it does.** Creates "blank" copies of the new release's templates in `baselines/<NEW>/blanks/`. Blank means headers are preserved but the data rows are removed. These are the clean template files you'd hand to a client. The clearing logic handles headers at different row positions across different templates.

**What you see on screen.** A file-by-file progress log. Occasionally you'll see a `*** TIMED OUT ***` message for a particular file. That just means the file took longer than 2 minutes to clear and was skipped. Those files need to be cleared manually later.

**Expected wall time.** 2–4 minutes.

**If it stalls.** The overall stage won't stall. Individual file timeouts are handled gracefully and don't block the rest. If the whole stage hangs, kill it and re-run. It's safe to repeat.

---

### Stage 5: Compare

**What it does.** Compares the old release and new release template by template, tab by tab, field by field. For each pair of matching files, it diffs the column headers and identifies what was added, removed, or modified. Results go into `Comparison_Report_<OLD>_<NEW>.xlsx`.

**What you see on screen.** A running count of file pairs being processed (around 210 pairs total), with any per-file errors or warnings noted inline. Occasional warnings about corrupt file metadata are normal and handled automatically.

**Expected wall time.** 3–5 minutes.

**If it stalls.** Individual file pairs run in isolated subprocesses, so a single bad file won't freeze the whole run. If the entire stage stalls past 10 minutes, kill it and re-run. It's safe to repeat.

---

### Stage 6: Catalog update

**What it does.** Builds or updates the FBDI master catalog for the new release. The catalog is a comprehensive inventory of every field in every template tab: position, label, technical column name, data type, length, scale, and whether it's required. This is separate from the comparison report; it's a snapshot of the new release on its own.

**What you see on screen.** Progress by file and tab. A summary at the end shows how many rows were written to the catalog.

**Expected wall time.** 3–5 minutes.

**If it stalls.** Same as Stage 5. Per-file isolation means a single bad file won't freeze the stage. Kill and re-run if the whole thing hangs.

---

### Stage 6.5: Populate Module column in mapping spreadsheet

**What it does.** Updates the Module column in `FBDI_to_ApplaudTables_Mapping.xlsx` based on the per-release `file_modules.json` written by the downloader. The pipeline asks you whether to back up the mapping file first (HITL #7); the default is yes. If `FBDI_to_ApplaudTables_Mapping.xlsx` isn't present at the repo root, this stage is skipped silently and the run continues.

**What you see on screen.** A short JSON-style summary showing how many cells were populated, how many stayed blank, and how many were overwritten.

**Expected wall time.** A few seconds.

**If it stalls.** It shouldn't. If `file_modules.json` is missing for either release, that means Stage 3 didn't complete cleanly. The skill halts and tells you which release is missing the file.

---

### Stage 7: Summary

**What it does.** Reads the comparison report and catalog and renders a human-readable summary to your terminal: total change rows, which files changed the most, the Module column update results from Stage 6.5, and any files that timed out during smart-clear and still need manual attention.

**What you see on screen.** A formatted text block. Something like: "748 change rows across 47 files. Top 5 most-changed: ..." followed by the Stage 6.5 summary and a list of any files from Stage 4 that need manual clearing.

**Expected wall time.** Under 10 seconds.

**If it stalls.** It shouldn't. If the summary script errors out, the comparison report and catalog are already on disk. The summary is just a convenience view of those.

---

### Stage 8: Post-run verification

**What it does.** Runs a quick sanity check comparing this run's outcomes against historical baselines. It checks whether the header detection success rate dropped, whether the catalog Issues tab grew significantly, and whether any new file errors appeared that weren't there in the prior release. If anything looks off, it appends a warning block to the summary.

**What you see on screen.** Either nothing extra (clean run) or a WARNING block listing specific metrics that regressed. A warning here doesn't mean the outputs are wrong. It means something is worth a second look.

**Expected wall time.** Under 10 seconds.

**If it stalls.** It shouldn't. If it errors out, the comparison report and catalog are unaffected. This stage is read-only.

---

### Stage 9: Compliance report

**What it does.** Asks you to validate the Excel outputs (HITL #8), then generates the formal HTML and PDF compliance report to hand to clients or auditors. This is the polished deliverable. It pulls from the master catalog and the FBDI-to-Applaud mapping, filters to the in-scope mapped tabs, and renders a single document in two formats.

**What you see on screen.** A short prompt asking you to spot-check the comparison report and catalog for anything obviously wrong (unreasonable change counts, blank columns, missing files). When you confirm, the script writes `FBDI_Compliance_Report_<OLD>_<NEW>.html` and `.pdf` to the repo root. If you say "skip," the run ends without generating the report; you can always come back and run `python -m fbdi report` later, or re-trigger the skill with a phrase like "generate the compliance report for 26A 26B."

**Expected wall time.** 5 to 15 seconds.

**If it stalls.** PDF generation is the part that fails most often, and the failure mode is loud. If you see a Pango or libgobject error, MSYS2 mingw64 GTK isn't installed correctly. The HTML file is usually written even when the PDF step fails, so check for that first.

---

## The 8 HITL checkpoints

The pipeline has eight human-in-the-loop checkpoints where it stops and asks you what to do. The labels HITL #1 through #8 are stable IDs from the design spec, not the order they appear during a run. HITL #3, for instance, fires before HITL #1 because version resolution happens before baseline presence is checked.

### HITL #1: Prior release missing

**Trigger.** The pipeline needs to compare against an older release, but `baselines/<OLD>/originals/` is empty or doesn't exist.

**Options.**

- Download the old release too (runs Stage 3 twice).
- Point the pipeline at an existing copy somewhere else on disk.

**How to decide.** If you've run a prior quarter's refresh and the `baselines/` folder is intact, confirm the path. If it's genuinely missing, download it. That adds another 15–20 minutes, but it's the clean path.

---

### HITL #2: `RapidImplementationForCashManagement.xlsm` missing

**Trigger.** After the download completes, this specific file isn't in `baselines/<NEW>/originals/`. It's a bank account setup template that Oracle doesn't host on the standard FBDI documentation pages, so the downloader never finds it.

**Options.**

- Copy it from the prior release's folder (fast and safe; Oracle rarely updates this template).
- Walk through the Oracle Fusion FSM manual download path.
- Drop it into the folder yourself and tell the pipeline to continue.

**How to decide.** Copying from the prior release is the right call in almost every case. Only go to Oracle Fusion if you have a specific reason to believe this template changed in the new release.

---

### HITL #3: Version-mismatch sanity check

**Trigger.** You passed explicit version numbers (or the pipeline auto-detected them), but the detected "old" release doesn't match what you asked for. For example, the pipeline sees 25D as the most recent baseline, but you said you want to compare 26A to 26C.

**Options.**

- Confirm you want to skip releases (unusual, but the run proceeds as asked).
- Correct to the detected version.

**How to decide.** Skipping releases is rarely intentional. If you're catching up after a missed quarter, confirm explicitly. If this prompt was unexpected, you probably mistyped a version number.

---

### HITL #4: Excessive compare failures

**Trigger.** The comparison step reports more than 5 per-pair file failures.

**Options.**

- Retry the comparison (some failures are transient).
- Skip the failed pairs and note the gap in the summary.
- Abort.

**How to decide.** A handful of failures (under 5) is normal and never triggers this prompt. The pipeline silently handles them. If you see this prompt, check whether the failures are concentrated in one or two large files or spread across many. A retry usually clears transient issues; if the same files fail again, proceed with the gap noted.

---

### HITL #5: Download still short after retry

**Trigger.** After the automatic retry, files in the expected inventory still didn't download.

**Options.**

- Retry again (available up to 3 total download attempts).
- Abort so the download script can be debugged directly.
- Proceed with the gap acknowledged (not recommended; the comparison output will be incomplete for those files).

**How to decide.** Try one more retry if the Oracle docs site might have been temporarily slow. If it fails a third time, abort and look at what's happening with the specific module pages that are missing. The Oracle docs site occasionally restructures its navigation between releases, which breaks the scraper for that module.

---

### HITL #6: Extra files present

**Trigger.** The downloader pulled files that aren't in the known inventory for this release.

**Options.**

- Add the extra files to the inventory (the right call when Oracle has added new templates).
- Quarantine them to a separate folder and keep them out of the comparison.
- Review them file by file before deciding.

**How to decide.** Adding the extras to the inventory is almost always correct. Oracle does add new templates in quarterly releases. Quarantine is only appropriate if you suspect the scraper accidentally pulled something from the wrong page, which is rare.

---

### HITL #7: Backup before mapping update

**Trigger.** Stage 6.5 is about to update the Module column in `FBDI_to_ApplaudTables_Mapping.xlsx`.

**Options.**

- Yes, copy to `FBDI_to_ApplaudTables_Mapping.bak.xlsx` (default).
- No, just go (the file is git-tracked, so you can revert).

**How to decide.** Take the backup. It's free, the file is small, and if something looks wrong after the run you have a clean copy to compare against. The default is yes for a reason.

---

### HITL #8: Compliance report validation gate

**Trigger.** Stage 9 is about to generate the compliance report. The skill wants you to spot-check the Excel outputs first so the report doesn't bake in obvious errors.

**Options.**

- Yes, generate the HTML and PDF.
- Skip (you can run `python -m fbdi report` later, or trigger the skill with a phrase like "generate compliance report for 26A 26B").

**How to decide.** Open `Comparison_Report_<OLD>_<NEW>.xlsx` and the new tab in `FBDI_Master_Catalog.xlsx`. Does the total change count look plausible for a quarterly Oracle release? Are there any obviously wrong files? Does the catalog have a reasonable row count and no blank columns? If anything jumps out, skip the report, fix the underlying issue, and re-run. Bad inputs produce a bad report, and the report is what gets sent to clients.

---

## Reading the outputs

You get four output files after a successful run, all in the repo root.

**`Comparison_Report_<OLD>_<NEW>.xlsx`** is the change log. It has seven columns: File, Tab, Position, Label, Technical, Change Type, and Details. Change Type is one of Added, Removed, or Modified. Each row represents a single field-level change in a single template tab. This is the file you hand to whoever owns the Applaud mapping. They use it to identify which integrations need updating before the new release goes live. When you're triaging, sort by File or Tab to group changes by area, or sort by Change Type to pull out all Removed fields first (those are highest risk for existing integrations).

**`FBDI_Master_Catalog.xlsx`** is the full snapshot. It has three kinds of sheets:

- Per-release tabs (such as `26A`, `26B`) with one row per file × tab × column, including position, label, technical name, data type, length, scale, and required flag.
- An `Issues` sheet flagging fields where Oracle's type strings are genuinely malformed. There are 9 rows currently, all known Oracle data-quality issues, not a sign that something went wrong in your run.
- A `Drift` sheet tracking how tabs moved across releases (added, removed, shifted, renamed, modified).

The catalog is useful when someone asks "what fields does this template have?" for any release. It's the reference you reach for before opening an actual xlsm file.

**`FBDI_Compliance_Report_<OLD>_<NEW>.html`** and **`FBDI_Compliance_Report_<OLD>_<NEW>.pdf`** are the polished deliverables. Same content in two formats. The HTML version has collapsible sections, which is the better choice for browsing changes interactively. The PDF is the formal print-rendered version, suitable for attaching to an email or dropping into an audit folder. Both are generated from the master catalog and the FBDI-to-Applaud mapping, so they only show changes that are actually in scope for Definian's integrations.

---

## When something goes sideways

**The FSM file is missing.** If Stage 3 warns that `RapidImplementationForCashManagement.xlsm` isn't in the download folder, don't ignore it. HITL #2 will guide you through it. The manual fetch path in Oracle Fusion: Setup and Maintenance, click the hamburger menu in the upper right, Search, type "Create Banks, Branches, and Accounts in Spreadsheet", click the task name to trigger the download. Drop the file into `baselines/<NEW>/originals/` and continue the run.

**You hit Ctrl-C mid-run.** That's fine. The pipeline is designed to resume. From Stage 3 onward, re-invoking the skill picks up where you left off. Files already on disk aren't re-downloaded, and the compare and catalog steps overwrite their outputs cleanly. One exception: if you Ctrl-C partway through Stage 3, re-running it will wipe `originals/` and start fresh. The skill warns you before doing this.

**The compliance report PDF won't generate.** If Stage 9 prints a Pango or libgobject error, MSYS2 mingw64 GTK isn't set up correctly. CLAUDE.md has the install steps under Known Hazards. The HTML file is usually written even when the PDF step fails, so check for `FBDI_Compliance_Report_<OLD>_<NEW>.html` in the repo root first. You can also re-run just the report step later with `python -m fbdi report --old <OLD> --new <NEW>`, or trigger the skill with a phrase like "regenerate the PDF for 26A 26B."

**For anything else.** The skill's built-in error handling covers the common cases. It'll never dump a raw Python error at you without also giving you a plain-English explanation and a choice of what to do next. If you're seeing something the skill doesn't handle gracefully, or if Stage 8 verification surfaces a regression you don't recognize, check the "Known hazards" section in `CLAUDE.md` first. If the catalog Issues count jumps significantly between releases, Stage 8 will call it out specifically.

---

## Next steps

If you need to understand what's happening inside any of these stages, extend a stage to handle a new edge case, or fix a bug you've run into, start with [`developer-guide.md`](developer-guide.md). That guide covers the Python package structure, how the comparison and catalog engines work, and how to run the test suite.
