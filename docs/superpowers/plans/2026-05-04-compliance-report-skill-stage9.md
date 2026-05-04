# Compliance Report Skill Stage 9 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.
>
> **REQUIRED FIRST ACTION:** Invoke `skill-creator:skill-creator` via the Skill tool before touching any files. This is mandatory — not optional, not substitutable with `superpowers:writing-skills` or `superpowers:executing-plans`. The spec and this plan both call this out explicitly because the session may have been cleared since brainstorming.

**Goal:** Add `## Stage 9 — Compliance Report` to `.claude/skills/fbdi-compare-release/SKILL.md`, placing a HITL validation gate after Stage 8 that pauses for human Excel review before running `python -m fbdi report`.

**Architecture:** Two surgical edits to a single markdown file — update the YAML frontmatter description to add report-only trigger phrases, then append the Stage 9 block at the end of the file. No Python code changes. No new files.

**Tech Stack:** Markdown, YAML frontmatter, `skill-creator:skill-creator` skill for guided edits.

---

## File Map

| Action | Path | What changes |
|---|---|---|
| Modify | `.claude/skills/fbdi-compare-release/SKILL.md` | Frontmatter description + Stage 9 block appended |

No other files are created or modified.

---

### Task 1: Invoke `skill-creator:skill-creator`

**Files:**
- (none — skill invocation only)

- [ ] **Step 1: Invoke the skill**

  Use the Skill tool to invoke `skill-creator:skill-creator`. This skill understands how to edit existing `SKILL.md` files. Follow its guidance for the edits described in Tasks 2 and 3.

  ```
  Skill({ "skill": "skill-creator:skill-creator" })
  ```

  Do not proceed to Task 2 until the skill has been invoked and acknowledged.

---

### Task 2: Update the frontmatter `description:` field

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/SKILL.md` lines 1–3

- [ ] **Step 1: Read the current frontmatter**

  Read lines 1–4 of `.claude/skills/fbdi-compare-release/SKILL.md` and confirm the current description ends with `'run the test suite'."`.

- [ ] **Step 2: Replace the description value**

  Replace the entire `description:` line (line 2) with the following. The value is one long string — do not line-wrap inside the YAML quotes:

  ```yaml
  description: "Use when Oracle ships a quarterly FBDI release and the user wants the full download → clear → compare → catalog pipeline run end-to-end. Triggers on phrases like 'Oracle released 26C', 'compare 26A to 26B', 'run the quarterly FBDI update', 'update the FBDI Master Catalog for 26B', 'new FBDI release dropped', 'FBDI refresh for Q1'. Also triggers on report-only phrases like 'generate compliance report', 'generate the report for 26A 26B', 'regenerate the PDF', 'generate the HTML report' — for these, verify FBDI_Master_Catalog.xlsx and FBDI_to_ApplaudTables_Mapping.xlsx exist at repo root, then jump directly to Stage 9. Does NOT trigger on near-miss phrases like 'compare these two spreadsheets' or 'run the test suite'."
  ```

- [ ] **Step 3: Verify the frontmatter**

  Read lines 1–4 of the file again. Confirm:
  - Line 1 is `---`
  - Line 2 starts with `description:` and contains both `'FBDI refresh for Q1'` and `'generate compliance report'` and `jump directly to Stage 9`
  - Line 3 is `---`

---

### Task 3: Append the Stage 9 block

**Files:**
- Modify: `.claude/skills/fbdi-compare-release/SKILL.md` (append after last line)

- [ ] **Step 1: Read the current end of the file**

  Read the last 10 lines of `.claude/skills/fbdi-compare-release/SKILL.md`. Confirm the file ends with the Stage 8 section (the `verify_rerun.py` exit-code block and the general error-handling / resumability sections).

- [ ] **Step 2: Append the Stage 9 block**

  Append the following to the end of `.claude/skills/fbdi-compare-release/SKILL.md`, separated from the existing content by a blank line:

  ````markdown
  ## Stage 9 — Compliance Report

  After Stage 8 finishes — or, for a standalone report-only invocation (triggered
  by a report phrase rather than a full-pipeline phrase), immediately — present
  HITL #8.

  **Standalone preflight:** When Stage 9 is reached via a report-only invocation
  (Stages 1–8 were skipped), first confirm both required inputs exist at repo root:

  - `FBDI_Master_Catalog.xlsx`
  - `FBDI_to_ApplaudTables_Mapping.xlsx`

  If either is missing, stop:

  > "`<filename>` not found at repo root. Run the full pipeline (Stages 1–6)
  > first to generate it, or check your working directory."

  **HITL #8 — validation gate:**

  For a **full pipeline run** (Stages 1–8 just completed), present:

  > "Pipeline complete. Before generating the formal compliance report, please
  > validate the Excel outputs:
  >
  > - `Comparison_Report_<OLD>_<NEW>.xlsx` — do the total change count and top
  >   files from Stage 7 look plausible? Any files that look wrong?
  > - `FBDI_Master_Catalog.xlsx` — open the `<NEW>` sheet and spot-check a few
  >   rows. Reasonable row count? No obvious gaps or blank data columns?
  >
  > Ready to generate the HTML and PDF compliance report? (yes / skip)"

  For a **standalone report-only invocation** (Stages 1–8 skipped), replace the
  opener:

  > "Ready to generate the compliance report for `<OLD>` → `<NEW>`. Before
  > generating, please validate the Excel outputs:
  >
  > - `Comparison_Report_<OLD>_<NEW>.xlsx` — do the total change count and top
  >   files look plausible? Any files that look wrong?
  > - `FBDI_Master_Catalog.xlsx` — open the `<NEW>` sheet and spot-check a few
  >   rows. Reasonable row count? No obvious gaps or blank data columns?
  >
  > Ready to generate the HTML and PDF compliance report? (yes / skip)"

  If the user says **skip**, log `Compliance Report skipped at user request` and end.

  **Report generation:**

  ```
  python -m fbdi report --old <OLD> --new <NEW>
  ```

  Expected wall time: ~5–15 seconds.

  On success, print:

  ```
  Compliance Report generated:
    HTML: FBDI_Compliance_Report_<OLD>_<NEW>.html
    PDF:  FBDI_Compliance_Report_<OLD>_<NEW>.pdf
  ```

  **Error handling:**

  - **Mapping file missing** (CLI exits 1): Surface as — *"The mapping file
    `FBDI_to_ApplaudTables_Mapping.xlsx` wasn't found. This file is required for
    the compliance report. Is it in the repo root? If not, the report can't be
    generated until it's present."*

  - **PDF rendering fails (GTK/weasyprint traceback):** Parse the exception type
    and say — *"PDF generation failed — this usually means MSYS2/GTK isn't set
    up. See Known Hazards in CLAUDE.md for the install steps. The HTML file was
    likely written successfully; check `FBDI_Compliance_Report_<OLD>_<NEW>.html`
    first."*
  ````

- [ ] **Step 3: Verify the Stage 9 block**

  Read the last 60 lines of `.claude/skills/fbdi-compare-release/SKILL.md`. Confirm:
  - `## Stage 9 — Compliance Report` heading is present
  - `HITL #8` is named in the block
  - Both the "full pipeline run" and "standalone" HITL prompt variants are present
  - `python -m fbdi report --old <OLD> --new <NEW>` command is present
  - The two error-handling bullets (mapping missing, PDF rendering fails) are present
  - The file does not end with a trailing Stage 8 section after Stage 9 (order check: Stage 9 is last)

---

### Task 4: Verify HITL numbering consistency and commit

**Files:**
- Read: `.claude/skills/fbdi-compare-release/SKILL.md`

- [ ] **Step 1: Grep for all HITL references**

  Run:
  ```
  grep -n "HITL #" .claude/skills/fbdi-compare-release/SKILL.md
  ```

  Confirm the output contains `HITL #1` through `HITL #7` (existing) and exactly one `HITL #8` (new). No gaps, no duplicates.

- [ ] **Step 2: Check Stage ordering**

  Run:
  ```
  grep -n "^## Stage" .claude/skills/fbdi-compare-release/SKILL.md
  ```

  Confirm output is:
  ```
  ## Stage 1 — Environment preflight
  ## Stage 2 — Resolve OLD and NEW releases
  ## Stage 3 — Download + verify
  ## Stage 4 — Smart-clear
  ## Stage 5 — Compare
  ## Stage 6 — Catalog update
  ## Stage 6.5 — Populate Module column in mapping spreadsheet
  ## Stage 7 — Summary
  ## Stage 8 — Post-run verification
  ## Stage 9 — Compliance Report
  ```

- [ ] **Step 3: Commit**

  ```bash
  git add .claude/skills/fbdi-compare-release/SKILL.md
  git commit -m "feat(skill): add Stage 9 compliance report with HITL #8 validation gate"
  ```

---

## Self-review checklist (run before marking plan complete)

- [ ] Spec requirement — frontmatter trigger phrases added: covered by Task 2
- [ ] Spec requirement — standalone shortcut (jump to Stage 9 on report-only phrase): covered by Task 2 description value + Task 3 standalone preflight
- [ ] Spec requirement — HITL #8 validation gate with "Pipeline complete" opener: covered by Task 3
- [ ] Spec requirement — standalone HITL variant with different opener: covered by Task 3
- [ ] Spec requirement — skip option: covered by Task 3
- [ ] Spec requirement — `python -m fbdi report --old <OLD> --new <NEW>` command: covered by Task 3
- [ ] Spec requirement — standalone confirmation block (separate from Stage 7): covered by Task 3
- [ ] Spec requirement — error handling (mapping missing, PDF fails): covered by Task 3
- [ ] Spec requirement — no new helper scripts: verified — Tasks 2–4 touch only SKILL.md
- [ ] Spec requirement — `skill-creator:skill-creator` invoked first: covered by Task 1
