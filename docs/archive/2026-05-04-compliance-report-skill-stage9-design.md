# FBDI Compliance Report — Skill Stage 9 Design Spec

**Date:** 2026-05-04
**Author:** Brad Hagstrom (with Claude Code)
**Status:** Approved for implementation planning

## Purpose

Add `## Stage 9 — Compliance Report` to the existing `fbdi-compare-release` skill so that the HTML and PDF compliance report (`FBDI_Compliance_Report_<OLD>_<NEW>.html` / `.pdf`) is generated as the final step of the quarterly pipeline, behind a human validation gate.

The report generator (`fbdi/report.py`, CLI: `python -m fbdi report`) is already fully built. This spec covers only the skill orchestration layer — no changes to Python code.

---

## Skill to invoke during implementation

> **IMPORTANT:** When implementing this spec, invoke `skill-creator:skill-creator` before making any edits to `SKILL.md`. This note is here explicitly because the session may be cleared between planning and execution.

---

## Decisions (closed — do not re-litigate)

| Decision | Choice | Rationale |
|---|---|---|
| New skill vs. extend existing | Extend `fbdi-compare-release` | Report is the natural end of the quarterly pipeline; one skill keeps the flow unified. |
| Stage position | After Stage 8 (current end) | By Stage 8 the user has the full summary + verification output in front of them, which gives context for validating the Excel files. |
| Output presentation | Standalone confirmation block, separate from Stage 7 summary | Stage 7 covers the data pipeline; the report confirmation is the delivery moment and deserves its own beat. |
| Validation gate | HITL #8 — mandatory pause before generation | Brad confirmed: the formal deliverable should not generate until a human has validated the catalog and comparison Excel. |
| Skip option | Yes — user can say "skip" | Pipeline should not be held hostage if the user deliberately doesn't need the report this run. |
| Helper scripts | None | `weasyprint` fails loudly; the CLI exits 1 with a clear message on missing inputs. No silent failure mode warrants a verify script. |

---

## Changes to `SKILL.md`

Two edits only. No new files.

### 1. Frontmatter `description:` update

Append report-specific trigger phrases so the skill also fires on standalone report reruns:

```
Also triggers on: "generate compliance report", "generate the report for 26A 26B",
"regenerate the PDF", "generate the HTML report".
```

When triggered this way (no download/compare intent), the skill checks that both
`FBDI_Master_Catalog.xlsx` and `FBDI_to_ApplaudTables_Mapping.xlsx` exist at the
repo root, then jumps directly to Stage 9's HITL — skipping Stages 1–8.

### 2. New `## Stage 9 — Compliance Report` block

Appended after Stage 8. Full text:

---

#### Stage 9 — Compliance Report

**Beat 1 — HITL #8 (validation gate):**

After Stage 8 finishes (or immediately, for a standalone report-only invocation),
pause and present:

> "Pipeline complete. Before generating the formal compliance report, please
> validate the Excel outputs:
>
> - `Comparison_Report_<OLD>_<NEW>.xlsx` — do the total change count and top
>   files from Stage 7 look plausible? Any files that look wrong?
> - `FBDI_Master_Catalog.xlsx` — open the `<NEW>` sheet and spot-check a few
>   rows. Reasonable row count? No obvious gaps or blank data columns?
>
> Ready to generate the HTML and PDF compliance report? (yes / skip)"

For standalone report-only invocations, replace the "Pipeline complete." opener
with: *"Ready to generate the compliance report for `<OLD>` → `<NEW>`."* — then
present the same two validation bullets and prompt.

If the user says **skip**, log `Compliance Report skipped at user request` and end.

**Beat 2 — Report generation:**

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

---

## HITL numbering

Existing stable IDs `#1`–`#7` are unchanged. The new gate is `HITL #8`.

---

## Resumability

Stage 9 is idempotent — rerunning `python -m fbdi report` overwrites the HTML
and PDF with fresh output. If the user skipped Stage 9 on a prior run and wants
to generate the report later, they can invoke the skill with a report-only phrase
(see frontmatter trigger update) and it will jump directly to the HITL #8 gate.

---

## Out of scope

- No changes to `fbdi/report.py`, `fbdi/cli.py`, or any Python module.
- No `verify_report.py` helper script.
- No changes to `summarize_report.py` or the Stage 7 summary format.
- No changes to Stages 1–8.

---

## Implementation skill

> **Invoke `skill-creator:skill-creator` at the start of the implementation
> session before touching any files.** This is the correct skill for editing
> an existing skill's `SKILL.md`. Do not use `superpowers:writing-skills`
> or `superpowers:executing-plans` as a substitute.
