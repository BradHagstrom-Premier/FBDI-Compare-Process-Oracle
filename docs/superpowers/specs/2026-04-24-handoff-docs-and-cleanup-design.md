# Handoff Docs and Repo Cleanup — Design

**Date:** 2026-04-24
**Author:** Brad Hagstrom (with Claude)
**Status:** Design approved, ready for writing-plans.

---

## Goal

Turn this repo into something a Definian consultant or developer can fork and work from cleanly on day one. Two specific deliverables:

1. An operator guide that documents what the `/fbdi-compare-release` skill does, stage by stage, so a new coworker can run the quarterly refresh without sitting next to Brad.
2. A developer guide that orients a new dev to the codebase — modules, testing conventions, how to extend it.

Plus a repo-cleanup pass that removes cruft, organizes narrative history into `docs/archive/`, and refreshes `CLAUDE.md` to match the post-cleanup state.

---

## Audience

Both docs target Definian team members — consultants or devs who may fork this repo. Coverage is dual-purpose (operator + developer) because any team member might end up in either role.

---

## Design decisions (from brainstorming)

| Axis | Choice |
|---|---|
| Audience | Both operator and developer (one guide each) |
| Archive strategy | Gitignored `Archive/` for dead code; tracked `docs/archive/` for narrative docs |
| Doc shape | Two separate files, linked from README |
| Depth | Medium (~2.5K words per doc) — self-serviceable on first run |
| Humanizer scope | Full pass on the two new docs; light pass on README; skip `CLAUDE.md` and `SKILL.md` |
| `CLAUDE.md` | Gets an improver pass via `claude-md-management:claude-md-improver` after cleanup |
| Workflow | Direct push to master (no PR), one commit per logical unit |

---

## Cleanup plan

### Bucket 1 — move to gitignored `Archive/` (preserved locally, removed from tracking)

| Item | Reason |
|---|---|
| `complete_mapping.py` | March 27 one-off; graph confirms no references; superseded by `fbdi/build_mapping.py` |
| `format_workbooks.py` | March 27 one-off; Community 14 isolated island; cosmetic-only |

### Bucket 2 — delete outright (no archive needed)

| Item | Reason |
|---|---|
| `~$FBDI_Master_Catalog.xlsx` | Excel lock file, regenerated on open |
| `_extract_cache/` | Empty dir, unused |
| `Comparison_Report.xlsx` (root, Mar 24) | Gitignored already; predates naming convention |
| `Diagnostic_Report_26B.xlsx` (root) | Gitignored; regenerable via CLI |

### Bucket 3 — move to tracked `docs/archive/` (narrative history preserved in fork)

| Source | Destination |
|---|---|
| `docs/Applaud Mapping Audit.md` | `docs/archive/applaud-mapping-audit-notes.md` |
| `docs/scraper-gap-findings-2026-04-23.md` | `docs/archive/scraper-gap-findings-2026-04-23.md` |
| `Claude_fbdi_applaud_mapping_audit.md` (root) | `docs/archive/claude-fbdi-applaud-mapping-audit.md` |

### Bucket 4 — relocate (keep tracked or ignored in new location)

| Item | Action |
|---|---|
| `applaud_snapshot.json` (3.3 MB, root) | Move to `baselines/applaud/applaud_snapshot.json`. `baselines/` is gitignored, so the file becomes untracked but lives in its logical home. Any code that reads it (`fbdi/audit.py`) must be updated to read from the new path. |

### Stays put (no change)

- `README.md`, `CLAUDE.md`, `requirements.txt`, `.python-version`, `.gitignore`
- `baseline_files.txt` (skill reads it in Stage 3)
- `Comparison_Report_26A_26B.xlsx`, `FBDI_Master_Catalog.xlsx` (live deliverables, tracked)
- `fbdi_applaud_mapping.xlsx` **and** `Claude_fbdi_applaud_mapping.xlsx` — Brad is actively comparing them; both remain at root
- `fbdi/`, `tools/`, `tests/`, `.claude/`, `docs/superpowers/`, `reference/`

---

## Doc structure

### New tracked files

```
docs/
├── operator-guide.md
├── developer-guide.md
└── archive/
    ├── applaud-mapping-audit-notes.md
    ├── scraper-gap-findings-2026-04-23.md
    └── claude-fbdi-applaud-mapping-audit.md
```

### README role

Lean landing page. Keeps current sections (Setup, Known hazards, Repo structure, Testing, Status). Adds two pointers near the top:

> **Running it:** see [`docs/operator-guide.md`](docs/operator-guide.md).
> **Developing on it:** see [`docs/developer-guide.md`](docs/developer-guide.md).

The `Repo structure` tree in README is updated to reflect the new `docs/archive/` dir and removed root files.

### Cross-links between guides

- Operator guide ends with a pointer to developer guide ("what's happening under the hood").
- Developer guide opens with a pointer to operator guide ("if you just want to run it").
- Both link to `SKILL.md` for the Claude-facing HITL numbering and to `CLAUDE.md` for project-facing context.

### Unchanged docs

- `SKILL.md` — Claude-facing, stays as-is.
- `CLAUDE.md` — Claude context, improved via the improver skill at the end.
- `docs/superpowers/specs/` and `docs/superpowers/plans/` — historical design artifacts.
- `reference/` — pre-Python VBA archive.

---

## `operator-guide.md` content outline

Target ~2.5K words. Tone: second person, concrete, self-serviceable.

1. **What this is and who it's for** (~100 words) — one paragraph setting the context.
2. **Before your first run** (~200 words) — env checklist, Windows sleep warning, time budget, the two invocation paths (Claude Code vs. CLI).
3. **The 8 stages, in order** (~1.5K words) — per stage: what it does (plain English), what you see on screen, expected wall time, what to do if it stalls.
4. **The 6 HITL checkpoints** (~400 words) — using the same `HITL #1`–`#6` IDs as `SKILL.md`. What triggers each, what the options mean, how to decide.
5. **Reading the outputs** (~200 words) — `Comparison_Report` column meaning, `FBDI_Master_Catalog` per-release / Issues / Drift tabs.
6. **When something goes sideways** (~150 words) — pointer to `SKILL.md` "Error handling" section; the non-auto-downloadable FSM file; Ctrl-C resume behavior.
7. **Next steps** (~50 words) — pointer to developer guide.

### Non-goals for this guide

- No reimplementation of HITL prompt text verbatim from `SKILL.md` — paraphrased to the operator's perspective.
- No terminal output screenshots except where stdout patterns matter (e.g., the `TIMED OUT` block from Stage 4).
- No Oracle Fusion UI screenshots — textual path is sufficient.

---

## `developer-guide.md` content outline

Target ~2.5K words. Tone: second person, specific.

1. **Orientation** (~150 words) — problem being solved, what's shipped, what's on the frontier.
2. **Local setup for development** (~200 words) — clone, Python 3.14 with pyenv-win, deps, baselines (gitignored — either run `tools/download_and_clear.py` or copy from teammate), pre-touch test run.
3. **Codebase tour** (~600 words) — module-by-module walk of `fbdi/` and `tools/`. Written for a reader, not a reference. Covers responsibility boundaries and how modules connect.
4. **The `/fbdi-compare-release` skill** (~250 words) — purpose (glue, not logic), location, the four bundled scripts and their exit codes, when to modify skill vs. CLI.
5. **Testing conventions** (~250 words) — pytest layout, the VBA spot-check script, synthetic-workbook pattern (no fixtures), the UPPER_SNAKE_CASE false-positive gotcha.
6. **How to add a new release handler** (~300 words) — concrete walkthrough anchored on "Oracle ships 27A."
7. **Design docs and how we work** (~200 words) — specs vs. plans, the resolved-hazards log in `CLAUDE.md`, CodeRabbit for PR review.
8. **Known hazards and gotchas** (~300 words) — phantom columns, corrupt xlsm, 5 MB cap in diagnose/build_mapping, VBA-output corrupt stylesheet. Reformatted from `CLAUDE.md` for a developer perspective.
9. **Where to ask for help** (~50 words) — pointers to specs, plans, `CLAUDE.md`, `SKILL.md`. No "DM Brad" as the only answer.

---

## Sequencing

Ordering matters — cleanup happens before docs so the docs describe the final state.

1. **Cleanup.** Execute Buckets 1, 2, 3, 4. Update `fbdi/audit.py` if it references `applaud_snapshot.json` by its old path.
2. **Write `operator-guide.md`.** Full draft per the operator-guide outline above, in normal voice.
3. **Write `developer-guide.md`.** Full draft per the developer-guide outline above, in normal voice.
4. **Light humanizer pass on README.** Load `humanizer-skill:humanizer`. Target README only. Remove AI-tells without restructuring.
5. **Full humanizer pass on both new docs.** Same skill, fuller treatment. Run independently so voices stay consistent with the humanized README.
6. **Update README pointers.** Add the two guide links near the top. Update the `Repo structure` tree.
7. **CLAUDE.md improver pass.** Run `/claude-md-management:claude-md-improver`. Final repo state is visible to the improver, so it can prune references to deleted files, add pointers to the new guides, and tighten the `reference/` vs. `docs/archive/` distinction.
8. **Test suite sanity check (verify only).** `python -m pytest tests/` — confirms the `applaud_snapshot.json` path change (and any other cleanup-related edit) didn't break an import. 241 tests expected to pass. No commit from this step.
9. **Graphify rebuild (verify only).** `python -c "from graphify.watch import _rebuild_code; from pathlib import Path; _rebuild_code(Path('.'))"` — `graphify-out/` is gitignored so this produces no commit; it keeps the local knowledge graph current for future sessions.
10. **Commit and push.** Direct push to master per Brad's workflow. One commit per logical unit (expected exactly 4 commits):
    - `chore(cleanup): archive orphan scripts, remove stale artifacts, relocate applaud_snapshot`
    - `docs: add operator and developer guides, archive narrative history`
    - `docs(readme): add guide pointers, update repo structure, light humanizer pass`
    - `docs(claude-md): refresh after handoff-docs cleanup`

---

## Scope boundaries

Explicitly **not** in scope:

- No changes to `fbdi/` package logic (only the `applaud_snapshot.json` path reference in `audit.py` if it exists).
- No changes to `SKILL.md`, the skill's `scripts/`, or its `references/`.
- No changes to `docs/superpowers/specs/` or `plans/` (historical).
- No changes to `reference/` (pre-Python VBA archive).
- No new CLI commands, no new tests, no Python refactors.

---

## Success criteria

- A fork of this repo can be handed to a Definian coworker and they can:
  - Run `/fbdi-compare-release` end-to-end using only `docs/operator-guide.md`.
  - Navigate the codebase and make a first change using only `docs/developer-guide.md`.
- The two new docs read as human-written (humanizer pass applied).
- README is the launchpad — no reader has to dig to find either guide.
- `CLAUDE.md` accurately reflects post-cleanup repo state.
- Root directory is clean: no orphan scripts, no Excel lock files, no stale reports.
- Test suite still passes. Graph is current.

---

## Open items for implementation

None — all design decisions resolved during brainstorming. Ready for writing-plans to produce the detailed implementation plan.
