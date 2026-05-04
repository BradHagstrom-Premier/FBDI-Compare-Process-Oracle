# `python -m fbdi run` — Headless Chained Pipeline

**Status:** Design approved 2026-05-04. Implementation pending.

## Summary

A new top-level CLI subcommand, `python -m fbdi run`, that chains the existing
quarterly pipeline end-to-end with no Claude in the loop. It produces the same
outputs as the `fbdi-compare-release` skill — `Comparison_Report_<OLD>_<NEW>.xlsx`,
`FBDI_Master_Catalog.xlsx`, an updated `FBDI_to_ApplaudTables_Mapping.xlsx`, and
the HTML+PDF compliance report — driven by Python alone.

The skill remains in place. It is the right tool for guided runs where Claude's
judgment helps (first-time users, runs where Oracle's docs site has changed
shape, conversational checkpoints). `fbdi run` is the right tool for unattended
or developer-iteration runs.

## Implementer notes — read first

These notes are deliberately at the top because the user (Brad) plans to clear
context before execution. A fresh-context implementer needs to see them
immediately.

- **Skill update via `/skill-creator:skill-creator`.** This work moves five
  helper scripts out of `.claude/skills/fbdi-compare-release/scripts/` and into
  the `fbdi/` package. Every invocation of those scripts inside `SKILL.md`
  must be rewritten from `python .claude/skills/.../scripts/<name>.py` to the
  new `python -m fbdi <subcommand>` form. **When you reach that step, invoke
  `/skill-creator:skill-creator`** to drive the SKILL.md rewrite — it handles
  skill edits correctly and will catch reference sites a manual edit might
  miss. Do not edit `SKILL.md` by hand.

- **Historical plans are immutable.** Files under `docs/superpowers/plans/` —
  in particular `2026-04-23-fbdi-compare-release-skill.md` and
  `2026-04-30-quarterly-rerun-and-module-capture.md` — reference the old
  `.claude/skills/.../scripts/<name>.py` paths because they describe what was
  built at the time. **Do not "fix" those references.** They are historical
  records and stay as-is.

- **No CLI overrides on HITL policy.** The unattended-mode behaviors in
  Section 4 are hard-coded. Do not add flags like `--on-extras=stop` even if
  it feels like a small win. Future override needs are a separate design.

- **Subprocess per stage, not in-process.** Each stage runs as its own
  subprocess. This matches the project's existing idiom (`compare.py` already
  uses subprocess isolation per file pair) and keeps stage failures isolated.

## Goals and non-goals

### Goals

- **Primary:** scheduled/unattended runs (Windows Task Scheduler triggered the
  night Oracle ships a release). Must finish or bail cleanly with a useful log.
- **Secondary:** fast developer iteration via `--from`/`--to` stage ranges.
- **Tertiary:** manual no-Claude runs from the terminal.

### Non-goals (deliberately deferred)

- **Auto-discovery of new releases on Oracle docs.** The trigger-model
  "polling" option is reserved for a future iteration. The flag namespace
  `--auto-detect` is reserved so the addition is non-breaking, but it is not
  implemented in v1.
- **Notification/email/Slack hooks.** The manifest JSON and exit code are the
  contract; external scripts (cron-mailer, custom watchers) handle delivery.
- **Replacing the skill.** The skill keeps working; this just adds a parallel,
  headless way to drive the same pipeline.

## Architecture

A thin Python orchestrator at `fbdi/run.py` drives seven user-facing work
stages by subprocess, plus two auto-final stages (`summary`, `verify`) that
always run at the end regardless of `--from`/`--to`.

```
python -m fbdi run --old 26A --new 26B
        |
        v
   [orchestrator: fbdi/run.py]
        |
        +--> subprocess: python -m fbdi preflight                              # stage 1
        +--> subprocess: python tools/download_and_clear.py 26B --skip-clear   # stage 2 (download)
        +--> subprocess: python -m fbdi verify-download --release 26B          #   (verify + auto-accept extras)
        +--> subprocess: python tools/download_and_clear.py 26B --clear-only   # stage 3 (clear)
        +--> subprocess: python -m fbdi compare --old 26A --new 26B ...        # stage 4
        +--> subprocess: python -m fbdi catalog --release 26B                  # stage 5
        +--> subprocess: python -m fbdi populate-module --new 26B --old 26A    # stage 6 (update-module)
        +--> subprocess: python -m fbdi report --old 26A --new 26B             # stage 7
        +--> subprocess: python -m fbdi summarize ...                          # auto-final
        +--> subprocess: python -m fbdi verify-run --release 26B               # auto-final
        +--> subprocess: python -m fbdi verify-rerun ...                       # auto-final
```

Each subprocess captures stdout+stderr to
`logs/fbdi_run_<OLD>_<NEW>_<timestamp>.log`, parses exit code, and updates the
in-memory manifest. After every stage the orchestrator overwrites
`logs/fbdi_run_latest.json` so external watchers always have a stable path with
current state. On stage failure the orchestrator writes the final manifest and
exits with the appropriate code.

The skill, after the helper-promotion lands, calls the same promoted CLI
subcommands (`python -m fbdi preflight`, `python -m fbdi verify-download`, etc.)
instead of `.claude/skills/.../scripts/*.py` paths. Single source of truth.

## Stages

| # | Stage | Subprocess(es) | On non-zero exit |
|---|---|---|---|
| 1 | `preflight` | `python -m fbdi preflight` | Exit 2 (env failure) |
| 2 | `download` | `python tools/download_and_clear.py <ver> --skip-clear` then `python -m fbdi verify-download --release <ver>`. See HITL absorption below. | One auto-retry of both. Still failing → exit 3 |
| 3 | `clear` | `python tools/download_and_clear.py <NEW> --clear-only` (and `<OLD>` if freshly downloaded) | Capture timeouts to manifest; never blocks pipeline (clear writes to `blanks/`, compare reads `originals/`) |
| 4 | `compare` | `python -m fbdi compare --old <OLD> --new <NEW> --output Comparison_Report_<OLD>_<NEW>.xlsx` | Exit 4 (mid-pipeline failure) |
| 5 | `catalog` | Snapshot `FBDI_Master_Catalog.xlsx` to `.bak.xlsx` if present, then `python -m fbdi catalog --release <NEW>` | Exit 4 |
| 6 | `update-module` | If `FBDI_to_ApplaudTables_Mapping.xlsx` absent → status `skipped_no_mapping_file`. Else: backup (timestamped if name collides) → `python -m fbdi populate-module --new <NEW> --old <OLD>` | Exit 4 |
| 7 | `report` | `python -m fbdi report --old <OLD> --new <NEW>` | Exit 4 (HTML may have written; manifest records partial success) |
| — | `summary` (auto-final, always runs) | `python -m fbdi summarize --report ... --catalog ...` | Failures here don't change exit code |
| — | `verify` (auto-final, always runs) | `python -m fbdi verify-run --release <NEW>` then `python -m fbdi verify-rerun --release <NEW> --compare-report ... --baseline-catalog ...` | Regression flagged → bump exit code from 0 → 5 |

`summary` and `verify` always run at the end of every `fbdi run` invocation,
even when `--from`/`--to` excludes the upstream stages. The underlying scripts
already no-op gracefully on absent inputs, so a partial run still produces a
sensible manifest.

### HITL absorption inside the `download` stage

Two of the original skill HITLs are absorbed inside `download` to keep the
stage list flat:

- **HITL #1 — `OLD` baseline missing.** Before downloading `<NEW>`, check
  `baselines/<OLD>/originals/`. If empty or absent, the orchestrator runs the
  download for `<OLD>` first. If that download itself fails, exit 3. (Wall
  time when triggered: ~70 minutes total instead of ~50.)

- **HITL #2 — `RapidImplementationForCashManagement.xlsm` missing in `<NEW>`.**
  After the `<NEW>` download completes, the orchestrator checks for the FSM
  file. If missing, copy from `baselines/<OLD>/originals/`. If `<OLD>` does not
  have it either, the orchestrator continues with a prominent warning recorded
  to the manifest under `warnings`; final exit code bumps from 0 to 5.

The `verify-download` step inside `download` also absorbs **HITL #6 (extras)**:
when `verify-download` reports extras, the orchestrator re-invokes it with
`--commit-inventory` to update `baseline_files.txt`, then re-verifies. No
prompt.

## CLI surface

```
python -m fbdi run --old <RELEASE> --new <RELEASE>
                   [--from <stage>] [--to <stage>]
                   [--log-dir <path>]
```

| Flag | Default | Purpose |
|---|---|---|
| `--old` | required | Prior release (e.g., `26A`) |
| `--new` | required | New release (e.g., `26B`) |
| `--from` | `preflight` | First stage to execute |
| `--to` | `report` | Last stage to execute (inclusive) |
| `--log-dir` | `./logs` | Where to write `.log` and `.json` artifacts |

**Stage names for `--from`/`--to`:** `preflight`, `download`, `clear`,
`compare`, `catalog`, `update-module`, `report`. `summary` and `verify` are
not addressable via the range — they always run at the end.

**Validation:**

- Release format must match `^\d{2}[A-D]$` (e.g., `26A`, `27B`); otherwise
  exit 1.
- `--old` and `--new` must differ; otherwise exit 1.
- `--new` should be lexically greater than `--old`. Lexically smaller is a
  warning, not an error — backfill comparisons are sometimes legitimate.
- Invalid stage name in `--from`/`--to` → exit 1.
- `--to` earlier than `--from` → exit 1.

**Reserved for future:** `--auto-detect` (the polling trigger model). Not
implemented in v1.

### Usage examples

```bash
fbdi run --old 26A --new 26B                       # full pipeline
fbdi run --old 26A --new 26B --from compare        # skip download/clear (dev)
fbdi run --old 26A --new 26B --from report         # regenerate just the report
fbdi run --old 26A --new 26B --to compare          # stop after compare
fbdi run --old 26A --new 26B --from compare --to catalog   # only compare + catalog
```

## HITL → unattended policy table

The skill has eight HITL prompts. In unattended mode each one collapses to
either a hard-coded auto-decision or a hard-stop with a non-zero exit. No
flags expose these as overrides.

| Skill HITL | Trigger | `fbdi run` behavior |
|---|---|---|
| #1 | `OLD` baseline missing | Auto-download `<OLD>` first. If that download itself fails → exit 3. Adds ~15–20 min to wall time when triggered. |
| #2 | `RapidImplementationForCashManagement.xlsm` missing in `<NEW>` | Auto-copy from `<OLD>`. If `<OLD>` also missing the file → continue with prominent warning; manifest records `fsm_file: missing_from_both`; final exit code bumps to 5. |
| #3 | Version-mismatch sanity check | N/A — explicit `--old`/`--new` flags; nothing to mismatch. |
| #4 | >5 compare-pair failures | Log + continue. Failure count captured under `compare.pair_failures` in the manifest. |
| #5 | Download still short after one retry | Hard-stop, exit 3. Manifest records missing files grouped by module under `download.missing_by_module`. |
| #6 | Extras present in download | Auto-accept. Re-run `verify-download --commit-inventory` to update `baseline_files.txt`; re-verify. Manifest records `extras_accepted: [...]`. |
| #7 | Backup mapping spreadsheet before overwrite | Always backup to `FBDI_to_ApplaudTables_Mapping.bak.xlsx` (timestamped suffix if name collides). |
| #8 | Validation gate before report | Skip the gate, generate the report. Report is regeneratable; if scheduled, you want the deliverable. |

## Logging, manifest, and exit codes

### Three artifacts per run

1. `logs/fbdi_run_<OLD>_<NEW>_<timestamp>.log` — full captured stdout/stderr
   from every subprocess, appended in order.
2. `logs/fbdi_run_<OLD>_<NEW>_<timestamp>.json` — structured manifest (schema
   below).
3. `logs/fbdi_run_latest.json` — overwriting copy of (2) for stable-path
   watchers (notifications, dashboards).

### Manifest schema

```json
{
  "old": "26A",
  "new": "26B",
  "started":  "2026-07-15T02:00:00Z",
  "ended":    "2026-07-15T02:48:23Z",
  "duration_s": 2903,
  "exit_code": 0,
  "from_stage": "preflight",
  "to_stage":   "report",
  "stages": {
    "preflight":     {"status": "ok",  "duration_s": 1.2},
    "download":      {"status": "ok",  "duration_s": 1182, "files": 195, "extras_accepted": ["X.xlsm"], "fsm_file": "copied_from_old"},
    "clear":         {"status": "ok",  "duration_s": 145,  "timeouts": ["PayablesCollectionDocuments.xlsm"]},
    "compare":       {"status": "ok",  "duration_s": 287,  "total_changes": 1247, "files_with_changes": 89, "pair_failures": 2},
    "catalog":       {"status": "ok",  "duration_s": 198},
    "update-module": {"status": "ok",  "duration_s": 4,    "populated": 412, "blank": 3, "overwritten": 18},
    "report":        {"status": "ok",  "duration_s": 11,   "html": "FBDI_Compliance_Report_26A_26B.html", "pdf": "FBDI_Compliance_Report_26A_26B.pdf"}
  },
  "summary": {"status": "ok"},
  "verify":  {"status": "ok", "regressions": []},
  "warnings": []
}
```

**Status values per stage:**

- `ok` — completed cleanly.
- `skipped_by_range` — outside the requested `--from`/`--to` range.
- `skipped_no_mapping_file` — `update-module` only; mapping spreadsheet absent.
- `failed` — subprocess returned non-zero. The manifest's top-level
  `exit_code` reflects which failure triggered the bail.
- `partial` — subprocess wrote some output then crashed. Used by `report`
  when the HTML wrote successfully but the PDF render failed.

**Warnings as a top-level array.** Any condition that doesn't fail the run but
should bump the exit code from 0 to 5 (FSM file missing from both, regression
flagged by verify, etc.) appends a string to `warnings`.

### Exit codes

| Code | Meaning |
|---|---|
| 0 | Clean — all stages OK, no warnings |
| 1 | Configuration error (bad flags, invalid stage name, malformed release) |
| 2 | Environment preflight failed |
| 3 | Download failed after retry, or `OLD` baseline auto-download failed |
| 4 | Mid-pipeline crash (compare, catalog, update-module, or report) |
| 5 | Completed-with-warnings (regression flagged, FSM file missing from both, partial report, etc.) |

The manifest is the source of truth. The exit code is a coarse signal for
schedulers and notification scripts.

## Implementation: file layout

### New files

- `fbdi/run.py` — the orchestrator. Public function
  `run_pipeline(old, new, from_stage, to_stage, log_dir) -> int`. CLI
  registration `python -m fbdi run`.
- `fbdi/manifest.py` — manifest object: in-memory state, schema validation,
  atomic write (`.tmp` then rename) of `latest.json` after each stage and the
  timestamped manifest at the end.
- `tests/test_run.py` — orchestrator tests (stage range parsing, exit-code
  propagation, manifest writing, HITL policy behavior using subprocess fakes).

### Promoted helpers (move from skill into `fbdi/`)

| Old path | New path | New CLI subcommand |
|---|---|---|
| `.claude/skills/fbdi-compare-release/scripts/check_env.py` | `fbdi/preflight.py` | `python -m fbdi preflight` |
| `.claude/skills/fbdi-compare-release/scripts/verify_download.py` | `fbdi/verify_download.py` | `python -m fbdi verify-download` |
| `.claude/skills/fbdi-compare-release/scripts/verify_run.py` | `fbdi/verify_run.py` | `python -m fbdi verify-run` |
| `.claude/skills/fbdi-compare-release/scripts/verify_rerun.py` | `fbdi/verify_rerun.py` | `python -m fbdi verify-rerun` |
| `.claude/skills/fbdi-compare-release/scripts/summarize_report.py` | `fbdi/summary.py` | `python -m fbdi summarize` |

The promotion is mechanical: keep file behavior identical, change import
paths, register a CLI subcommand in `fbdi/cli.py`. The empty
`scripts/__init__.py` and the now-empty `scripts/` directory are removed.

### Test files that need updating

- `tests/test_skill_scripts.py` — uses `from scripts import check_env /
  verify_download / summarize_report / verify_run` with a `sys.path` injection
  at the top. After the move, these become `from fbdi.preflight import …`,
  `from fbdi.verify_download import …`, etc. The file can be split into
  per-module test files (`tests/test_preflight.py`,
  `tests/test_verify_download.py`, etc.) or stay as one — implementer's
  choice. The `sys.path` injection block at the top of the file goes away.

- `tests/test_verify_rerun.py` — uses
  `importlib.util.spec_from_file_location("verify_rerun", SKILL_SCRIPT)` to
  load by file path. Becomes a normal `from fbdi.verify_rerun import …`.

### Skill update (drive via `/skill-creator:skill-creator`)

Six invocation lines and five reference mentions inside
`.claude/skills/fbdi-compare-release/SKILL.md` need rewriting from
`python .claude/skills/.../scripts/<name>.py` to
`python -m fbdi <subcommand>`. Drive this through
`/skill-creator:skill-creator` rather than hand-editing.

The skill's external behavior must not change: same stages, same HITL
prompts, same outputs. Only the invocation paths inside it change.

## Testing strategy

- **Unit tests** for the orchestrator (`tests/test_run.py`):
  - Stage range parsing: valid ranges, invalid stage names, `--to` before
    `--from`.
  - Manifest writing: shape matches schema; `latest.json` overwritten after
    each stage; atomic write semantics.
  - Exit-code mapping: each failure mode maps to the correct code.
  - HITL policies: simulated subprocess outputs trigger the right
    auto-decisions (extras-accept, FSM-copy, OLD auto-download, etc.) — uses
    fakes/mocks since real downloads are too slow for unit tests.

- **Integration test** (slow, opt-in via `pytest -m integration`): full
  `fbdi run --old <test-fixture> --new <test-fixture> --from compare` against
  tiny fixture xlsm files in `tests/fixtures/`. Validates manifest
  end-to-end without real Selenium downloads.

- **Existing test suite (320 tests)** must continue to pass with zero
  regressions. The promoted helpers keep their behavior contracts — only
  their import paths change.

## Resumability

Each stage is idempotent on output-existence terms (same as the skill today):

- `preflight`: cheap, always re-runs.
- `download`: re-running wipes `originals/` first (destructive); the
  orchestrator does not auto-rerun a successful download in the same
  invocation.
- `clear`, `compare`, `catalog`, `update-module`, `report`: re-running is
  safe; they overwrite their outputs.

If a run crashes at, say, `compare`, the user can re-invoke with
`--from compare` once they've fixed the underlying issue. The downloads and
clears from the prior run are preserved.

## Wall-time expectations

- Full pipeline, both baselines present: 35–50 minutes (downloads dominate).
- Full pipeline, `OLD` baseline missing (HITL #1 triggers):
  ~70 minutes total.
- `--from compare` (downloads/clears skipped, e.g., dev iteration): ~10
  minutes.
- `--from report`: ~5–15 seconds.

These match the existing skill's wall-time profile.
