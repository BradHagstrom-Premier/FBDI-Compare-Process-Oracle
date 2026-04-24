# Scraper Gap Findings — 2026-04-23

**Context:** Running the 26A→26B pipeline end-to-end (to ground the `fbdi-compare-release` skill design) surfaced a real bug in `tools/download_and_clear.py`. Capturing this so it can be fixed before the skill is implemented against the spec.

---

## The numbers

|                | Brad's Windows machine (authoritative) | Mac dev machine (two runs) | Gap |
|---|---|---|---|
| 26A originals  | 212 | 196, 197 | **-15 files** |
| 26B originals  | 212 | 138, 139 | **-73 files** |

Reference: `baseline_files.txt` (committed by Brad, generated from his work machine).

Two independent download runs on the Mac produced near-identical counts (±1), ruling out transient network flakiness. The gap is deterministic.

---

## The pattern

**All 15 missing 26A files are from the procurement module** (15/15 → 100%). Examples:
- `POBlanketPurchaseAgreementImportTemplate.xlsm`
- `POContractPurchaseAgreementImportTemplate.xlsm`
- `POPurchaseOrderImportTemplate.xlsm`
- `RequisitionImportTemplate.xlsm`
- `SupplierImportTemplate.xlsm` (+ 9 other Supplier-prefixed)
- `PoiDataSetImportTemplate.xlsm`
- `PONNegotiationLinesImportTemplate.xlsm`
- `SchExternalPurchasePricesImportTemplate.xlsm`

**The 73 missing 26B files are from procurement *and* financials.** Same 15 procurement files as 26A, plus 58 financials files:

| Prefix | Count |
|---|---|
| FixedAsset | 11 |
| Supplier | 9 |
| Lease / Revenue | 5 each |
| PO / Payables / CashManagement | 4 each |
| Receivables / Import / Cross / Upload / ChartOf | 2 each |
| (24 other prefixes) | 1 each |

Full lists in `/tmp/missing_26a.txt` and `/tmp/missing_26b.txt` at time of this investigation.

---

## Root cause (hypothesis)

Looking at `tools/download_and_clear.py:113–150`, `download_files()`:

1. Navigates to each of the 4 Oracle doc base URLs (`project-management`, `financials`, `procurement`, `supply-chain-and-manufacturing`).
2. Waits for `#navigationDrawer` to render (60s timeout with one refresh-and-retry).
3. Finds all `#navigationDrawer li` elements.
4. For each section, clicks any `.oj-clickable-icon-nocontext` expand icon and sleeps 1s.
5. Harvests `.xlsm` links under each section.

The logs for both runs show all 4 URLs completing without timeout — so step 2 worked. The gap is at step 4: **some section expansions silently fail**, and the downstream harvest only sees the links that are visible at that moment.

Evidence for this hypothesis:
- The gap is **module-concentrated** (only procurement for 26A; procurement + financials for 26B), not randomly distributed across modules. If it were a network issue, we'd expect random drops across all modules.
- The `supply-chain-and-manufacturing` module pulls down fully both runs — consistent with "some pages expand reliably, others don't."
- The procurement URL is the most-affected (missed on both runs, both releases). Plausible that its navigationDrawer has deeper nesting or slower-rendering sub-sections.

Why this hasn't bitten Brad on Windows: Chrome timing on Windows vs Mac differs. Oracle's JET framework's click-to-expand animation may be fast enough on Windows to complete within the 1s `time.sleep(1)`, and slow enough on Mac to regularly miss. A more robust fix would wait on a DOM signal rather than a fixed sleep.

---

## One inventory correction (minor)

`ProjectBudgetsImportTemplate.xlsm` appears in our 26B download. The original `baseline_files.txt` listed it as 26A-only (per Oracle's deprecation note), but Oracle's 26B docs page still serves the file, and the scraper correctly pulls it down. A sha256 check on 2026-04-23 confirmed the 26A and 26B copies are bit-identical.

Originally flagged as a "false positive" in this doc; that framing was wrong. The file is a legitimate 26B download — Oracle's served set is ground truth, not the deprecation note. **Fix: update `baseline_files.txt` to include `ProjectBudgetsImportTemplate.xlsm` in the 26B section** (done 2026-04-23; 26B count is now 213). The compare engine will show "no changes" for the file in 26A→26B, which is the correct representation ("Oracle didn't revise this deprecated template across the release").

This correction also reshapes the skill's extras handling (see "Update — 2026-04-23" below): extras against `baseline_files.txt` are almost always stale-inventory signal, not bad downloads — the skill should default to updating the inventory rather than quarantining.

---

## Implications for the skill design

The `fbdi-compare-release` skill spec (§6) already calls for a **file-count delta check vs the prior release** with a warn-and-retry threshold. That partially catches this, but:

- The threshold (>15% or >20 files) triggers on the 26B case (-73 files is 34%) but would miss the 26A case (-15 files is 7%).
- Retry on the same scraper doesn't help — the bug is deterministic on a given machine.

Better approach, to fold into the spec when it's revisited:

1. **Absolute expected count.** Maintain a `baseline_file_counts.json` (committed) that records the authoritative count per release from Brad's machine. Skill errors out if the downloaded count is short of that baseline by >N, with a clear message pointing at the scraper.
2. **Skill verification step** (new `scripts/verify_download.py` under Stage 3) — compares actual filenames against `baseline_files.txt` if present, and prints the module-grouped gap.
3. **Longer-term — fix the scraper:** replace `time.sleep(1)` after section expansion with an explicit wait on section-content appearance; add per-section verification that expected child count matches.

These are spec revisions, not new work. Capture them in the writing-plans step.

---

## Recommended next action (before skill implementation)

Fix `tools/download_and_clear.py` to reliably expand all navigationDrawer sections before harvesting. Verify by re-running 26A on Mac and confirming 212 files (or whatever the authoritative count is). Only then proceed with `writing-plans` for the skill — otherwise the skill's Stage 3 will keep producing incomplete data on Mac, which undermines the whole comparison.

---

## Artifacts

- `baseline_files.txt` — committed by Brad (authoritative inventory from Windows).
- `docs/superpowers/specs/2026-04-23-fbdi-compare-release-skill-design.md` — current design spec.
- `/tmp/missing_26a.txt`, `/tmp/missing_26b.txt`, `/tmp/extra_26b.txt` — diff outputs (ephemeral).

---

## Update — 2026-04-23 (Windows verification, post-fix)

After `b1150c4` (`fix(scraper): DOM-signal wait for section expand`) landed, re-verified on Brad's Windows machine with empty baselines folders:

|                   | 26A            | 26B (first run)                                       | 26B (retry)  |
|---                |---             |---                                                    |---           |
| Files downloaded  | 211            | 107                                                   | 211          |
| + manual drop     | 212            | 108                                                   | 212          |
| Diff vs baseline  | clean (0/0)    | **-105 missing (all SCM/Mfg), +1 extra (ProjectBudgets)** | 0 missing, +1 extra (ProjectBudgets) |

Same code, same machine, two consecutive runs, different outcomes. **This overturns the doc's earlier "deterministic per-machine-per-release" hypothesis.** The scraper gap is transient — the entire `supply-chain-and-manufacturing/26b` module silently returned zero files on run 1 (no timeout, no error, `Navigating to...` immediately followed by `Completed:` with no downloads in between), then harvested all 104 files normally on run 2.

### What this changes

- **The scraper fix from `b1150c4` is sufficient, not a blocker for the skill.** It reduces failure rate significantly — the prior Mac baseline was 138/212 on 26B, a 73-file deficit — but doesn't eliminate transient module-silent-failures. A verify-then-retry wrapper at the skill level closes the remaining gap.
- **Deterministic-per-machine was wrong.** The doc originally argued "retry on the same scraper doesn't help — the bug is deterministic on a given machine." Today's data shows a naive retry actually works on this class of failure. A retry is cheap and the verification step tells us whether it was needed.
- **ProjectBudgets is a legitimate 26B download, not a false positive.** On the first run of this verification Brad flagged the mischaracterization: Oracle serves this file from the 26B docs page; the scraper correctly pulls it down; it belongs in the 26B baseline. The original `baseline_files.txt` inventory was wrong (it listed the file as 26A-only). `baseline_files.txt` was updated 2026-04-23 to include `ProjectBudgetsImportTemplate.xlsm` in the 26B section (26B now 213 files); the "Only in 26A" difference line was removed. See the "One inventory correction" section above.

### Revised recommendation

Supersedes §"Recommended next action" above:

1. **Keep `b1150c4` as-is.** It's necessary but not sufficient on its own.
2. **In the `fbdi-compare-release` skill, add a post-download verification step** that diffs downloaded filenames against the per-release section of `baseline_files.txt` using `LC_ALL=C sort | comm -23`. If missing count > 0 (excluding known manual files like `RapidImplementationForCashManagement.xlsm`), retry the scraper once. If still short, group missing files by Oracle module URL and surface the gap to the user — most likely a genuine Oracle docs restructure that needs scraper-code attention.
3. **Extras handling:** anything downloaded that's not in `baseline_files.txt` is almost always stale-inventory signal (Oracle is trusted — what the scraper downloads is legitimate). The skill's default is to offer to update `baseline_files.txt` with the new filenames. Quarantine to `baselines/<ver>/_extras/` remains an option for the rare case the scraper pulled something genuinely unexpected. Details in the spec's §5 #6.
4. ~~Open question: how should the skill behave on the first run of a brand-new release?~~ **Resolved in the spec's §5 #6 "First-run sanity check" (2026-04-23 design-approval pass):** the skill proposes a bootstrap inventory from whatever the scraper downloaded, warns the user if the filename count deviates >15% from the prior release (guards against silent scraper failures on the very first run of a new release), and writes the confirmed inventory to `baseline_files.txt`.

These land in the `fbdi-compare-release` spec Stage 3 revision.
