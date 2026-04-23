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

## One false positive (minor)

`ProjectBudgetsImportTemplate.xlsm` appears in our 26B download but is **not** on Brad's 26B machine (per the `baseline_files.txt` header note, it's 26A-only and was deprecated in 26B). Oracle's 26B docs still renders a link resolving to the deprecated template file.

Not a blocker — the compare engine will show it as "present in both releases" instead of "removed in 26B", which is slightly misleading but not catastrophic. The fix is probably in the scraper: detect 404/redirect and drop.

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
