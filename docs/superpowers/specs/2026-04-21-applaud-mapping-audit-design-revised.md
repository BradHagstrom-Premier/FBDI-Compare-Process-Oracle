# FBDI ↔ Applaud Mapping Audit — Design (Revised)

**Date:** 2026-04-21
**Status:** Revised after Applaud technical audit — ready for implementation plan
**Trigger:** `docs/Applaud Mapping Audit.md` handoff — full re-audit of all 183 Applaud tables in `fbdi_applaud_mapping.xlsx` against the 26B FBDI catalog and the live Applaud MDB.
**Revision notes:** Corrected DataDictionary query patterns against live MDB schema; clarified prefix extraction; hardened snapshot field contract; added `@`-prefix field handling; minor adjudication fixes.

---

## 1. Goal

Produce `Claude_fbdi_applaud_mapping.xlsx` — a fully audited, evidence-backed mapping of every Applaud T_ table to its authoritative FBDI template/tab. Every current mapping decision (YES / UNMAPPED / multi-template) must be re-challenged against authoritative sources and either re-verified, revised, or escalated to Brad via a dedicated review queue.

The audit is keyed by Applaud table, not by FBDI file, because the handoff deliverable is one row per Applaud table.

---

## 2. Sources of truth

1. **`applaud-mcp`** — live access to `C:/Users/10193/Definian/MDB_for_ApplaudMCP/AP0STE.mdb`. Authoritative for Applaud table names, prefixes, key sequences, and DataDictionary fields.
2. **`FBDI_Master_Catalog.xlsx` — 26B tab** — authoritative for what FBDI files, tabs, and columns exist in the current release.
3. **`fbdi_applaud_mapping.xlsx`** — prior mapping representing Brad's accumulated human judgment. Read-only input; the audit re-challenges its assertions but preserves its annotations where verdicts are unchanged.

All three are consumed; none are written to. Output is written only to `Claude_fbdi_applaud_mapping.xlsx` + `Claude_fbdi_applaud_mapping_audit.md`.

---

## 3. Architecture

Two-step, Claude-to-Python handoff via a JSON snapshot:

```
┌────────────────────────┐        ┌────────────────────────┐        ┌─────────────────────────┐
│  applaud-mcp           │        │  FBDI_Master_Catalog   │        │  fbdi_applaud_mapping   │
│  (live MCP → .mdb)     │        │  .xlsx (26B tab)       │        │  .xlsx (current)        │
└───────────┬────────────┘        └───────────┬────────────┘        └────────────┬────────────┘
            │ Step A: Claude extracts         │                                  │
            │  (one-time, regen on MDB        │                                  │
            │   change)                       │                                  │
            ▼                                 │                                  │
┌────────────────────────┐                    │                                  │
│  applaud_snapshot.json │                    │                                  │
│  (checked into repo)   │                    │                                  │
└───────────┬────────────┘                    │                                  │
            │                                 │                                  │
            │         ┌───────────────────────┴──────────────────────────────────┘
            │         │
            ▼         ▼
      ┌──────────────────────────────────────────┐
      │  fbdi/audit.py (Step B)                  │
      │                                          │
      │  Pass 1: Build candidate index from      │
      │  catalog × snapshot (name, keys, cols)   │
      │                                          │
      │  Pass 2: For each Applaud table in the   │
      │  prior mapping, apply rubric →           │
      │  verdict + confidence + rationale        │
      │                                          │
      │  Emit outputs                            │
      └──────────────┬───────────────────────────┘
                     │
        ┌────────────┼──────────────────────┐
        ▼            ▼                      ▼
┌───────────────┐ ┌────────────────┐ ┌─────────────────────┐
│ Claude_...    │ │ Claude_...     │ │ stdout summary      │
│ mapping.xlsx  │ │ audit.md       │ │ counts by tier      │
│ (3 sheets)    │ │ (hard-case     │ │                     │
│               │ │  prose)        │ │                     │
└───────────────┘ └────────────────┘ └─────────────────────┘
```

**Key properties:**

- `applaud_snapshot.json` is the clean cut-line. Left of it: Claude + MCP. Right of it: deterministic Python with pytest.
- The audit is **idempotent**: same snapshot + same catalog = same outputs. Re-runnable on any dev's machine without MCP or an MDB.
- Prior `fbdi_applaud_mapping.xlsx` is **read-only input**. Output is always a fresh write to `Claude_fbdi_applaud_mapping.xlsx`.
- `fbdi/audit.py` is a new module alongside `build_mapping.py`. Not wired into `python -m fbdi` as a subcommand — standalone for the one-shot full audit; can be promoted to CLI later if it becomes routine.

---

## 4. Step A — Applaud snapshot extraction (Claude-driven)

### 4.1 Output: `applaud_snapshot.json`

```json
{
  "mdb_path": "C:/Users/10193/Definian/MDB_for_ApplaudMCP/AP0STE.mdb",
  "extracted_at": "2026-04-21T13:05:00Z",
  "extractor_version": "1",
  "tables": [
    {
      "name": "T_RA_INTERFACE_LINES_ALL",
      "prefix": "TA4",
      "description": "T_RA_INTERFACE_LINES_ALL (TA4)",
      "type": "1",
      "key_sequences": [
        { "seq": "1", "keys": ["TA4INTERFACE_LINE_ATTRIBUTE1"] },
        { "seq": "2", "keys": ["TA4INTERFACE_LINE_CONTEXT", "TA4INTERFACE_LINE_ATTRIBUTE1"] }
      ],
      "fields": [
        {
          "name": "TA4INTERFACE_LINE_ATTRIBUTE1",
          "bare_name": "INTERFACE_LINE_ATTRIBUTE1",
          "is_legacy_tracking": false,
          "data_type": "X",
          "length": 30
        }
      ]
    }
  ],
  "missing_tables": [
    { "name": "T_EXAMPLE_GHOST", "reason": "Not found in DatabaseTable" }
  ]
}
```

#### Field contract corrections

**`type` field:** The `Type` column in `DatabaseTable` is stored as a text value (e.g., `"1"`), not a numeric integer. Emit it as a string in the snapshot.

**`prefix` extraction:** The prefix is embedded in the `Description` field as a parenthetical suffix — e.g., `"T_RA_INTERFACE_LINES_ALL (TA4)"`. Parse using: strip the string, find the last `(`, extract everything between the last `(` and `)`. This is the only reliable extraction method — do not derive the prefix from the table name by regex, as it will fail for tables with non-obvious truncations.

**`fields` — DataDictionary query:** The DataDictionary table does **not** have a `DDID1` column. Query fields by prefix using the `Name` column:

```sql
SELECT Name, DataType, Size FROM DataDictionary WHERE Name LIKE '<prefix>%'
```

For example, to get fields for prefix `TA4`:
```sql
SELECT Name, DataType, Size FROM DataDictionary WHERE Name LIKE 'TA4%'
```

The `DataType` column (not `Type`) contains the Applaud type code (`X`, `N`, `D`). The `Size` column contains the field length.

**`is_legacy_tracking` flag (NEW):** DataDictionary fields whose `Name` starts with `@` (e.g., `@TA4LEGACY_HEADER1`, `@TA4SITE`) are legacy tracking fields populated by `CS^SET_@LEG_FIELDS`. These are infrastructure — not FBDI-mapped business fields. The snapshot must flag them, and Pass 1 must exclude them from `column_overlap` computation. Include the flag on each field object as shown above.

**`bare_name` derivation:** Strip the prefix (e.g., `TA4`) from the beginning of the `Name` value to get `bare_name`. For `@`-prefixed fields, strip both the `@` and the prefix — e.g., `@TA4SITE` → `bare_name = "SITE"`, `is_legacy_tracking = true`.

**`data_type` values:** The `DataType` column stores single-character Applaud type codes: `X` (character), `N` (numeric), `D` (date). Emit these directly into the snapshot.

### 4.2 Schema probe

Before the per-table loop, execute:

```sql
SELECT Name, DataType, Size FROM DataDictionary WHERE Name LIKE 'TA4%'
```

Expected: returns rows with `Name`, `DataType`, and `Size` columns. If this query returns 0 rows or the columns are absent, abort loudly with the actual columns returned. This validates both the column names and that the known prefix `TA4` has records.

### 4.3 Extraction procedure

1. **Working set = the 183 Applaud table names already in `fbdi_applaud_mapping.xlsx` Sheet2.** The audit scope is explicitly "every row in Sheet2"; newly-discovered tables (in the MDB but not in Sheet2) are out of scope.
2. **Schema probe** (see §4.2) — run before the loop.
3. **Per-table loop** (all 183):
   - `get_table_definition(name)` → description, prefix (parsed from parenthetical suffix in description), type (as string), key sequences.
   - `query_table("DataDictionary", where_clause="Name LIKE '<prefix>%'")` → all fields for that prefix. Strip prefix from `Name` to get `bare_name`. Flag `@`-prefixed fields as `is_legacy_tracking = true`.
4. **Tables not found** in `DatabaseTable` are written to the `missing_tables` array with reason. Pass 2 turns these into automatic `UNMAPPED` verdicts.
5. **Write** `applaud_snapshot.json` at repo root. Commit to Git so the audit is reproducible across machines.

### 4.4 Execution shape

Single Claude-driven agent run. ~183 tables × 2 MCP calls each. Estimated latency <5 min. No human-in-the-loop per table — it is mechanical data extraction.

---

## 5. Step B — Pass 1: candidate index

### 5.1 Purpose

Build a deterministic lookup `{applaud_table_name → [Candidate, ...]}` before any adjudication. Every candidate gets scored; nothing is decided here.

### 5.2 Per-candidate signals

For every `(fbdi_file, fbdi_tab)` in the 26B catalog, compute four signals against the Applaud table:

| Signal | Values |
|---|---|
| `name_alignment` | `EXACT` — Applaud name minus `T_` equals FBDI tab technical name. `PARTIAL` — matches after stripping `_ALL` / `_INT` / `_INTERFACE` suffix from either side. `NONE` — no name relation. |
| `key_coverage` | Fraction of the Applaud table's key fields (bare names) present in that FBDI tab's column set. `0.0` to `1.0`. |
| `column_overlap` | Fraction of the Applaud table's DataDictionary fields (bare names, **excluding `is_legacy_tracking = true` fields**) present in that FBDI tab's column set. `0.0` to `1.0`. |
| `prefix_conformance` | Boolean. True if the Applaud prefix follows the `T_` + FBDI-tab-name convention (exact match only — the set of "known truncations / collisions" is **not pre-enumerated**; mismatches are flagged to Brad via the markdown sidecar, and the audit builds up a confirmed exception list over iterations). Diagnostic only — does not affect verdict or confidence banding. |

**Key clarification on `column_overlap` denominator:** Use only non-legacy-tracking fields (those with `is_legacy_tracking = false`) in both numerator and denominator. Legacy tracking fields (`@`-prefixed) are infrastructure common to all T_ tables and are not present in any FBDI tab — including them would uniformly deflate overlap scores for no diagnostic value.

### 5.3 Candidate filter

Keep only candidates where **any** of:
- `name_alignment != NONE`
- `key_coverage ≥ 0.5`
- `column_overlap ≥ 0.3`

Anything below all three floors isn't worth showing even as a candidate.

### 5.4 Data structures

```python
@dataclass
class Candidate:
    fbdi_file: str
    fbdi_tab: str
    name_alignment: str           # EXACT | PARTIAL | NONE
    key_coverage: float           # 0.0..1.0
    column_overlap: float         # 0.0..1.0 (legacy tracking fields excluded)
    prefix_conformance: bool
    applaud_key_fields_matched: list[str]
    applaud_fields_matched: list[str]
    applaud_fields_missing: list[str]

candidate_index: dict[str, list[Candidate]]  # sorted by signal strength, strongest first
```

Full match/miss field lists are kept so pass 2 and the markdown sidecar can cite specific column names without recomputing.

### 5.5 Edge cases handled in pass 1

- Applaud table with no DataDictionary fields (0 non-tracking fields) → `column_overlap` is N/A; `key_coverage` alone drives candidacy. Do not divide by zero; treat overlap as 0.0 and note in rationale.
- Applaud table with only `@`-prefixed (legacy tracking) fields and no business fields → same as above.
- FBDI tabs with phantom column counts → respects the 500-column cap from the comparison engine.
- Case-insensitive matching throughout.

**Pass 1 does NOT** produce verdicts, collapse multi-mappings, or band confidence.

---

## 6. Step B — Pass 2: adjudication

### 6.1 Per-Applaud-table algorithm

```
Load prior state from Sheet2: prior_status, prior_prefix, prior_mapping_text
Load candidates from pass 1: candidates (sorted by signal strength)

1. PREFLIGHT — catch structural problems first
   - If table flagged NOT_IN_APPLAUD in snapshot:
       verdict = UNMAPPED, confidence = High
       rationale = "Applaud table not present in MDB snapshot"
       stop.
   - If prior_status in {FILE_TOO_LARGE, FILE_ERROR}:
       verdict = carry through, confidence = High
       rationale = "Sized out / unreadable in 26B — unchanged from prior"
       stop.

2. PRIOR MAPPING PARSE
   claimed = parse "Template / Tab; Template / Tab" → [(file, tab), ...]

3. SINGLE prior claim:
   Match claim against candidates.
   - High signals → verdict = claim, confidence = High
   - Medium signals → verdict = claim, confidence = Medium
   - Low / not in candidates → NEEDS_REVIEW

4. MULTI prior claims (31 rows currently):
   For each claim, check against candidates.
   - KEEP MULTI if: every claim independently scores High on the rubric
   - KEEP MULTI with NOTE if: every claim scores High-or-Medium, but at least one is Medium
     (note reminds Brad that one leg is PARTIAL-name-aligned)
   - COLLAPSE to highest-scoring if: only one claim scores High/Medium, others Low or absent
     from candidates
   - NEEDS_REVIEW if: two-or-more claims are Low, absent, or the mix is otherwise indecisive
   Rejected or demoted claims captured in evidence for the sidecar.

   Note: a shared technical tab name is a common pattern (e.g. T_EGP_COMPONENTS_INTERFACE
   mapped to ChangeOrderImportTemplate / EGP_COMPONENTS_INTERFACE and
   ItemStructureImportTemplate / EGP_COMPONENTS_INTERFACE) but is not a precondition —
   some legitimate multi-mappings span differently-named tabs. The rubric's column-overlap
   and key-coverage signals carry the weight, not tab-name identity.

5. UNMAPPED (37) or YES-with-blank (5):
   Re-challenge using candidates.
   - Any candidate scores High → verdict = that (file, tab), confidence = High (promoted)
   - Best candidate Medium → NEEDS_REVIEW (potential new mapping)
   - No candidate clears threshold → verdict = UNMAPPED, confidence = High
     rationale = "No FBDI tab in 26B catalog scores above threshold"

6. PREFIX AUDIT (all verdicts):
   Check prefix_conformance for the chosen (file, tab).
   False → append note: "Prefix mismatch — expected T_<tab>, got <actual>"
   Never changes verdict or confidence; surfaces in Notes.

7. EMIT ROW with verdict, confidence, one-line rationale, changed flag, evidence bundle.
```

### 6.2 Confidence tiers (per handoff rubric — Q6/A)

Evaluated in order; first match wins.

- **High** — `name_alignment == EXACT` AND (`key_coverage == 1.0` OR `column_overlap ≥ 0.7`)
- **Medium** — `name_alignment == PARTIAL`, OR (`0 < key_coverage < 1.0` AND `column_overlap ≥ 0.4`)
- **Low** — any other candidate retained by pass 1 (i.e. one signal cleared a pass-1 floor but the combination doesn't reach High or Medium)

**UNMAPPED** is assigned when pass 1 retained no candidate for that Applaud table at all (no signal cleared any floor).

§6.1 steps 3, 4, and 5 reference "High / Medium / Low" — all resolve against this rubric.

### 6.2.1 Which verdicts can carry which confidence

| Verdict | Allowed confidence | Notes |
|---|---|---|
| `YES` | High or Medium | Auto-verdict paths (§6.1 steps 3, 4 KEEP, 5 promotion) |
| `NEEDS_REVIEW` | High / Medium / Low | Confidence = best candidate's tier. Surfaces to Brad what quality of match the escalation is based on. |
| `UNMAPPED` | High only | Either the Applaud table isn't in the MDB, or pass 1 retained no candidate at all — either way, the "no match" conclusion itself is high-confidence. |
| `FILE_TOO_LARGE` / `FILE_ERROR` | blank | Carried through from prior without re-evaluation. |

This prevents the anti-pattern the handoff calls out: a forced low-confidence YES. Any Low-confidence match becomes NEEDS_REVIEW, not YES.

### 6.3 Data structure

```python
@dataclass
class AuditRow:
    applaud_table: str
    prefix: str
    verdict: str                   # YES | UNMAPPED | NEEDS_REVIEW | FILE_TOO_LARGE | FILE_ERROR
    fbdi_mapping: str              # rebuilt "Template / Tab[; Template / Tab]" or ""
    confidence: str                # H | M | L (blank for FILE_* rows)
    rationale: str                 # one-line inline rationale
    prior_verdict: str             # for change tracking
    changed: bool
    needs_deep_rationale: bool     # triggers markdown sidecar entry
    evidence: EvidenceBundle       # keys, columns matched/missed, rejected alts
```

### 6.4 Deep-rationale trigger

`needs_deep_rationale = True` when:
- `verdict == NEEDS_REVIEW`, OR
- verdict changed from prior (promoted, demoted, collapsed, etc.), OR
- `confidence == Low`, OR
- prefix_conformance failed

These rows get markdown-sidecar sections. High-confidence unchanged rows get only the one-line inline rationale on Sheet2 and nothing more.

### 6.5 Change tracking

Every row carries `prior_verdict` + `changed`. audit.md summary uses these to produce counts like "Of 183 rows: 14 changed, 6 promoted from UNMAPPED, 3 multi collapsed to single, 5 newly flagged Needs Review."

---

## 7. Outputs

### 7.1 `Claude_fbdi_applaud_mapping.xlsx` — 3 sheets

**Sheet 1: `FBDI Mapping`** — file × tab inventory, one row per `(file, tab)` in the 26B catalog.

| Col | Header |
|---|---|
| A | FBDI Template |
| B | FBDI Tab |
| C | Applaud Table |
| D | Prefix |
| E | Status (YES / UNMAPPED / NEEDS_REVIEW / FILE_TOO_LARGE / FILE_ERROR) |
| F | Module |
| G | Notes |
| H | Match Type (EXACT / PARTIAL / PRIOR-CARRYOVER) |
| I | Confidence (H / M / L) |

Module and Notes carried from prior mapping verbatim where verdict is unchanged — Brad's manual annotations are preserved.

**Sheet 2: `Applaud Tables`** — 183 rows, keyed by Applaud table. Primary deliverable.

| Col | Header |
|---|---|
| A | # |
| B | Applaud Table |
| C | Status (direct value — no XLOOKUP) |
| D | Prefix |
| E | FBDI Template Mappings (`Template / Tab` or multi) |
| F | Confidence |
| G | Rationale (one line) |
| H | Changed From Prior (✓ / blank) |
| I | Prior Status |

Row order preserved from original Sheet2. Status is stored as a direct value, removing the cross-sheet XLOOKUP coupling from the current workbook.

**Sheet 3: `Needs Review`** — filtered to rows where `needs_deep_rationale == True`. Same column schema as Sheet 2. Sorted: NEEDS_REVIEW first, then changed-from-prior, then Low confidence. Rationale column ends with "→ see audit.md". Expected size: 15-40 rows.

### 7.2 `Claude_fbdi_applaud_mapping_audit.md` — prose sidecar

One markdown section per Needs Review row + one per changed-from-prior + one per Low confidence + one per prefix mismatch. High-confidence unchanged rows get zero prose. Expected size: 30-80 sections, ~200-600 lines.

Structure:

```markdown
# FBDI ↔ Applaud Mapping Audit — 26B

**Generated:** <timestamp>
**Snapshot:** applaud_snapshot.json @ <extracted_at>
**Catalog:** FBDI_Master_Catalog.xlsx 26B tab
**Prior mapping:** fbdi_applaud_mapping.xlsx

## Summary

Of 183 Applaud tables audited: [counts by verdict + change type]

## Needs Review (N rows)

### T_XXX (prefix: TXX) — NEEDS_REVIEW
- **Prior:** YES → `TemplateA / TabA; TemplateB / TabB`
- **Decision:** <decision>
- **Candidates evaluated:**
  - `TemplateA / TabA` — name=EXACT, keys=5/5, cols=82% → High
  - `TemplateB / TabB` — name=EXACT, keys=3/5, cols=41% → Medium
- **Question for Brad:** <specific question>

## Prefix Mismatches

| Applaud Table | Prefix | Chosen FBDI Tab | Expected Prefix | Status |
```

---

## 8. Testing strategy

`tests/test_audit.py` — pytest, consistent with existing 139-test pattern.

### 8.1 Unit tests

1. **Signal computation** — EXACT / PARTIAL / NONE name alignment; key coverage 1.0 / 0.5 / empty; column overlap known ratios; case-insensitive.
2. **Legacy tracking field exclusion** — `column_overlap` denominator excludes `is_legacy_tracking = true` fields; table with only tracking fields treated as 0-field table.
3. **Pass 1 thresholding** — candidate kept/dropped by thresholds; sort order.
4. **Prior-mapping parser** — single, multi, blank, malformed input.
5. **Adjudication branches** — every path in §6.1 algorithm gets its own test:
   - NOT_IN_APPLAUD → UNMAPPED High
   - FILE_TOO_LARGE / FILE_ERROR carry-through
   - Single prior, High → YES High
   - Single prior, Low / missing → NEEDS_REVIEW
   - Multi, both High → multi retained
   - Multi, one High + one Low → collapsed
   - Multi, signals conflict → NEEDS_REVIEW
   - UNMAPPED + High candidate → promoted
   - UNMAPPED + Medium candidate → NEEDS_REVIEW
   - UNMAPPED + no viable candidate → stays UNMAPPED High
6. **Prefix conformance** — matching / mismatch with note.
7. **Prefix extraction from description** — parenthetical suffix parsing; handles unusual table descriptions cleanly.
8. **`bare_name` derivation** — regular field strips prefix; `@`-prefixed field strips `@` + prefix and sets `is_legacy_tracking = true`.
9. **`column_overlap` with zero non-tracking fields** — no division by zero; treated as 0.0; noted in rationale.
10. **Output writer** — three sheets, correct headers, NEEDS_REVIEW content on Sheet 3.

### 8.2 Integration test

`test_audit_end_to_end`:
- Synthetic snapshot with 5 Applaud tables covering EXACT / PARTIAL / UNMAPPED / multi-collapse / needs-review. Include at least one table with `@`-prefixed legacy tracking fields to confirm they are excluded from overlap.
- Synthetic catalog with matching FBDI files.
- Synthetic prior-mapping xlsx.
- Run `run_audit`, assert output xlsx + md exist and contain expected rows + rationales.

### 8.3 Not tested

- Snapshot extraction (Claude-driven, live MCP). Verified by inspection + downstream audit output.
- applaud-mcp server itself.
- Current `fbdi_applaud_mapping.xlsx` content (tested indirectly through pass 2).

### 8.4 Test-data reminder

Per `CLAUDE.md` test-data gotcha: `detect_header_row` scores rows by UPPER_SNAKE_CASE content. Synthetic sample values like `"CREATE"` false-positive as headers. Audit tests build catalog-shaped sheets with explicit header rows, so this doesn't apply directly — but if we ever add synthetic sample data rows, use lowercase/mixed-case.

**Target:** ~30-40 new tests (up from original 25-35 estimate due to added coverage for legacy tracking field handling and prefix extraction). Full suite goes from 139 → ~175. Runtime under 30s.

---

## 9. Error handling

### 9.1 Step A — snapshot extraction

| Failure | Handling |
|---|---|
| MCP unreachable / MDB path wrong | Abort loudly. No partial snapshot written. |
| DataDictionary schema probe fails | Abort. Error cites the probe query and actual columns returned. |
| Single table has no definition | Record `missing_tables` entry, continue. Pass 2 → UNMAPPED. |
| Single table's DataDictionary returns 0 rows | Record table with empty `fields`. Pass 2 falls back to key-coverage-only. |
| Single table's DataDictionary returns only `@`-prefixed rows | Record table with empty business fields (`fields` list contains tracking entries with `is_legacy_tracking = true` but no `false` entries). Pass 2 treats same as 0-field table for overlap purposes. |

### 9.2 Snapshot freshness

`extracted_at` recorded in file header. Pass 2 logs warning (not error) if snapshot is >30 days old. Brad decides whether to regenerate.

### 9.3 Step B — audit runtime

| Failure | Handling |
|---|---|
| `applaud_snapshot.json` missing | Hard error: "run Step A first" |
| `FBDI_Master_Catalog.xlsx` missing or no 26B tab | Hard error with explicit path |
| `fbdi_applaud_mapping.xlsx` missing | Hard error |
| Output file exists | Overwrite (matches `build_mapping.py` / `compare.py` behavior) |
| Prior mapping text malformed | Parser warning, best-effort parse, continue |
| Prior claims `(file, tab)` not in 26B catalog | Signals = NONE → NEEDS_REVIEW with rationale "Prior references file/tab not in 26B catalog" |
| Signal ties between candidates | Deterministic tiebreaker: alphabetical by `(file, tab)`. Tie noted in rationale. |
| Snapshot has fields but catalog tab has 0 columns | `column_overlap = 0`; name + key signals carry candidacy. |

### 9.4 Data boundaries

- Column-name matching: case-insensitive, whitespace-trimmed, otherwise literal. **No fuzzy matching** — the handoff's "don't force mappings" principle prefers miss+flag over similarity-based auto-bind.
- Prefix stripping for bare-name comparison: uses the exact prefix from the snapshot (e.g. `TA4`), not a general regex. Fields not starting with the known prefix keep their full name as `bare_name`.
- `@`-prefixed fields: strip `@` and prefix for `bare_name` derivation; mark `is_legacy_tracking = true`; exclude from all FBDI column-overlap comparisons.
- Multi-mapping parse: split on `;`, strip, split each half on ` / ` with exact spacing. Anything that doesn't cleanly split → malformed-parser warning path.

### 9.5 Out of scope

- Concurrent edits to input workbooks during audit (openpyxl reads on-disk state; no locking).
- Re-running against a different release (25D / 26A). Current design hardcodes 26B from catalog; later enhancement if needed.
- Partial re-audit of a subset of rows. Always full sweep.

---

## 10. Key decisions (locked)

| Decision | Choice | Rationale |
|---|---|---|
| Audit loop shape | Two-pass: candidate index first, then Sheet2 iteration | Multi-mapping rule is easiest to enforce against a pre-built index |
| Evidence format | Inline one-liners on Sheet2 + markdown sidecar for hard cases only | Most rows need ≤1 column of rationale; contested rows deserve prose |
| Multi-mapping policy | Auto-decide on overwhelming evidence, escalate contested to Needs Review | "Prove every multi is correct" — neither silent acceptance nor forced manual triage on all 31 |
| Output shape | Mirror existing two-sheet layout + `Needs Review` triage tab | Preserves Excel muscle memory; surfaces rows that need eyes |
| Confidence rubric | `name_alignment` × `key_coverage` × `column_overlap` with fixed thresholds (§6.2) | Deterministic, testable, matches handoff's signal language |
| Architecture | Step A (Claude/MCP → JSON snapshot) + Step B (deterministic Python audit) | Clean cut-line; audit is reproducible and testable |
| Module location | New `fbdi/audit.py` alongside `build_mapping.py` | Matches existing pattern; promote to `python -m fbdi` subcommand later if needed |
| DataDictionary query pattern | `WHERE Name LIKE '<prefix>%'` on the `Name` column | Confirmed against live MDB schema — `DataDictionary` has no `DDID1` column; `Name` column holds prefixed field names |
| Legacy tracking field exclusion | `@`-prefixed fields excluded from `column_overlap` denominator and numerator | These fields are conversion infrastructure present on all T_ tables; including them uniformly deflates overlap with no diagnostic value |

---

## 11. Implementation notes for Step A (Claude Code guidance)

The following are concrete implementation details verified against the live MDB schema.

### DataDictionary column names (verified)

The `DataDictionary` table has these relevant columns:

| Column | Type | Purpose |
|---|---|---|
| `Name` | text(60) | Fully prefixed field name (e.g. `TA4INTERFACE_LINE_ATTRIBUTE1`, `@TA4SITE`) |
| `DataType` | text(2) | Applaud type code: `X`, `N`, or `D` |
| `Size` | integer | Field length in characters (or total digits for numeric) |
| `DecPlaces` | byte | Decimal places (for N-type fields) |
| `Description` | text(60) | Human-readable field description |

**Do not use `DDID1`, `DDID2`, or `TableId` as query filters for field lookup.** The correct filter is `Name LIKE '<prefix>%'`.

### DatabaseTable column names (verified)

| Column | Type | Purpose |
|---|---|---|
| `Name` | text(60) | Table name (e.g. `T_RA_INTERFACE_LINES_ALL`) |
| `Description` | text(60) | Human-readable description including prefix in parens (e.g. `T_RA_INTERFACE_LINES_ALL (TA4)`) |
| `Type` | text(2) | Table type, stored as string (e.g. `"1"`) |

### `get_table_definition` return format

The `get_table_definition` MCP tool returns the table description plus key sequences in a structured format. The prefix is NOT returned as a separate field — it must be parsed from the `Description` string. Confirmed extraction pattern:

```python
import re
match = re.search(r'\(([A-Z0-9]+)\)\s*$', description)
prefix = match.group(1) if match else None
```

### MCP tool to use for field extraction

Use `applaud-mcp:query_table` (not `applaud-mcp:execute_query`) for field extraction, since the query is a simple filtered select with no joins:

```
query_table(
    table_name="DataDictionary",
    where_clause="Name LIKE 'TA4%'"
)
```

`execute_query` is available for complex cross-table SQL but is unnecessary here.

---

## 12. Next step

`superpowers:writing-plans` — produce a step-by-step implementation plan with review checkpoints. The plan should cover: (a) snapshot extraction agent task spec, (b) `fbdi/audit.py` build order (data classes → signal functions → pass 1 → pass 2 → output writers → CLI shim), (c) test file build order, (d) verification command to prove the audit is correct before handing off.
