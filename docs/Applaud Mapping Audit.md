# Handoff: FBDI ↔ Applaud Mapping — Full Audit & Completion

## Context

`fbdi_applaud_mapping.xlsx` is partially populated. Brad needs an audit of all 183 Applaud tables, with every mapping validated against two authoritative sources:

1. **`applaud-mcp`** — live access to Applaud table definitions (database tables, columns, IFs, EFs, code sections)
2. **`FBDI_Master_Catalog.xlsx`** — authoritative snapshot of all Oracle FBDI files, tabs, and columns for the current release

The prior audit treated multi-template mappings as valid Oracle design. **This audit challenges that assumption.** Every multi-mapping must be challenged. You must prove that a multi mapping is correct.

The output is a new `Claude_fbdi_applaud_mapping.xlsx`

---

## Scope

**Every row in Sheet2 is in scope.** All 183 Applaud tables get audited, including:

- `YES`-status rows with blank mappings (must be resolved, it is possible that Applaud contains tables that are no longer in use by Oracle)
- `UNMAPPED` rows (re-challenge — verify no FBDI counterpart exists, it is possible that Applaud contains tables that are no longer in use by Oracle)
- Existing multi-template mappings (collapse to single best; flag rejected templates)
- "Clean" single-template mappings (field-level validation — prefix correctness, column alignment, IF/EF compatibility)

---

## Files in Scope

**NEW:**
- `Claude_fbdi_applaud_mapping.xlsx` (primary output)

**Read-only sources:**
- `FBDI_Master_Catalog.xlsx` — Oracle files/tabs/columns ground truth
- `applaud-mcp` tools — Applaud table definitions, IFs, EFs, code sections

---

## Hard Rules

1. Do not silently allow multi-mappings. Prove and verify with Brad a multi mapping is correct.

2. **Prefix correctness is part of the audit.** The Applaud naming convention is typically `T_` + FBDI tab name (with known collisions, truncatations, etc.). Verify every proposed mapping respects this pattern. Flag mismatches.

3. **Evidence is required for every proposed change.** Every modification must cite:
   - Applaud column names / IF-EF structure from `applaud-mcp`
   - FBDI tab + column list from `FBDI_Master_Catalog.xlsx`
   - A confidence score: **High / Medium / Low**

4. **Don't force mappings.** If no FBDI tab plausibly matches an Applaud table, the correct status is `UNMAPPED`. A forced low-confidence mapping is worse than no mapping.

---

## Workflow

### Step 0 — Load Context

Read in order:
- `CLAUDE.md`
- `graphify-out/GRAPH_REPORT.md`
- `fbdi_applaud_mapping.xlsx` (Sheet2 — inventory the 183 rows and current state)
- `FBDI_Master_Catalog.xlsx` (inventory tabs/files/columns available for matching)

### Step 1 — Brainstorm the Audit Approach

**Invoke the `brainstorming` skill (superpowers).**

Specifically scope the brainstorm to:
- How to structure the audit loop (per-table vs. batched by FBDI template)
- What signals from `applaud-mcp` most strongly indicate a correct FBDI match (column-name overlap, IF patterns, EF patterns, data-type compatibility)
- How to rank confidence (what makes a match High vs. Medium vs. Low)
- How to cleanly represent evidence in the xlsx

Output: an approach document in the conversation before writing the plan.

### Follow the rest of the superpowers path, writing-plans, etc.


### Step X — Keep Graphify Current

After xlsx and doc changes:

```bash
python3 -c "from graphify.watch import _rebuild_code; from pathlib import Path; _rebuild_code(Path('.'))"
```

---

## Plugin & Skill Prescriptions

| Step | Invoke |
|------|--------|
| 1 | `brainstorming` skill (superpowers) |
| 2 | `writing-plans` skill (superpowers) |
| 3 | `executing-plans` skill (superpowers); `applaud-mcp` tools throughout |
| 6 | `verification-before-completion` skill (superpowers) |
| 7 | `commit-commands` plugin |
| Throughout | `context7` plugin if openpyxl edge cases arise |


---

## Notes for Claude Code

- `applaud-mcp` is the highest-signal source for Applaud table intent. Prefer its evidence over assumptions from prior mappings.
- `FBDI_Master_Catalog.xlsx` is the ground truth for what FBDI tabs and columns exist in the current release. If a tab isn't in the catalog, it isn't a valid mapping target.
- When in doubt on a specific mapping, `UNMAPPED` + strong reasoning in the audit report is the correct call. Brad would rather review 20 genuine ambiguities than chase down 5 forced low-confidence matches.
- Do not invent columns, tabs, or Applaud tables. If `applaud-mcp` or the catalog doesn't confirm it, it doesn't exist.