# Skill Evals — fbdi-compare-release

Per spec §8 Layer 2, these prompts exercise the skill's triggering and
end-to-end behavior.

## Running manually

Use `skill-creator`'s eval-viewer loop:

```bash
python ~/.claude/plugins/cache/claude-plugins-official/skill-creator/unknown/skills/skill-creator/scripts/run_eval.py \
  --skill-dir .claude/skills/fbdi-compare-release \
  --prompts .claude/skills/fbdi-compare-release/evals/prompts.jsonl \
  --output .claude/skills/fbdi-compare-release/evals/results/
```

Then open the generated HTML report:
```bash
python ~/.claude/plugins/cache/claude-plugins-official/skill-creator/unknown/skills/skill-creator/eval-viewer/generate_review.py \
  .claude/skills/fbdi-compare-release/evals/results/
```

## Ground-truth reference (eval #2)

Eval #2 ("Compare 26A to 26B") expects the skill to reproduce the
2026-04-23 end-to-end run:
- `Comparison_Report_26A_26B.xlsx`: 706 change rows, 19 files
- `FBDI_Master_Catalog.xlsx`: 9 Issues-tab rows, 748 Drift rows
- `baselines/26A/originals/`: 212 files
- `baselines/26B/originals/`: 213 files

These artifacts are already on disk and can be used as the oracle for
pass/fail scoring.

## Success criteria

- All `expected_trigger: true` prompts invoke the skill.
- All `expected_trigger: false` prompts do NOT invoke the skill.
- Eval #2 produces a `Comparison_Report_26A_26B.xlsx` byte-equivalent (or
  row-count-equivalent) to the committed ground truth.

If triggering misfires, move to Layer 3 (description optimization — see
Task 12).