# Oracle Quarterly Release Version Format

Oracle Cloud Fusion Applications ships quarterly releases. Each release
has a two-character label: `YYx` where `YY` is a two-digit year and `x` is
one of `A`, `B`, `C`, `D` (Q1–Q4 respectively).

## Examples

| Label | Period                |
|-------|-----------------------|
| 25A   | Feb 2025 – Apr 2025  |
| 25B   | May 2025 – Jul 2025  |
| 25C   | Aug 2025 – Oct 2025  |
| 25D   | Nov 2025 – Jan 2026  |
| 26A   | Feb 2026 – Apr 2026  |
| 26B   | May 2026 – Jul 2026  |

Quarterly cadence is stable — Oracle has not skipped or renamed these in
the years leading up to this skill's authorship (2026-04).

## Canonical form

In this repo, release labels are **uppercase** everywhere they are
user-visible (folders `baselines/26A/`, tab names in catalog workbook,
`baseline_files.txt` section headers, comparison report filenames). The
CLI accepts any case and upper-cases internally.

## How to find the latest Oracle release

1. Visit https://docs.oracle.com/en/cloud/saas/ — Oracle's Cloud SaaS
   landing page.
2. Pick a module (Financials, Procurement, Project Management, Supply
   Chain) — each lists "What's New" for the current release at the top.
3. The current-release URL pattern is:
   `https://docs.oracle.com/en/cloud/saas/<module>/<release_lowercase>/oe<code>/index.html`
   (e.g., `.../financials/26b/oefbf/index.html`).

## Expected release count (historical)

| Release | File count (originals) |
|---------|------------------------|
| 26A     | 212                    |
| 26B     | 213 (added `ItemImportReferenceOrgTemplate.xlsm`) |

Oracle rarely adds or removes more than ~5–10 templates in a quarterly
release — §5 #6's 15% delta guard in `verify_download.py` catches larger
swings on the first run of a new release.
