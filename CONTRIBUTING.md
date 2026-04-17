# Contributing to TIBMM

TIBMM is a framework-as-code project. Every visible change to the workbook
comes from a change to one of three upstream sources — the rubric, the
questionnaire, or the generator. Please do not hand-edit the generated
workbook.

## What the sources of truth are

| Change you want to make | Edit this file | Regenerate |
|---|---|---|
| Add / remove / reword a criterion | `src/rubric/rubric.yaml` | yes |
| Change a maturity-level anchor (text or quantitative threshold) | `src/rubric/rubric.yaml` | yes |
| Add / remove / reword an assessment question | `src/questionnaire/questionnaire.yaml` | yes |
| Change the scoring logic, tier gate, or layout | `src/scorer/generate_spreadsheet.py` | yes |
| Add a glossary term | `src/scorer/generate_spreadsheet.py` (the `GLOSSARY` list) | yes |
| Change the PPTX snapshot layout | `src/scorer/generate_report.js` | no — rerun on demand |
| Change the PPTX trend layout | `src/scorer/generate_trend.js` | no — rerun on demand |

To regenerate the workbook after a YAML or generator change:

```bash
python src/scorer/generate_spreadsheet.py \
  --output src/templates/TIBMM-Assessment-Workbook.xlsx
```

## Pull request guidelines

- Keep PRs small and focused. A criterion change, a new glossary term, and a
  layout tweak should all be separate PRs.
- Include the regenerated workbook in the same commit as the YAML change so
  reviewers can open the xlsx without rebuilding locally. The regeneration is
  deterministic — no manual touch-ups should be needed.
- When adding or reshaping a criterion, update the narrative dashboard
  explanation in `generate_spreadsheet.py` (search for `tier_explanations`) so
  the Results Dashboard summary stays consistent.
- Behavior-anchored phrasing is expected: a practice statement must describe
  something an external observer could verify (an artifact, a cadence, a
  published document, a measured metric). "The team understands X" is not a
  practice statement.

## Discussion

For larger changes to the model (adding a tier, rebalancing criteria,
extending the enrichment dimensions), open a GitHub issue first with the
problem statement, the proposed change, and the anchor references from the
literature. Opinionated changes require sourced justification — the model is
a synthesis, not a personal opinion.
