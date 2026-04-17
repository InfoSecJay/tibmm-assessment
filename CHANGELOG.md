# Changelog

All notable changes to TIBMM will be documented in this file. The format is
loosely based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); the
project uses [Semantic Versioning](https://semver.org/).

## [0.1.0] — 2026-04-16

First public release.

### Added
- Rubric (5 core tiers × 4 criteria + 2 enrichment dimensions × 3 criteria) in
  `src/rubric/rubric.yaml`
- Questionnaire (35 dropdown questions; 1–5 scale and Yes/No) in
  `src/questionnaire/questionnaire.yaml`
- Python openpyxl workbook generator (`src/scorer/generate_spreadsheet.py`)
  producing an 8-tab xlsx: Instructions, Assessment, Results Dashboard, Tier
  Scores Chart, Readiness Chart, Rubric Reference, Glossary, Report Data
- Python xlsx → JSON extractor (`src/scorer/extract_data.py`) with optional
  `--history` upsert for longitudinal reporting
- Node / pptxgenjs point-in-time report generator (`src/scorer/generate_report.js`)
  producing a 4-slide PPTX
- Node / pptxgenjs trend report generator (`src/scorer/generate_trend.js`)
  producing a 3-slide PPTX from `history.json`
- Glossary tab embedded in workbook covering CTI tradecraft, telco stack, and
  Canadian regulatory terms
- Synthetic test fixture generator (`src/scorer/_test_fixture.py`) for
  smoke-testing the PPTX pipeline without filling a real assessment
- `references/` directory holding the source maturity models surveyed during
  synthesis (CTI-CMM, CREST CTI-MAT Detailed, SANS Playbook, ThreatConnect
  TIMM, Recorded Future TIMA, CMM Maturity Model), each retaining its own
  license

### Scoring design
- Tier advancement gate: a tier is achieved only when all criteria in that
  tier and all lower tiers score ≥ 3.0 (Defined)
- Enrichment dimensions contribute to the overall score but do not gate tier
  determination
- Conditional formatting bands: ≥4.5 green, 3.5–4.49 cyan, 3.0–3.49 yellow,
  2.0–2.99 orange, <2.0 red
