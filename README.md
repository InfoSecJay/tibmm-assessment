# TIBMM — Threat Intelligence Behavior Maturity Model

**A behavior-anchored maturity model for Cyber Threat Intelligence teams that operate alongside a Security Operations Center. Ships as a single Excel workbook plus a scripted reporting pipeline; tuned out of the box for a Canadian telecommunications operating context but industry-neutral in structure.**

TIBMM is a sibling to DEBMM (Detection Engineering Behavior Maturity Model). The two share workbook conventions, scoring scales, and reporting tooling so a security program can present both on a single executive scorecard.

**Status.** v0.1 — first public release.
**Licence.** MIT. See [`LICENSE`](./LICENSE).

---

## Quick start

```bash
# 1. Install dependencies (one-time)
python -m pip install -r src/scorer/requirements.txt
npm install

# 2. Regenerate the blank workbook (only needed after editing the rubric/questionnaire)
python src/scorer/generate_spreadsheet.py \
  --output src/templates/TIBMM-Assessment-Workbook.xlsx

# 3. Open src/templates/TIBMM-Assessment-Workbook.xlsx in Excel, fill 35 dropdowns, save.

# 4. Extract to JSON and append to a rolling history file
python src/scorer/extract_data.py src/templates/TIBMM-Assessment-Workbook.xlsx \
  -o reports/2026-04.json --history reports/history.json --date 2026-04

# 5. Produce the point-in-time report (4 slides)
node src/scorer/generate_report.js reports/2026-04.json reports/2026-04-snapshot.pptx

# 6. Produce the trend report (3 slides) — requires >=2 entries in history.json
node src/scorer/generate_trend.js reports/history.json reports/trend.pptx
```

---

## Model at a glance

**Maturity scale.** 1 (Initial) → 2 (Repeatable) → 3 (Defined) → 4 (Managed) → 5 (Optimized).

**5 Core Tiers — 20 criteria, 4 per tier:**

| Tier | Name | Focus |
|---|---|---|
| **T0** | Foundation | Mission/charter, IOC handling, source inventory, stakeholder touchpoints |
| **T1** | Basic | PIRs, BLUF product template, calibrated language, Collection Management Framework |
| **T2** | Operational | TIP + IOC pipeline, Detection Engineering handoff, Threat Hunt-team partnership, IR enrichment |
| **T3** | Advanced | Vulnerability & exploit intel, telco-fraud intel, Third-Party Risk, Cyber Threat Landscape reporting |
| **T4** | Leading | Strategic brief to CISO, Structured Analytic Techniques, outcome attribution, forecast calibration |

**2 Enrichment Dimensions — 6 criteria, 3 per dimension:**

- **Tradecraft & Analyst Capability** — ICD 203 adherence, ATT&CK/Diamond/Kill Chain literacy, peer review process
- **Governance, Legal & Improvement** — Canadian legal/regulatory review (PIPEDA, CASL, CRTC, Bill C-26 CCSPA), external sharing via CCTX / ISACs / trust groups, continuous improvement via repeated TIBMM self-assessment

**Tier advancement rule.** A tier is achieved only when *all* criteria in that tier *and all lower tiers* score ≥ 3.0 (Defined). This prevents claiming operational maturity while leaving foundations weak. Enrichment contributes to the overall score but does not gate tier determination — a strong tier score with a weak Tradecraft enrichment predicts fragility under scrutiny.

**Industry focus.** The framework structure is industry-neutral; the practice text is tuned for a Canadian telecommunications operating context. Expect references to the telco stack (RAN, core, SS7, Diameter, GTP, 5G SBA, OSS/BSS, CPE), telco fraud patterns (smishing, SIM swap, number-porting abuse, brand impersonation), carrier TPRM (roaming partners, interconnect, equipment vendors), and Canadian regulatory context (PIPEDA, CASL, CRTC, Bill C-26 CCSPA, CCCS advisories, CCTX). A non-telco or non-Canadian user can adapt the rubric YAML with minimal structural change — most edits are localized to three criteria at T3 (the stakeholder-specific layer).

---

## Research methodology

TIBMM is a synthesis, not an original framework. The design goal was to combine the strengths of the existing publicly-available Cyber Threat Intelligence maturity models while fixing two observed gaps: (a) most models either score the analytic-tradecraft floor or skip it entirely, and (b) most models under-weight the technical integration surface between CTI and the rest of a modern SOC (TIP, SIEM, SOAR, Detection-as-Code pipeline). The survey covered seven community frameworks and four pieces of supporting doctrine. Full source materials are bundled in [`references/`](./references/README.md).

### Sources surveyed and what TIBMM took from each

| Source | Contribution to TIBMM |
|---|---|
| [**CTI-CMM v1.3**](https://cti-cmm.org/) (Intel 471, IBM X-Force, Kroger, Mandiant, Bank of America + community; Jan 2026) | Behavior-anchored practice-statement style; 4-level conservative scale (adopted as 1–5 here); stakeholder-first philosophy; discouragement of pursuing the ceiling uniformly. |
| [**CREST CTI-MAT Detailed**](https://www.crest-approved.org/buying-building-cyber-services/cyber-threat-intelligence-maturity-assessment-tools/) (CREST, Apr 2022) | Explicit Legal/Regulatory/Ethics enrichment criterion; process-rigor decomposition. CREST's weighted-target-profile mechanic was *not* adopted — TIBMM's tier-gate rule achieves similar effect with less ceremony. |
| [**ThreatConnect TIMM**](https://threatconnect.com/resource/maturing-a-threat-intelligence-program/) (2017, refreshed 2025) | Narrative "what a typical team at this level looks like" descriptions on the Instructions tab and the Results Dashboard dynamic-tier-explanation block. The vendor-aligned TIP-adoption-as-maturity framing was explicitly *not* inherited. |
| [**Recorded Future TIMA**](https://www.recordedfuture.com/resources/maturity-assessment) | Four of the nine Recorded Future questions inform specific TIBMM criteria at T2 and T3 — most notably "response across the lifecycle," threat hunt cadence, automation footprint, and stack integration. |
| [**Mandiant Intelligence Capability Discovery**](https://cloud.google.com/blog/products/identity-security/cti-program-maturity-assessment/) (Feb 2024) | Promotion of "Analyst Capability and Expertise" from a soft skill to a first-class scored dimension. This became the Tradecraft enrichment dimension in TIBMM. |
| [**CTIM**](https://ctim.eu/) (Delft / Brainframe, academic) | The Generation-vs-Integration conceptual split, used as an implicit scoring lens in the IR and Detection handoff criteria — it is the reason the DE-handoff criterion measures *rule provenance tagging* alongside *request cadence*. |
| [**Slavkey CTI-MM**](https://github.com/Slavkey/CTI_Maturity_Model) | Strategic / Operational / Tactical coverage split, preserved as an implicit scoring lens inside the Cyber Threat Landscape Reporting criterion at T3. |
| [**SANS Intelligence Analyst's Playbook**](https://www.sans.org/posters/) (Matt Edmondson, CD_OSINT_02-26) | The entire T1 tradecraft floor: BLUF, calibrated probability language, Admiralty / NATO source rating, CRAAP source evaluation. Also the SAT expectation at T4 and cognitive-bias-countermeasure awareness in the Tradecraft enrichment. |
| [**ICD 203**](https://www.dni.gov/files/documents/ICD/ICD-203.pdf) (ODNI, 2015; technical amendments 2022) | The nine analytic tradecraft standards that anchor the `icd203_adherence` enrichment criterion. TIBMM's T1 calibrated-language, source-rating, and T4 SAT criteria all map back to ICD 203 standards 1–4; T4 forecast calibration maps to standard 8. |
| [**TaHiTI**](https://www.betaalvereniging.nl/en/safety/tahiti/) (FI-ISAC / Dutch Payments Association) | The entire structure of the T2 "Collaboration with the Threat Hunting Team" criterion, including the three-phase Initialize/Hunt/Finalize flow and the four sources of hunt hypotheses. This is why TIBMM measures CTI–Hunt *partnership* rather than asking whether CTI runs hunts itself. |
| [**MITRE ATT&CK**](https://attack.mitre.org/) | Listed as a baseline framework literacy expectation in the Tradecraft enrichment and as a scoped-artifact expectation in several criteria (hunt requests, IR enrichment, CTL reports, detection requests). |
| [**Lockheed Martin Cyber Kill Chain** / **Diamond Model of Intrusion Analysis**](./references/README.md) | Both framework literacies are scored alongside ATT&CK in the `framework_literacy` criterion. The three are treated as complementary rather than competing. |

### Design decisions that are not inherited from any single source

Three structural choices are TIBMM-original:

1. **Tier-based progression with per-criterion floor gate rather than per-domain target profile.** Most surveyed models either (a) use an 11–13 stakeholder-domain structure (CTI-CMM, CREST) with targets per domain, or (b) use an intelligence-cycle-phase structure (Slavkey, CTIM). TIBMM uses a tier progression (T0→T4) with a strict "all criteria in this tier and below must be ≥ 3.0" rule. This mirrors DEBMM's conventions and makes single-number maturity reporting honest — you cannot be "partially T3" by being strong in some domains and weak in others.

2. **Enrichment dimensions that do not gate tier advancement.** Tradecraft and Governance scores contribute to the overall maturity score but are structurally separate from the tier gate. This acknowledges that a CTI team can be structurally integrated and operationally fluid while still being fragile on tradecraft or governance — and the assessment surfaces that as an explicit gap rather than hiding it inside a domain average.

3. **CTI-to-Hunt-team collaboration rather than CTI-runs-hunts.** Several surveyed models bundle threat hunting with CTI or treat it as a CTI responsibility. TIBMM assumes a dedicated Threat Hunting team exists (common in mature SOCs) and measures the quality of CTI's partnership with that team — the hypothesis feed, the enriched intel package, and the findings return path — directly from TaHiTI's three-phase model.

### What was *not* adopted, and why

- **Domain-stakeholder structure (CTI-CMM, CREST).** Useful for large organizations with differentiated domain leadership. Unnecessary complexity for a small team; produces a score that is harder to communicate to leadership in one number.
- **CMMI-style levels (Mandiant ICD).** The five CMMI labels (Initial, Managed, Defined, Quantitatively Managed, Adaptive) are familiar to engineering audiences but produce a "Quantitatively Managed" label that is easier to aspire to than to audit. TIBMM keeps the 1–5 scale but uses Initial / Repeatable / Defined / Managed / Optimized with quantitative anchors at each level.
- **Vendor-platform-specific practices (TIMM).** TIBMM is platform-neutral. A team using open-source tooling (OpenCTI, MISP, Elastic SIEM) and a team using commercial tooling (ThreatConnect / Anomali / Recorded Future + Splunk) should be scorable on equivalent terms.
- **Workforce-screening / HR-vetting criteria (present in several community models).** Not in scope for a small CTI team. TIBMM scopes "workforce" to analyst tradecraft, framework literacy, and peer review — the things a CTI lead can actually influence.

---

## Repository layout

```
tibmm_assessment/
├── README.md                           # this file
├── LICENSE                             # MIT
├── CONTRIBUTING.md                     # contribution guidance (edit the YAML, don't edit the xlsx)
├── CHANGELOG.md                        # versioned changes
├── package.json                        # Node dependencies (pptxgenjs for PPTX output)
├── references/                         # third-party source material + links
│   ├── README.md
│   ├── CTI-CMM Assessment Tool v1.3.xlsx
│   ├── CTI-MMAT-Detailed-level_Apr2022.xlsx
│   ├── CMM_maturity_model.xlsx
│   ├── SANS-The-Intelligence-Analysts-Playbook-030926.pdf
│   └── recorded_future_cti_mm.txt
└── src/
    ├── rubric/rubric.yaml                          # 5 tiers × 4 criteria + 2 enrichment × 3 criteria
    ├── questionnaire/questionnaire.yaml            # 35 questions (1–5 scale or Yes/No)
    ├── scorer/
    │   ├── generate_spreadsheet.py                 # openpyxl workbook generator
    │   ├── extract_data.py                         # xlsx → JSON extractor
    │   ├── generate_report.js                      # point-in-time PPTX (4 slides)
    │   ├── generate_trend.js                       # over-time PPTX (3 slides)
    │   ├── _test_fixture.py                        # synthetic JSON for smoke tests
    │   └── requirements.txt
    └── templates/TIBMM-Assessment-Workbook.xlsx    # canonical generated workbook
```

---

## Workbook tabs

The generated workbook has 8 tabs:

1. **Instructions** — how to use, tier definitions, scoring rules, telco-operational context
2. **Assessment** — metadata + 35 dropdown questions (amber answer cells, auto-calculated score column)
3. **Results Dashboard** — overall score hero, achieved TIBMM tier, per-criterion heatmap, dynamic narrative explanation of the achieved tier
4. **Tier Scores Chart** — bar chart per core tier with 3.0 target reference
5. **Readiness Chart** — bar chart per enrichment dimension
6. **Rubric Reference** — full 1–5 maturity-level descriptions and quantitative anchors for every criterion
7. **Glossary** — definitions of every term a TIBMM question references, grouped by category (Direction, Tradecraft, Collection, Tooling, Framework, Hunt Methodology, Telco Stack, Telco Fraud, Canadian Regulatory, Sharing, Metrics, Reporting)
8. **Report Data** — machine-readable flat table consumed by the PPTX generators (summary, tier progression, criterion breakdown)

---

## Recommended usage cadence

- **Full TIBMM self-assessment:** bi-annual (every 6 months)
- **Quarterly KPI check-in** between full runs (tracks operational metrics without re-running the full assessment)
- **Suggested starting target profile for a new CTI function:** T2 (Operational) across the board, with T3 (Advanced) pursued in Fraud and Cyber Threat Landscape reporting first
- **Smoke test (no real data):**
  ```bash
  python src/scorer/_test_fixture.py --out-dir testdata
  node src/scorer/generate_report.js testdata/baseline.json testdata/snapshot.pptx
  node src/scorer/generate_trend.js testdata/history.json testdata/trend.pptx
  ```

---

## Prerequisites

- Python 3.10+ (for the generator and extractor)
- Node.js 18+ (for the PPTX report generators)
- Microsoft Excel — the extractor reads evaluated formula values, which requires the file to have been opened and saved in Excel at least once after answering questions. LibreOffice Calc produces acceptable output for most sheets but Excel is the canonical target.

---

## See also

- [`CONTRIBUTING.md`](./CONTRIBUTING.md) — how to make changes; every edit starts in the YAML, not the xlsx
- [`references/README.md`](./references/README.md) — full source list with links
- [`CHANGELOG.md`](./CHANGELOG.md) — version history
