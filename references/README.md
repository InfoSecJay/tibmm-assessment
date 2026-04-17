# References

This directory bundles the third-party maturity models, assessment tools,
and tradecraft doctrine that TIBMM was synthesized from. Every tier, every
criterion, and every practice statement in TIBMM traces back to at least one
of these sources.

The files are kept here for research-traceability: if you want to see *why*
TIBMM structures things the way it does, or *what specifically* TIBMM
borrowed, these are the originals. The main [README](../README.md#research-methodology)
explains which model contributed which design decision.

> **Licensing.** Each work in this directory retains its own license as
> published by its original author. This repository holds them purely as
> reference copies for research purposes. Links to the canonical sources are
> below — please obtain current versions from the original authors rather than
> relying on the files here.

---

## Cyber Threat Intelligence maturity frameworks

### CTI-CMM — Cyber Threat Intelligence Capability Maturity Model

- **File:** [`CTI-CMM Assessment Tool v1.3.xlsx`](./CTI-CMM%20Assessment%20Tool%20v1.3.xlsx)
- **Canonical source:** https://cti-cmm.org/
- **Authors:** Intel 471, IBM X-Force, Kroger, Mandiant, Bank of America, and
  a volunteer community of ~27 CTI practitioners.
- **What it is:** 11 stakeholder domains × 2–5 objectives × 4 maturity levels
  (CTI0 Pre-Foundational → CTI3 Leading) with ~230 practice statements.
  Inspired by the US DOE Cybersecurity Capability Maturity Model (C2M2).
- **What TIBMM took from it:** the behavior-anchored practice-statement style,
  the 4-level conservative scale, the discouragement of pursuing the ceiling
  uniformly, the stakeholder-first philosophy, and the scoring scale (0–3 per
  practice, promoted to 1–5 in TIBMM).

### CREST CTI Maturity Assessment Tool (CTI-MAT) — Detailed level

- **File:** [`CTI-MMAT-Detailed-level_Apr2022.xlsx`](./CTI-MMAT-Detailed-level_Apr2022.xlsx)
- **Canonical source:** https://www.crest-approved.org/buying-building-cyber-services/cyber-threat-intelligence-maturity-assessment-tools/
  (detailed/intermediate/summary variants) — also archived by Curated Intelligence.
- **Author:** CREST (Council of Registered Ethical Security Testers).
- **What it is:** 4 phases × 18 steps × 5 maturity levels (1 least effective
  through 5 most effective), with Governance, Program Planning and
  Requirements, Threat Intelligence Operation, and Functional Management as
  the four phases. Adopted by several European regulators (e.g., the Bank of
  England STAR-FS regime) as a maturity KPI.
- **What TIBMM took from it:** the explicit Legal/Regulatory/Ethics objective,
  granular process-rigor decomposition, target-profile mechanic (TIBMM
  simplifies this to the tier-gate rule).

### ThreatConnect TIMM (Threat Intelligence Maturity Model)

- **File:** [`CMM_maturity_model.xlsx`](./CMM_maturity_model.xlsx) (partial
  representation of several community models; TIMM narratives extracted from
  the ThreatConnect whitepaper below)
- **Canonical source:** https://threatconnect.com/resource/maturing-a-threat-intelligence-program/
  and https://threatconnect.com/wp-content/uploads/Maturity-Model-Whitepaper-2017-1.pdf
- **Author:** ThreatConnect.
- **What it is:** 5 levels — L0 Unclear Where to Start, L1 Warming Up, L2
  Expanding, L3 Driving Defense, L4 Well-Defined — each described
  narratively ("a typical team at this level looks like…") rather than as a
  question set. Vendor-oriented (TIP adoption as a maturity marker) but
  unusually strong on executive-communicable narrative.
- **What TIBMM took from it:** the "what a typical team at this level looks
  like" narrative descriptions on the Instructions and Results Dashboard tabs.

### Recorded Future TIMA (Threat Intelligence Maturity Assessment)

- **File:** [`recorded_future_cti_mm.txt`](./recorded_future_cti_mm.txt)
- **Canonical source:** https://www.recordedfuture.com/resources/maturity-assessment
- **Author:** Recorded Future.
- **What it is:** A 9-question, 4-option-per-question triage designed to be
  completable in 5 minutes. Covers resourcing, stakeholder service, personnel,
  strategic alignment, lifecycle, threat hunt cadence, data sources,
  automation, and stack integration.
- **What TIBMM took from it:** question 3 (response-across-the-lifecycle),
  question 6 (threat hunt cadence), question 8 (where have you automated),
  and question 9 (stack integration) each inform a specific TIBMM criterion
  in T2 and T3.

### Other maturity-model research surveyed (not in this directory)

- **Mandiant Intelligence Capability Discovery (ICD).** Web-based, gated,
  Google-account-required; 42 questions × 6 capability areas × 5 CMMI-style
  levels. https://cloud.google.com/blog/products/identity-security/cti-program-maturity-assessment/
  TIBMM promoted Mandiant's "Analyst Capability and Expertise" dimension into
  a first-class enrichment criterion.
- **CTIM (Delft / Brainframe).** Academically-grounded research instrument.
  https://ctim.eu/  TIBMM borrowed the Generation-vs-Integration conceptual
  split as a scoring lens inside the IR and Detection handoff criteria.
- **Slavkey CTI-MM.** Open-source practitioner contribution.
  https://github.com/Slavkey/CTI_Maturity_Model  TIBMM preserved the
  Strategic/Operational/Tactical coverage split as an implicit scoring lens
  in the Cyber Threat Landscape Reporting criterion.

---

## Analytic tradecraft doctrine

### SANS Intelligence Analyst's Playbook

- **File:** [`SANS-The-Intelligence-Analysts-Playbook-030926.pdf`](./SANS-The-Intelligence-Analysts-Playbook-030926.pdf)
- **Canonical source:** https://www.sans.org/posters/  (current SANS OSINT
  poster series; author Matt Edmondson, code CD_OSINT_02-26)
- **What it is:** A two-page reference poster codifying the day-to-day
  tradecraft of structured intelligence analysis: the intelligence cycle,
  Intelligence Product Structure (BLUF through Outlook), calibrated
  probability language, Admiralty / NATO source-evaluation, the CRAAP test,
  Analysis of Competing Hypotheses, STEEP+S scanning, and the catalog of
  cognitive biases with structured countermeasures.
- **What TIBMM took from it:** the entire Tier-1 tradecraft floor —
  BLUF, calibrated language, source reliability — plus the SAT expectation
  at Tier 4 and the bias-countermeasure awareness in the Tradecraft
  enrichment dimension.

### ICD 203 — Intelligence Community Directive 203

- **Canonical source:** https://www.dni.gov/files/documents/ICD/ICD-203.pdf
  (ODNI, 2 Jan 2015, technical amendments 2022)
- **What it is:** The US Intelligence Community's nine analytic tradecraft
  standards that all finished intelligence products must "implement and
  exhibit": (1) describe source quality, (2) express uncertainty with
  calibrated language, (3) distinguish information from assumption, (4)
  incorporate analysis of alternatives, (5) demonstrate customer relevance,
  (6) use clear argumentation, (7) explain change or consistency, (8) make
  accurate judgments, (9) use effective visuals.
- **What TIBMM took from it:** the nine standards are the referenced quality
  floor for the `icd203_adherence` enrichment criterion. TIBMM's Tier-1
  calibrated-language and source-rating criteria and Tier-4 SAT criterion
  all map back to ICD 203 standards 1–4.

### TaHiTI — Targeted Hunting integrating Threat Intelligence

- **Canonical source:** https://www.betaalvereniging.nl/en/safety/tahiti/
  (FI-ISAC / Dutch Payments Association)
- **What it is:** A three-phase threat-hunting methodology — Initialize,
  Hunt, Finalize — explicitly designed around CTI integration. Hunt
  hypotheses are generated from four sources (CTI, IR findings, red team,
  MaGMa framework gaps) and CTI provides enriched adversary packages to
  scope the hunt.
- **What TIBMM took from it:** the full structure of the T2 "Collaboration
  with the Threat Hunting Team" criterion. TaHiTI is the reason TIBMM
  measures *partnership* between CTI and a separate Hunt team rather than
  asking whether CTI runs hunts itself.

### MITRE ATT&CK

- **Canonical source:** https://attack.mitre.org/
- **What it is:** The de facto taxonomy of adversary tactics, techniques,
  and sub-techniques. Current v18 (2025) has 216 techniques and 475
  sub-techniques across 14 tactics.
- **How it shows up in TIBMM:** ATT&CK literacy is a required enrichment
  competency; several criteria (hunt scoping, IR enrichment, CTL reporting,
  detection request templates) expect ATT&CK IDs in the product.

### Lockheed Martin Cyber Kill Chain and the Diamond Model of Intrusion Analysis

- **Canonical sources:**
  - Kill Chain: Hutchins, Cloppert, Amin (Lockheed Martin, 2011) —
    https://www.lockheedmartin.com/en-us/capabilities/cyber/cyber-kill-chain.html
  - Diamond Model: Caltagirone, Pendergast, Betz (2013), with the 2025
    Cisco Talos Relationship Layer extension
- **How they show up in TIBMM:** both are listed alongside ATT&CK as the
  frameworks scored in the `framework_literacy` enrichment criterion. The
  three are treated as complementary, not competing.

---

## How this directory is used

During a TIBMM assessment, you generally do not need to open the files in
this directory. They are here for:

1. **Reviewers** who want to audit the derivation of a specific TIBMM
   criterion ("where did this come from?").
2. **Framework maintainers** proposing changes — see
   [`CONTRIBUTING.md`](../CONTRIBUTING.md), which expects sourced
   justification from this set of references.
3. **Researchers** surveying the CTI maturity-model landscape who want one
   place to compare the five major community frameworks side-by-side.

If you would like to cite TIBMM in academic work, please cite the original
sources alongside TIBMM itself — the synthesis is opinionated but unoriginal
in its constituents.
