"""Generate synthetic TIBMM assessment data for testing the report generators.

Produces two JSON files: a baseline snapshot and a two-period history.
Usage:
    python scorer/_test_fixture.py [--out-dir testdata]
"""

import argparse
import json
from pathlib import Path

# Target tier + criterion structure must mirror what the real Report Data tab emits.
# Keeping this list in sync with rubric.yaml is a manual task — only used for
# testing the snapshot / trend layout, not for real scoring.

CORE_TIERS = [
    ("T0", "Foundation", [
        "Mission, Charter, and Service Catalog",
        "Indicator of Compromise Handling",
        "Intelligence Source Inventory",
        "Stakeholder Awareness and Access",
    ]),
    ("T1", "Basic", [
        "Priority Intelligence Requirements (PIRs)",
        "Product Template and BLUF Structure",
        "Calibrated Language and Source Reliability",
        "Collection Management Framework (CMF)",
    ]),
    ("T2", "Operational", [
        "Threat Intelligence Platform and IOC Pipeline",
        "Detection Engineering Handoff",
        "Collaboration with the Threat Hunting Team",
        "Incident Response Enrichment",
    ]),
    ("T3", "Advanced", [
        "Vulnerability and Exploit Intelligence",
        "Telco Fraud Intelligence",
        "Third-Party and Supply-Chain Intelligence",
        "Cyber Threat Landscape Reporting",
    ]),
    ("T4", "Leading", [
        "Strategic Intelligence to Leadership",
        "Structured Analytic Techniques (SATs)",
        "Attributable Outcomes and Impact Metrics",
        "Forecast Calibration",
    ]),
]

ENRICHMENT = [
    ("Tradecraft and Analyst Capability", [
        "ICD 203 Analytic Standards Adherence",
        "Framework Literacy (ATT&CK, Diamond, Kill Chain)",
        "Peer Review Process",
    ]),
    ("Governance, Legal, Ethics, and Improvement", [
        "Legal and Regulatory Review (PIPEDA, CASL, CRTC)",
        "External Sharing and Trust-Group Participation",
        "Continuous Improvement and Self-Assessment",
    ]),
]


def level_for(score):
    if score >= 4.5:
        return "Optimized"
    if score >= 3.5:
        return "Managed"
    if score >= 2.5:
        return "Defined"
    if score >= 1.5:
        return "Repeatable"
    return "Initial"


def status_for(score):
    return "\u2713 Pass" if score >= 3 else "\u2717 Below Target"


def progression_for(tier_idx, tier_scores):
    """Complete = this and all lower tiers pass; Current = previous pass; else In Progress."""
    if all(s >= 3 for s in tier_scores[: tier_idx + 1]):
        return "Complete"
    if tier_idx == 0 or all(s >= 3 for s in tier_scores[:tier_idx]):
        return "Current"
    return "In Progress"


def snapshot(org, date, core_scores, enrichment_scores):
    """Build a Report-Data-shaped snapshot dict.

    core_scores: list of 20 floats (4 per tier × 5 tiers)
    enrichment_scores: list of 6 floats (3 per dimension × 2 dimensions)
    """
    criteria = []
    tier_avgs = []
    ci = 0
    for tid, tname, crits in CORE_TIERS:
        tier_crit_scores = []
        for crit in crits:
            s = core_scores[ci]
            criteria.append({
                "section": "TIBMM Core",
                "category": tname,
                "criterion": crit,
                "score": s,
                "level": level_for(s),
                "status": status_for(s),
            })
            tier_crit_scores.append(s)
            ci += 1
        tier_avgs.append(sum(tier_crit_scores) / len(tier_crit_scores))

    # Enrichment
    ei = 0
    for cat, crits in ENRICHMENT:
        for crit in crits:
            s = enrichment_scores[ei]
            criteria.append({
                "section": "Enrichment",
                "category": cat,
                "criterion": crit,
                "score": s,
                "level": level_for(s),
                "status": status_for(s),
            })
            ei += 1

    # Build tier rows with progression
    tiers = []
    tier_display = {
        "T0": "Tier 0: Foundation", "T1": "Tier 1: Basic",
        "T2": "Tier 2: Operational", "T3": "Tier 3: Advanced",
        "T4": "Tier 4: Leading",
    }

    # Determine achieved tier (highest where all tiers up to it fully pass at criterion level)
    achieved_tier = "Below Foundation"
    for i, (tid, tname, crits) in enumerate(CORE_TIERS):
        crit_slice = core_scores[i * 4: (i + 1) * 4]
        lower_pass = all(core_scores[j] >= 3 for j in range(0, (i + 1) * 4))
        if lower_pass:
            achieved_tier = tier_display[tid]

    for i, (tid, tname, crits) in enumerate(CORE_TIERS):
        tiers.append({
            "id": tid,
            "name": tname,
            "score": round(tier_avgs[i], 2),
            "level": level_for(tier_avgs[i]),
            "status": status_for(tier_avgs[i]),
            "progression": progression_for(i, tier_avgs),
        })

    overall = sum(c["score"] for c in criteria) / len(criteria)
    n_answered = len(core_scores) + len(enrichment_scores)

    return {
        "org": org,
        "assessor": "Jay Tymchuk",
        "date": date,
        "type": "Self-Assessment",
        "overallScore": round(overall, 2),
        "achievedTier": achieved_tier,
        "completion": f"35 / 35",
        "tiers": tiers,
        "criteria": criteria,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", type=Path, default=Path("testdata"))
    args = ap.parse_args()
    args.out_dir.mkdir(exist_ok=True, parents=True)

    # Baseline (2026-04): mostly Foundation/Basic done, Operational partial, higher tiers sparse
    baseline_core = [
        3.0, 3.5, 3.0, 3.0,      # T0 Foundation — passing
        3.0, 3.5, 3.0, 3.0,      # T1 Basic — passing
        2.0, 2.5, 1.5, 2.0,      # T2 Operational — gaps
        1.5, 1.0, 1.0, 1.5,      # T3 Advanced — weak
        1.0, 1.0, 1.0, 1.0,      # T4 Leading — not started
    ]
    baseline_enrichment = [3.0, 2.5, 2.0, 2.0, 2.0, 1.5]

    # Follow-up (2026-10): six months of uplift work
    followup_core = [
        3.5, 4.0, 3.5, 3.5,      # T0 Foundation — stronger
        3.5, 3.5, 3.5, 3.0,      # T1 Basic — solid
        3.0, 3.0, 3.0, 3.0,      # T2 Operational — now passing
        2.5, 2.5, 2.0, 2.5,      # T3 Advanced — emerging
        1.5, 1.5, 1.0, 1.0,      # T4 Leading — still early
    ]
    followup_enrichment = [3.5, 3.0, 3.0, 2.5, 3.0, 3.0]

    snap1 = snapshot("Canadian Telco", "2026-04-15", baseline_core, baseline_enrichment)
    snap2 = snapshot("Canadian Telco", "2026-10-15", followup_core, followup_enrichment)

    (args.out_dir / "baseline.json").write_text(
        json.dumps(snap1, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    # History: two entries with date in YYYY-MM format
    snap1_hist = {"date": "2026-04", **snap1}
    snap2_hist = {"date": "2026-10", **snap2}
    (args.out_dir / "history.json").write_text(
        json.dumps([snap1_hist, snap2_hist], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print(f"Wrote {args.out_dir / 'baseline.json'}")
    print(f"  org={snap1['org']} overall={snap1['overallScore']} tier={snap1['achievedTier']}")
    print(f"Wrote {args.out_dir / 'history.json'}  (2 periods)")
    print(f"  latest overall={snap2['overallScore']} tier={snap2['achievedTier']}")


if __name__ == "__main__":
    main()
