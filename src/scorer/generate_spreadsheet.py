#!/usr/bin/env python3
"""
TIBMM Assessment Spreadsheet Generator

Generates an all-in-one Excel workbook for the Threat Intelligence Behavior
Maturity Model (TIBMM). Mirrors the DEBMM workbook structure so that both
assessments can share reporting tooling.

Tabs produced:
  1. Instructions
  2. Assessment (fillable with dropdowns and auto-scoring)
  3. Results Dashboard (auto-calculated scores and narrative)
  4. Tier Scores Chart
  5. Readiness Chart
  6. Rubric Reference
  7. Glossary (definitions for CTI, tradecraft, telco, and regulatory terms)
  8. Report Data (machine-readable — consumed by PDF/PPT generator)

Usage:
    python generate_spreadsheet.py [--output TIBMM-Assessment-Workbook.xlsx]
    python generate_spreadsheet.py --mode audit --output TIBMM-Audit.xlsx
"""

import argparse
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
DEFAULT_RUBRIC = PROJECT_ROOT / "rubric" / "rubric.yaml"
DEFAULT_QUESTIONNAIRE = PROJECT_ROOT / "questionnaire" / "questionnaire.yaml"

# ── Color palette ─────────────────────────────────────────────────────────────

# DEBMM Core — navy/blue family
DARK_NAVY = "0F1D32"
NAVY = "1B2A4A"
STEEL = "2D3E50"
MED_BLUE = "3B82F6"
LIGHT_BLUE_BG = "EBF4FF"
BLUE_ACCENT = "93C5FD"

# Enrichment — teal family
TEAL = "0D7377"
DARK_TEAL = "115E60"
LIGHT_TEAL_BG = "E6F7F7"
TEAL_ACCENT = "5EEAD4"

# Neutrals
WHITE = "FFFFFF"
OFF_WHITE = "FAFBFC"
LIGHT_GRAY = "F1F5F9"
HAIRLINE_COLOR = "E2E8F0"
DARK_TEXT = "1E293B"
MED_TEXT = "475569"

# Answer cells — warm amber (the only warm tone)
ANSWER_BG_COLOR = "FFFBEB"
ANSWER_BORDER_COLOR = "F59E0B"

# Conditional formatting
SCORE_GREEN = "D1FAE5"
SCORE_YELLOW = "FEF3C7"
SCORE_ORANGE = "FED7AA"
SCORE_RED = "FEE2E2"

# Rubric level tints (5 distinct)
LEVEL_1_BG = "FFF1F2"
LEVEL_2_BG = "FFF7ED"
LEVEL_3_BG = "FEFCE8"
LEVEL_4_BG = "F0FDF4"
LEVEL_5_BG = "ECFDF5"

# ── Font constants ────────────────────────────────────────────────────────────

FN = "Aptos"  # Modern Excel 365 default; graceful fallback to Calibri

FONT_TITLE = Font(name=FN, size=20, bold=True, color=WHITE)
FONT_SUBTITLE = Font(name=FN, size=10, italic=True, color=BLUE_ACCENT)
FONT_SECTION = Font(name=FN, size=12, bold=True, color=NAVY)
FONT_SECTION_TEAL = Font(name=FN, size=12, bold=True, color=DARK_TEAL)
FONT_TIER_BANNER = Font(name=FN, size=12, bold=True, color=WHITE)
FONT_COL_HEADER = Font(name=FN, size=10, bold=True, color=WHITE)
FONT_BODY = Font(name=FN, size=10.5, color=DARK_TEXT)
FONT_BODY_BOLD = Font(name=FN, size=10.5, bold=True, color=DARK_TEXT)
FONT_BODY_ITALIC = Font(name=FN, size=10.5, italic=True, color=MED_TEXT)
FONT_SMALL = Font(name=FN, size=9.5, color=MED_TEXT)
FONT_SMALL_ITALIC = Font(name=FN, size=9.5, italic=True, color=MED_TEXT)
FONT_ANSWER = Font(name=FN, size=11, bold=True, color=DARK_TEXT)
FONT_SCORE_HERO = Font(name=FN, size=28, bold=True, color=NAVY)
FONT_SCORE_LARGE = Font(name=FN, size=16, bold=True, color=NAVY)
FONT_SCORE_LABEL = Font(name=FN, size=10, bold=True, color=STEEL)
FONT_LEVEL_BOLD = Font(name=FN, size=10, bold=True, color=DARK_TEXT)
FONT_CONTEXT = Font(name=FN, size=9.5, italic=True, color=MED_TEXT)

# ── Fill constants ────────────────────────────────────────────────────────────

FILL_DARK_NAVY = PatternFill("solid", fgColor=DARK_NAVY)
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_STEEL = PatternFill("solid", fgColor=STEEL)
FILL_TEAL = PatternFill("solid", fgColor=TEAL)
FILL_WHITE = PatternFill("solid", fgColor=WHITE)
FILL_OFF_WHITE = PatternFill("solid", fgColor=OFF_WHITE)
FILL_LIGHT_GRAY = PatternFill("solid", fgColor=LIGHT_GRAY)
FILL_LIGHT_BLUE = PatternFill("solid", fgColor=LIGHT_BLUE_BG)
FILL_LIGHT_TEAL = PatternFill("solid", fgColor=LIGHT_TEAL_BG)
FILL_ANSWER = PatternFill("solid", fgColor=ANSWER_BG_COLOR)

FILL_SCORE_GREEN = PatternFill("solid", fgColor=SCORE_GREEN)
FILL_SCORE_YELLOW = PatternFill("solid", fgColor=SCORE_YELLOW)
FILL_SCORE_ORANGE = PatternFill("solid", fgColor=SCORE_ORANGE)
FILL_SCORE_RED = PatternFill("solid", fgColor=SCORE_RED)

FILL_LEVEL = {
    1: PatternFill("solid", fgColor=LEVEL_1_BG),
    2: PatternFill("solid", fgColor=LEVEL_2_BG),
    3: PatternFill("solid", fgColor=LEVEL_3_BG),
    4: PatternFill("solid", fgColor=LEVEL_4_BG),
    5: PatternFill("solid", fgColor=LEVEL_5_BG),
}

# ── Alignment constants ──────────────────────────────────────────────────────

ALIGN_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

# ── Border constants ──────────────────────────────────────────────────────────

HAIRLINE_BOTTOM = Border(bottom=Side(style="thin", color=HAIRLINE_COLOR))
THIN_BORDER = Border(
    left=Side(style="thin", color=HAIRLINE_COLOR),
    right=Side(style="thin", color=HAIRLINE_COLOR),
    top=Side(style="thin", color=HAIRLINE_COLOR),
    bottom=Side(style="thin", color=HAIRLINE_COLOR),
)
ANSWER_BORDER = Border(
    left=Side(style="thin", color=ANSWER_BORDER_COLOR),
    right=Side(style="thin", color=ANSWER_BORDER_COLOR),
    top=Side(style="thin", color=ANSWER_BORDER_COLOR),
    bottom=Side(style="thin", color=ANSWER_BORDER_COLOR),
)
BLUE_ACCENT_LEFT = Border(left=Side(style="medium", color=MED_BLUE))
TEAL_ACCENT_LEFT = Border(left=Side(style="medium", color=TEAL_ACCENT))


# ── Helpers ───────────────────────────────────────────────────────────────────


def load_yaml(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def style_cell(cell, font=None, fill=None, alignment=None, border=None):
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border


def style_range(ws, row, col_start, col_end, font=None, fill=None, alignment=None, border=None):
    for c in range(col_start, col_end + 1):
        style_cell(ws.cell(row=row, column=c), font, fill, alignment, border)


def is_enrichment(tier_id):
    return str(tier_id).startswith("enrichment")


def apply_conditional_formatting(ws, cell_range):
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="between", formula=["1", "1.49"], fill=FILL_SCORE_RED),
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="between", formula=["1.5", "2.49"], fill=FILL_SCORE_ORANGE),
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="between", formula=["2.5", "3.49"], fill=FILL_SCORE_YELLOW),
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="between", formula=["3.5", "5"], fill=FILL_SCORE_GREEN),
    )


def level_formula(score_ref):
    return (
        f'=IF({score_ref}="","",IF({score_ref}>=4.5,"Optimized",'
        f'IF({score_ref}>=3.5,"Managed",IF({score_ref}>=2.5,"Defined",'
        f'IF({score_ref}>=1.5,"Repeatable","Initial")))))'
    )


def status_formula(score_ref):
    return f'=IF({score_ref}="","",IF({score_ref}>=3,"\u2713 Pass","\u2717 Below Target"))'


# ── Tab 1: Instructions ──────────────────────────────────────────────────────


def build_instructions_tab(wb: Workbook, mode: str):
    ws = wb.active
    ws.title = "Instructions"
    ws.sheet_properties.tabColor = NAVY

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 4

    # ── Title banner ──────────────────────────────────────────────────
    ws.merge_cells("A1:D1")
    ws.row_dimensions[1].height = 56
    c = ws["A1"]
    c.value = "  TIBMM Assessment Tool"
    style_cell(c, FONT_TITLE, FILL_DARK_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCD":
        ws[f"{col}1"].fill = FILL_DARK_NAVY

    # Subtitle row
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 26
    mode_label = "Self-Assessment" if mode == "self" else "Audit Assessment"
    c = ws["A2"]
    c.value = f"  Threat Intelligence Behavior Maturity Model \u2014 {mode_label}"
    style_cell(c, FONT_SUBTITLE, FILL_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCD":
        ws[f"{col}2"].fill = FILL_NAVY

    row = 4

    def section_card(title, lines, accent="blue"):
        nonlocal row
        fill = FILL_LIGHT_BLUE if accent == "blue" else FILL_LIGHT_TEAL
        border = BLUE_ACCENT_LEFT if accent == "blue" else TEAL_ACCENT_LEFT
        font = FONT_SECTION if accent == "blue" else FONT_SECTION_TEAL

        ws.merge_cells(f"B{row}:C{row}")
        ws.cell(row=row, column=2, value=title)
        style_cell(ws.cell(row=row, column=2), font, fill, ALIGN_LEFT, border)
        ws.cell(row=row, column=3).fill = fill
        ws.row_dimensions[row].height = 28
        row += 1

        for line in lines:
            ws.merge_cells(f"B{row}:C{row}")
            ws.cell(row=row, column=2, value=line)
            style_cell(ws.cell(row=row, column=2), FONT_BODY, alignment=ALIGN_WRAP)
            row += 1

        row += 1  # Spacing

    # ── About ─────────────────────────────────────────────────────────
    section_card("About This Assessment", [
        "This workbook assesses the maturity of a Cyber Threat Intelligence (CTI) function "
        "using the Threat Intelligence Behavior Maturity Model (TIBMM). TIBMM is tuned "
        "for a Canadian telco CTI team operating alongside a SOC, and is a sibling to "
        "the Detection Engineering Behavior Maturity Model (DEBMM).",
        "",
        "It covers 26 criteria across 5 progressive tiers (T0\u2013T4) plus 2 enrichment "
        "dimensions (Tradecraft, Governance), with 35 dropdown questions total.",
        "All questions use dropdowns (Yes/No or Scale 1\u20145). No free-text required.",
    ])

    # ── How to Use ────────────────────────────────────────────────────
    section_card("How to Use", [
        "1.  Go to the Assessment tab and fill in your organization details",
        "2.  Answer each question using the dropdown menus in the amber column",
        "3.  For Yes/No questions \u2014 select from the dropdown",
        "4.  For Scale questions (1\u20145) \u2014 select your maturity rating",
        "5.  If any term in a question is unfamiliar (PIR, ICD 203, CMF, TaHiTI, "
        "Admiralty Code, etc.), open the Glossary tab for a short definition",
        "6.  Switch to the Results Dashboard tab to see scores calculated automatically",
    ])

    # ── Understanding the Model (two-column) ──────────────────────────
    ws.merge_cells(f"B{row}:C{row}")
    ws.cell(row=row, column=2, value="Understanding the Model")
    style_cell(ws.cell(row=row, column=2), FONT_SECTION, FILL_LIGHT_BLUE, ALIGN_LEFT, BLUE_ACCENT_LEFT)
    ws.cell(row=row, column=3).fill = FILL_LIGHT_BLUE
    ws.row_dimensions[row].height = 28
    row += 1

    # Left column header — TIBMM
    ws.cell(row=row, column=2, value="TIBMM Core Tiers")
    style_cell(ws.cell(row=row, column=2), FONT_BODY_BOLD, FILL_LIGHT_BLUE, ALIGN_LEFT, BLUE_ACCENT_LEFT)
    # Right column header — Enrichment
    ws.cell(row=row, column=3, value="Enrichment Dimensions")
    style_cell(ws.cell(row=row, column=3), Font(name=FN, size=10.5, bold=True, color=DARK_TEAL),
               FILL_LIGHT_TEAL, ALIGN_LEFT, TEAL_ACCENT_LEFT)
    ws.row_dimensions[row].height = 24
    row += 1

    core_lines = [
        "Tier 0 \u2014 Foundation",
        "    Role, IOC handling, source list, stakeholder touchpoints",
        "Tier 1 \u2014 Basic",
        "    PIRs, product template (BLUF), calibrated language, CMF",
        "Tier 2 \u2014 Operational",
        "    TIP + IOC pipeline, DE handoff, Hunt-team partnership, IR enrichment",
        "Tier 3 \u2014 Advanced",
        "    Vuln intel, telco fraud, TPRM, Cyber Threat Landscape reporting",
        "Tier 4 \u2014 Leading",
        "    Strategic intel, SATs, outcome attribution, forecast calibration",
    ]
    enrichment_lines = [
        "Tradecraft & Analyst Capability",
        "    ICD 203 adherence, framework literacy, peer review",
        "Governance, Legal & Improvement",
        "    PIPEDA / CRTC / Bill C-26 review, CCTX sharing, self-assessment",
        "",
        "These enrichment dimensions capture",
        "tradecraft and governance factors.",
        "They contribute to the overall score",
        "but do not affect TIBMM tier determination.",
        "",
    ]

    for i, (left, right) in enumerate(zip(core_lines, enrichment_lines)):
        ws.cell(row=row, column=2, value=left)
        style_cell(ws.cell(row=row, column=2), FONT_BODY, alignment=ALIGN_WRAP)
        ws.cell(row=row, column=3, value=right)
        style_cell(ws.cell(row=row, column=3), FONT_BODY, alignment=ALIGN_WRAP)
        row += 1

    row += 1  # Spacing

    # ── Maturity Levels ───────────────────────────────────────────────
    section_card("Maturity Levels", [
        "Each criterion is scored on a 1\u20145 scale:",
        "",
        "  1 \u2014 Initial          Minimal or no structured activity",
        "  2 \u2014 Repeatable    Sporadic, inconsistent efforts",
        "  3 \u2014 Defined         Regular, documented processes followed consistently",
        "  4 \u2014 Managed       Comprehensive, well-integrated with measurable outcomes",
        "  5 \u2014 Optimized     Fully automated, continuously improving",
    ])

    # ── Tier Determination ────────────────────────────────────────────
    section_card("Tier Determination", [
        "Your achieved TIBMM tier is the highest tier where ALL criteria in that tier "
        "(and all lower tiers) score \u2265 3.0 (Defined level). This enforces the progressive "
        "nature of the model \u2014 solid foundations must precede claiming advanced maturity.",
        "",
        "Enrichment dimensions (Tradecraft, Governance) contribute to the overall maturity "
        "score but are not part of the tier determination logic. A weak enrichment score "
        "with a strong tier score predicts fragility under scrutiny.",
    ])

    # ── Canadian Telco Context ────────────────────────────────────────
    section_card("Canadian Telco Context", [
        "TIBMM practice text is tuned to a Canadian telco CTI function. Expect references "
        "to PIPEDA, CASL, CRTC expectations, Bill C-26 (CCSPA), CCCS advisories, CCTX "
        "membership, telco-stack vulnerabilities (RAN, core, SS7, Diameter, GTP, 5G), and "
        "telco fraud patterns (smishing, SIM swap, number porting, brand impersonation). ",
        "",
        "The framework itself is industry-neutral; only examples are telco-specific.",
    ])

    # ── Scoring Paths ─────────────────────────────────────────────────
    section_card("Scoring Paths", [
        "1.  This Workbook \u2014 fill it out; scores calculate automatically in the Dashboard.",
        "2.  Report Data tab \u2014 machine-readable; consumed by the PDF/PPT generator.",
        "3.  Bi-annual cadence recommended; quarterly KPI check-ins between full runs.",
    ])

    # ── References ────────────────────────────────────────────────────
    section_card("References", [
        "CTI-CMM (Intel 471 community): https://cti-cmm.org/",
        "CREST CTI-MAT: https://www.crest-approved.org/",
        "SANS Intelligence Analyst's Playbook (Matt Edmondson, CD_OSINT_02-26)",
        "ICD 203 (ODNI, 2015): https://www.dni.gov/files/documents/ICD/ICD-203.pdf",
        "TaHiTI threat hunting methodology (FI-ISAC / Dutch Payments Association)",
        "MITRE ATT&CK: https://attack.mitre.org/",
    ])

    return ws


# ── Tab 2: Assessment ─────────────────────────────────────────────────────────


def build_assessment_tab(wb: Workbook, questionnaire: dict, rubric: dict, mode: str):
    ws = wb.create_sheet("Assessment")
    ws.sheet_properties.tabColor = MED_BLUE

    # Column layout — gutter | ID | Criterion | Question | Answer | Score | [Evidence] | gutter
    if mode == "audit":
        col_widths = {"A": 2, "B": 7, "C": 26, "D": 78, "E": 14, "F": 10, "G": 45, "H": 2}
        last_col = "H"
        last_col_num = 8
        evidence_col = 7  # Column G
    else:
        col_widths = {"A": 2, "B": 7, "C": 26, "D": 78, "E": 14, "F": 10, "G": 2}
        last_col = "G"
        last_col_num = 7
        evidence_col = None
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # ── Title banner ──────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    ws.row_dimensions[1].height = 52
    c = ws["A1"]
    c.value = "  TIBMM Assessment"
    style_cell(c, FONT_TITLE, FILL_DARK_NAVY, Alignment(horizontal="left", vertical="center"))
    for i in range(1, last_col_num + 1):
        ws.cell(row=1, column=i).fill = FILL_DARK_NAVY

    # Subtitle row
    ws.merge_cells(f"A2:{last_col}2")
    ws.row_dimensions[2].height = 24
    c = ws["A2"]
    c.value = "  Threat Intelligence Behavior Maturity Model"
    style_cell(c, FONT_SUBTITLE, FILL_NAVY, Alignment(horizontal="left", vertical="center"))
    for i in range(1, last_col_num + 1):
        ws.cell(row=2, column=i).fill = FILL_NAVY

    # ── Metadata section ──────────────────────────────────────────────
    ws.row_dimensions[3].height = 8  # Spacer
    labels = ["Organization:", "Assessor Name:", "Assessor Role:", "Date:", "Assessment Type:"]
    for i, label in enumerate(labels):
        r = 4 + i
        ws.cell(row=r, column=3, value=label)
        style_cell(ws.cell(row=r, column=3), FONT_BODY_BOLD, alignment=ALIGN_RIGHT)
        ws.merge_cells(f"D{r}:E{r}")
        style_cell(ws.cell(row=r, column=4), FONT_BODY, FILL_ANSWER, ALIGN_LEFT, ANSWER_BORDER)
        ws.cell(row=r, column=5).fill = FILL_ANSWER
        ws.cell(row=r, column=5).border = ANSWER_BORDER

    # Pre-fill assessment type
    ws.cell(row=8, column=4, value="Self-Assessment" if mode == "self" else "Audit")

    # Type dropdown
    dv_type = DataValidation(type="list", formula1='"Self-Assessment,Audit"', allow_blank=False)
    dv_type.error = "Please select Self-Assessment or Audit"
    dv_type.errorTitle = "Invalid Entry"
    ws.add_data_validation(dv_type)
    dv_type.add(ws.cell(row=8, column=4))

    # ── Data validations ──────────────────────────────────────────────
    dv_yesno = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv_yesno.error = "Please select Yes or No"
    dv_yesno.errorTitle = "Invalid Entry"
    dv_yesno.prompt = "Select Yes or No"
    dv_yesno.promptTitle = "Answer"
    ws.add_data_validation(dv_yesno)

    dv_scale = DataValidation(type="list", formula1='"1,2,3,4,5"', allow_blank=True)
    dv_scale.error = "Please select a value from 1 to 5"
    dv_scale.errorTitle = "Invalid Entry"
    dv_scale.prompt = "Select 1 (Initial) through 5 (Optimized)"
    dv_scale.promptTitle = "Maturity Rating"
    ws.add_data_validation(dv_scale)

    # ── Build tier / criterion lookups ────────────────────────────────
    tier_names = {}
    tier_labels = {}
    tier_descs = {}
    for tier in rubric["tiers"]:
        tid = tier["id"]
        tier_names[tid] = tier["name"]
        desc = tier.get("description", "").strip()
        # Truncate long descriptions to first sentence
        if desc and ". " in desc:
            desc = desc[: desc.index(". ") + 1]
        tier_descs[tid] = desc
        if tid.startswith("tier_"):
            num = tid.replace("tier_", "")
            tier_names[int(num)] = tier["name"]
            tier_labels[int(num)] = f"TIER {num}: {tier['name'].upper()}"
            tier_labels[tid] = tier_labels[int(num)]
        else:
            tier_labels[tid] = tier["name"].upper().replace("ENRICHMENT: ", "")

    crit_names = {}
    for tier in rubric["tiers"]:
        for crit in tier["criteria"]:
            crit_names[crit["id"]] = crit["name"]

    # ── DEBMM Core Assessment context header ──────────────────────────
    row = 10
    ws.merge_cells(f"B{row}:F{row}")
    ws.cell(row=row, column=2, value="TIBMM CORE ASSESSMENT \u2014 Tiers 0\u20144")
    style_cell(ws.cell(row=row, column=2), FONT_SECTION, FILL_LIGHT_BLUE, ALIGN_LEFT, BLUE_ACCENT_LEFT)
    for c in range(3, 7):
        ws.cell(row=row, column=c).fill = FILL_LIGHT_BLUE
    ws.row_dimensions[row].height = 32
    row += 1

    ws.merge_cells(f"B{row}:F{row}")
    ws.cell(row=row, column=2,
            value="Rate your team across the 5 progressive TIBMM tiers. "
                  "Your achieved tier is the highest where all criteria score \u2265 3.0.")
    style_cell(ws.cell(row=row, column=2), FONT_CONTEXT, FILL_LIGHT_BLUE, ALIGN_LEFT)
    for c in range(3, 7):
        ws.cell(row=row, column=c).fill = FILL_LIGHT_BLUE
    ws.row_dimensions[row].height = 22
    row += 1

    # ── Column headers ────────────────────────────────────────────────
    if mode == "audit":
        headers = ["ID", "Criterion", "Question", "Answer", "Score", "Evidence / Notes"]
        hdr_cols = list(range(2, 8))
    else:
        headers = ["ID", "Criterion", "Question", "Answer", "Score"]
        hdr_cols = list(range(2, 7))

    header_row = row
    for col_idx, header in zip(hdr_cols, headers):
        c = ws.cell(row=header_row, column=col_idx, value=header)
        style_cell(c, FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    ws.row_dimensions[header_row].height = 26
    row += 1

    # ── Question rows ─────────────────────────────────────────────────
    current_tier = None
    question_rows = []
    q_key = "question" if mode == "self" else "question_audit"
    enrichment_headers_inserted = False
    prev_criterion = None
    q_idx_in_tier = 0  # For alternating rows

    for q in questionnaire["questions"]:
        q_tier = q["tier"]

        # ── Enrichment section transition ─────────────────────────────
        if is_enrichment(q_tier) and not enrichment_headers_inserted:
            enrichment_headers_inserted = True
            prev_criterion = None
            q_idx_in_tier = 0

            # Spacer / chapter break
            ws.row_dimensions[row].height = 20
            for c in range(1, last_col_num + 1):
                ws.cell(row=row, column=c).fill = FILL_LIGHT_GRAY
            row += 1

            # Enrichment context header
            ws.merge_cells(f"B{row}:F{row}")
            ws.cell(row=row, column=2,
                    value="ENRICHMENT DIMENSIONS \u2014 Tradecraft and Governance")
            style_cell(ws.cell(row=row, column=2), FONT_SECTION_TEAL, FILL_LIGHT_TEAL,
                       ALIGN_LEFT, TEAL_ACCENT_LEFT)
            for c in range(3, 7):
                ws.cell(row=row, column=c).fill = FILL_LIGHT_TEAL
            ws.row_dimensions[row].height = 32
            row += 1

            ws.merge_cells(f"B{row}:F{row}")
            ws.cell(row=row, column=2,
                    value="Cross-cutting tradecraft and governance factors. They contribute to the "
                          "overall score but do not affect TIBMM tier determination.")
            style_cell(ws.cell(row=row, column=2), FONT_CONTEXT, FILL_LIGHT_TEAL, ALIGN_LEFT)
            for c in range(3, 7):
                ws.cell(row=row, column=c).fill = FILL_LIGHT_TEAL
            ws.row_dimensions[row].height = 22
            row += 1

            # Repeated column headers with teal
            for col_idx, hdr in zip(hdr_cols, headers):
                c = ws.cell(row=row, column=col_idx, value=hdr)
                style_cell(c, FONT_COL_HEADER, FILL_TEAL, ALIGN_CENTER, THIN_BORDER)
            ws.row_dimensions[row].height = 26
            row += 1

        # ── Tier separator banner ─────────────────────────────────────
        if q_tier != current_tier:
            current_tier = q_tier
            prev_criterion = None
            q_idx_in_tier = 0

            banner_fill = FILL_TEAL if is_enrichment(q_tier) else FILL_NAVY
            banner_text = tier_labels.get(q_tier, str(q_tier).upper())

            ws.merge_cells(f"B{row}:F{row}")
            ws.cell(row=row, column=2, value=f"  {banner_text}")
            style_cell(ws.cell(row=row, column=2), FONT_TIER_BANNER, banner_fill,
                       Alignment(horizontal="left", vertical="center"))
            for c in range(2, 7):
                ws.cell(row=row, column=c).fill = banner_fill
            if evidence_col:
                ws.cell(row=row, column=evidence_col).fill = banner_fill
            ws.row_dimensions[row].height = 34
            row += 1

            # Tier description sub-row
            desc = tier_descs.get(q_tier if isinstance(q_tier, str) else f"tier_{q_tier}", "")
            if not desc:
                desc = tier_descs.get(q_tier, "")
            if desc:
                sub_fill = FILL_LIGHT_TEAL if is_enrichment(q_tier) else FILL_LIGHT_BLUE
                ws.merge_cells(f"B{row}:F{row}")
                ws.cell(row=row, column=2, value=desc)
                style_cell(ws.cell(row=row, column=2), FONT_SMALL_ITALIC, sub_fill, ALIGN_LEFT)
                for c in range(3, 7):
                    ws.cell(row=row, column=c).fill = sub_fill
                ws.row_dimensions[row].height = 20
                row += 1

        # ── Question data ─────────────────────────────────────────────
        qid = q["id"]
        qtype = q["type"]
        criterion_id = q["criterion"]
        criterion = crit_names.get(criterion_id, criterion_id)
        question_text = q.get(q_key, q["question"])
        yes_value = q.get("scoring", {}).get("yes_value", 3) if qtype == "checklist" else None

        # Alternating row tint
        enrichment_q = is_enrichment(q_tier)
        if q_idx_in_tier % 2 == 1:
            row_fill = FILL_LIGHT_TEAL if enrichment_q else FILL_OFF_WHITE
        else:
            row_fill = FILL_WHITE

        # Criterion grouping — bold for first question of each criterion
        is_first_of_criterion = (criterion_id != prev_criterion)
        prev_criterion = criterion_id
        crit_font = FONT_BODY_BOLD if is_first_of_criterion else FONT_SMALL

        # Column B — ID
        ws.cell(row=row, column=2, value=qid)
        style_cell(ws.cell(row=row, column=2), FONT_SMALL, row_fill, ALIGN_CENTER, HAIRLINE_BOTTOM)

        # Column C — Criterion
        ws.cell(row=row, column=3, value=criterion)
        style_cell(ws.cell(row=row, column=3), crit_font, row_fill, ALIGN_LEFT_TOP, HAIRLINE_BOTTOM)

        # Column D — Question (with scale options)
        if qtype == "scale" and "options" in q:
            option_lines = "\n".join(f"{k} \u2014 {v}" for k, v in q["options"].items())
            full_question = f"{question_text}\n\n{option_lines}"
        else:
            full_question = question_text
        ws.cell(row=row, column=4, value=full_question)
        style_cell(ws.cell(row=row, column=4), FONT_BODY, row_fill, ALIGN_LEFT_TOP, HAIRLINE_BOTTOM)

        # Column E — Answer (warm amber)
        answer_cell = ws.cell(row=row, column=5)
        style_cell(answer_cell, FONT_ANSWER, FILL_ANSWER, ALIGN_CENTER, ANSWER_BORDER)
        if qtype == "checklist":
            dv_yesno.add(answer_cell)
        elif qtype == "scale":
            dv_scale.add(answer_cell)

        # Column F — Score (auto-calculated)
        score_cell = ws.cell(row=row, column=6)
        answer_ref = f"E{row}"
        if qtype == "checklist":
            score_cell.value = f'=IF({answer_ref}="Yes",{yes_value},IF({answer_ref}="No",1,""))'
        elif qtype == "scale":
            score_cell.value = f'=IF({answer_ref}="","",{answer_ref})'
        style_cell(score_cell, FONT_BODY_BOLD, row_fill, ALIGN_CENTER, HAIRLINE_BOTTOM)
        score_cell.number_format = "0.0"

        # Column G — Evidence (audit only)
        if evidence_col:
            evidence_cell = ws.cell(row=row, column=evidence_col)
            style_cell(evidence_cell, FONT_BODY, FILL_ANSWER, ALIGN_LEFT_TOP, ANSWER_BORDER)

        # Row height
        ws.row_dimensions[row].height = 115 if qtype == "scale" else 35

        question_rows.append({
            "row": row,
            "id": qid,
            "type": qtype,
            "criterion": criterion_id,
            "tier": q_tier,
            "yes_value": yes_value,
        })
        q_idx_in_tier += 1
        row += 1

    # Conditional formatting on score column (F)
    apply_conditional_formatting(ws, f"F{header_row + 1}:F{row - 1}")

    # Freeze panes below first column header row
    ws.freeze_panes = f"A{header_row + 1}"

    return ws, question_rows, header_row


# ── Tab 3: Results Dashboard ──────────────────────────────────────────────────


def build_dashboard_tab(wb: Workbook, rubric: dict, questionnaire: dict,
                        question_rows: list, header_row: int):
    ws = wb.create_sheet("Results Dashboard")
    ws.sheet_properties.tabColor = NAVY

    col_widths = {"A": 3, "B": 6, "C": 34, "D": 14, "E": 14, "F": 14, "G": 3}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width
    # ── Title banner ──────────────────────────────────────────────────
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 56
    c = ws["A1"]
    c.value = "  TIBMM Assessment Results"
    style_cell(c, FONT_TITLE, FILL_DARK_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCDEFG":
        ws[f"{col}1"].fill = FILL_DARK_NAVY

    # Subtitle — org name
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 24
    c = ws["A2"]
    c.value = '="  "&Assessment!D4'
    style_cell(c, FONT_SUBTITLE, FILL_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCDEFG":
        ws[f"{col}2"].fill = FILL_NAVY

    # ── Executive Summary Cards ───────────────────────────────────────
    row = 5

    # Labels row
    ws.merge_cells(f"B{row}:C{row}")
    ws.cell(row=row, column=2, value="Overall Maturity Score")
    style_cell(ws.cell(row=row, column=2), FONT_SCORE_LABEL, FILL_LIGHT_BLUE, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=3).fill = FILL_LIGHT_BLUE
    ws.cell(row=row, column=3).border = THIN_BORDER

    ws.merge_cells(f"D{row}:E{row}")
    ws.cell(row=row, column=4, value="Achieved TIBMM Tier")
    style_cell(ws.cell(row=row, column=4), FONT_SCORE_LABEL, FILL_LIGHT_BLUE, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=5).fill = FILL_LIGHT_BLUE
    ws.cell(row=row, column=5).border = THIN_BORDER

    ws.cell(row=row, column=6, value="Completion")
    style_cell(ws.cell(row=row, column=6), FONT_SCORE_LABEL, FILL_LIGHT_BLUE, ALIGN_CENTER, THIN_BORDER)
    ws.row_dimensions[row].height = 22
    row += 1

    # Values row
    ws.row_dimensions[row].height = 60
    ws.merge_cells(f"B{row}:C{row}")
    style_cell(ws.cell(row=row, column=2), FONT_SCORE_HERO, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=3).border = THIN_BORDER

    ws.merge_cells(f"D{row}:E{row}")
    style_cell(ws.cell(row=row, column=4), FONT_SCORE_LARGE, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=5).border = THIN_BORDER

    total_q = len(question_rows)
    # Reference only actual question cells to avoid counting header rows
    q_cells = ",".join(f"Assessment!E{qr['row']}" for qr in question_rows)
    completion_formula = f'=COUNTA({q_cells})&" / {total_q}"'
    ws.cell(row=row, column=6, value=completion_formula)
    style_cell(ws.cell(row=row, column=6), FONT_SCORE_LARGE, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)

    overall_row = row
    row += 1

    # Spacer
    ws.row_dimensions[row].height = 8
    row += 1

    # Tier explanation — dynamic text based on achieved tier
    tier_explanations = {
        "Below Foundation": (
            "The CTI function has not yet achieved Tier 0 (Foundation). One or more foundational "
            "criteria \u2014 charter/service catalog, IOC handling, source inventory, or stakeholder "
            "touchpoints \u2014 score below the Defined level (3.0). Focus on documenting what CTI "
            "delivers and to whom, and on establishing a minimally disciplined IOC process, before "
            "pursuing higher tiers."
        ),
        "Tier 0: Foundation": (
            "The CTI function has achieved Tier 0 (Foundation). The role is documented, IOCs are "
            "handled consistently, sources are inventoried, and stakeholders know how to engage. "
            "To reach Tier 1, publish Priority Intelligence Requirements, adopt a BLUF product "
            "template, require calibrated probability language and source reliability ratings, "
            "and build a Collection Management Framework."
        ),
        "Tier 1: Basic": (
            "The CTI function has achieved Tier 1 (Basic). PIRs and the tradecraft floor "
            "(BLUF, calibrated language, source rating, CMF) are in place. To advance to Tier 2, "
            "deploy a Threat Intelligence Platform with automated IOC flow to the SIEM, "
            "formalize the CTI\u2192Detection-Engineering handoff, establish a working cadence "
            "with the Threat Hunting team, and enrich every major incident with CTI context."
        ),
        "Tier 2: Operational": (
            "The CTI function has achieved Tier 2 (Operational). TIP + IOC pipeline, DE handoff, "
            "CTI-Hunt partnership, and IR enrichment are running on a cadence. To advance to Tier 3, "
            "stand up recurring products for Vulnerability Management (EPSS + telco-stack tagging), "
            "Fraud (smishing, SIM swap, brand), and Third-Party Risk, and publish a tiered Cyber "
            "Threat Landscape report."
        ),
        "Tier 3: Advanced": (
            "The CTI function has achieved Tier 3 (Advanced). Stakeholder-specific intelligence "
            "products (VM, Fraud, TPRM, CTL) are running on cadence and influencing downstream "
            "decisions. To advance to Tier 4, produce a standing strategic brief for the CISO, "
            "apply Structured Analytic Techniques to major products, track CTI-attributable "
            "outcomes quantitatively, and measure forecast calibration."
        ),
        "Tier 4: Leading": (
            "The CTI function has achieved Tier 4 (Leading) \u2014 the highest TIBMM tier. "
            "Strategic intelligence reaches the CISO and board, SATs are standard practice, "
            "attributable outcomes are reported to leadership, and forecast calibration is "
            "measured. Focus on sustaining this posture, contributing to the telco-sector "
            "intelligence community, and using outcome data at budget review."
        ),
    }

    # Build nested IF formula for dynamic explanation
    tier_labels_ordered = [
        "Tier 4: Leading", "Tier 3: Advanced", "Tier 2: Operational",
        "Tier 1: Basic", "Tier 0: Foundation",
    ]
    tier_cell = f"D{overall_row}"
    explanation_formula = f'=IF({tier_cell}="","",'
    for label in tier_labels_ordered:
        text = tier_explanations[label]
        explanation_formula += f'IF({tier_cell}="{label}","{text}",'
    explanation_formula += f'"{tier_explanations["Below Foundation"]}"'
    explanation_formula += ")" * (len(tier_labels_ordered) + 1)

    ws.merge_cells(f"B{row}:F{row}")
    ws.cell(row=row, column=2, value=explanation_formula)
    style_cell(ws.cell(row=row, column=2), FONT_BODY_ITALIC, FILL_LIGHT_GRAY, ALIGN_WRAP)
    for c in range(3, 7):
        ws.cell(row=row, column=c).fill = FILL_LIGHT_GRAY
    ws.row_dimensions[row].height = 62
    row += 2

    # ── Build mappings ────────────────────────────────────────────────
    crit_to_rows = {}
    for qr in question_rows:
        crit_to_rows.setdefault(qr["criterion"], []).append(qr)

    tier_criteria_list = []
    for tier in rubric["tiers"]:
        for crit in tier["criteria"]:
            tier_criteria_list.append({
                "tier_id": tier["id"],
                "tier_name": tier["name"],
                "crit_id": crit["id"],
                "crit_name": crit["name"],
            })

    core_tier_ids = {"tier_0", "tier_1", "tier_2", "tier_3", "tier_4"}
    core_criteria = [tc for tc in tier_criteria_list if tc["tier_id"] in core_tier_ids]
    enrichment_criteria = [tc for tc in tier_criteria_list if tc["tier_id"] not in core_tier_ids]

    # ── DEBMM Core Section ────────────────────────────────────────────
    ws.merge_cells(f"B{row}:F{row}")
    ws.cell(row=row, column=2, value="  TIBMM CORE ASSESSMENT")
    style_cell(ws.cell(row=row, column=2), FONT_TIER_BANNER, FILL_NAVY,
               Alignment(horizontal="left", vertical="center"))
    for c in range(3, 7):
        ws.cell(row=row, column=c).fill = FILL_NAVY
    ws.row_dimensions[row].height = 30
    row += 1

    # Column headers
    core_hdr_row = row
    for col_idx, hdr in [(3, "Category / Criterion"), (4, "Score"), (5, "Level"), (6, "Status")]:
        ws.cell(row=row, column=col_idx, value=hdr)
        style_cell(ws.cell(row=row, column=col_idx), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    ws.row_dimensions[row].height = 24
    row += 1

    criterion_score_cells = []
    core_criterion_cells = []
    tier_score_cells = {}
    current_tier_id = None
    tier_crit_cells = []

    def _finalize_tier(tid):
        nonlocal tier_crit_cells
        if tid and tier_crit_cells and tid in tier_score_cells:
            trow = tier_score_cells[tid]["row"]
            avg_refs = ",".join(tier_crit_cells)
            ws.cell(row=trow, column=4, value=f'=IF(COUNT({avg_refs})=0,"",AVERAGE({avg_refs}))')
            ws.cell(row=trow, column=4).number_format = "0.00"
            ws.cell(row=trow, column=5, value=level_formula(f"D{trow}"))
            ws.cell(row=trow, column=6, value=status_formula(f"D{trow}"))
        tier_crit_cells = []

    for tc in core_criteria:
        if tc["tier_id"] != current_tier_id:
            _finalize_tier(current_tier_id)
            current_tier_id = tc["tier_id"]

            # Tier summary row
            ws.cell(row=row, column=3, value=tc["tier_name"])
            style_cell(ws.cell(row=row, column=3), FONT_BODY_BOLD, FILL_LIGHT_BLUE, ALIGN_LEFT, THIN_BORDER)
            for c in range(4, 7):
                style_cell(ws.cell(row=row, column=c), FONT_BODY_BOLD, FILL_LIGHT_BLUE, ALIGN_CENTER, THIN_BORDER)
            tier_score_cells[current_tier_id] = {"row": row, "name": tc["tier_name"]}
            ws.row_dimensions[row].height = 28
            row += 1

        # Criterion row
        crit_id = tc["crit_id"]
        rows_for_crit = crit_to_rows.get(crit_id, [])

        ws.cell(row=row, column=3, value=f"    {tc['crit_name']}")
        style_cell(ws.cell(row=row, column=3), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)

        if rows_for_crit:
            score_refs = [f"Assessment!F{qr['row']}" for qr in rows_for_crit]
            refs_str = ",".join(score_refs)
            ws.cell(row=row, column=4, value=f'=IF(COUNT({refs_str})=0,"",AVERAGE({refs_str}))')
        ws.cell(row=row, column=4).number_format = "0.00"
        style_cell(ws.cell(row=row, column=4), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=5, value=level_formula(f"D{row}"))
        style_cell(ws.cell(row=row, column=5), FONT_BODY, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=6, value=status_formula(f"D{row}"))
        style_cell(ws.cell(row=row, column=6), FONT_BODY, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        criterion_score_cells.append(f"D{row}")
        core_criterion_cells.append(f"D{row}")
        tier_crit_cells.append(f"D{row}")
        row += 1

    _finalize_tier(current_tier_id)

    # Conditional formatting on core section
    core_end_row = row - 1
    apply_conditional_formatting(ws, f"D{core_hdr_row + 1}:D{core_end_row}")
    ws.conditional_formatting.add(
        f"F{core_hdr_row + 1}:F{core_end_row}",
        CellIsRule(operator="equal", formula=['"✓ Pass"'], fill=FILL_SCORE_GREEN),
    )
    ws.conditional_formatting.add(
        f"F{core_hdr_row + 1}:F{core_end_row}",
        CellIsRule(operator="equal", formula=['"✗ Below Target"'], fill=FILL_SCORE_RED),
    )

    # ── Spacer ────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 14
    for c in range(1, 8):
        ws.cell(row=row, column=c).fill = FILL_LIGHT_GRAY
    row += 1

    # ── Enrichment Section ────────────────────────────────────────────
    ws.merge_cells(f"B{row}:F{row}")
    ws.cell(row=row, column=2, value="  ENRICHMENT DIMENSIONS")
    style_cell(ws.cell(row=row, column=2), FONT_TIER_BANNER, FILL_TEAL,
               Alignment(horizontal="left", vertical="center"))
    for c in range(3, 7):
        ws.cell(row=row, column=c).fill = FILL_TEAL
    ws.row_dimensions[row].height = 30
    row += 1

    ws.merge_cells(f"B{row}:F{row}")
    ws.cell(row=row, column=2,
            value="Tradecraft and governance factors \u2014 do not affect TIBMM tier determination")
    style_cell(ws.cell(row=row, column=2), FONT_SMALL_ITALIC, FILL_LIGHT_TEAL, ALIGN_LEFT)
    for c in range(3, 7):
        ws.cell(row=row, column=c).fill = FILL_LIGHT_TEAL
    ws.row_dimensions[row].height = 20
    row += 1

    # Column headers (teal)
    enrich_hdr_row = row
    for col_idx, hdr in [(3, "Category / Criterion"), (4, "Score"), (5, "Level"), (6, "Status")]:
        ws.cell(row=row, column=col_idx, value=hdr)
        style_cell(ws.cell(row=row, column=col_idx), FONT_COL_HEADER, FILL_TEAL, ALIGN_CENTER, THIN_BORDER)
    ws.row_dimensions[row].height = 24
    row += 1

    enrich_tier_score_cells = {}
    current_tier_id = None
    tier_crit_cells = []

    for tc in enrichment_criteria:
        if tc["tier_id"] != current_tier_id:
            _finalize_tier(current_tier_id)
            current_tier_id = tc["tier_id"]

            # Category summary row
            display_name = tc["tier_name"].replace("Enrichment: ", "")
            ws.cell(row=row, column=3, value=display_name)
            style_cell(ws.cell(row=row, column=3), FONT_BODY_BOLD, FILL_LIGHT_TEAL, ALIGN_LEFT, THIN_BORDER)
            for c in range(4, 7):
                style_cell(ws.cell(row=row, column=c), FONT_BODY_BOLD, FILL_LIGHT_TEAL, ALIGN_CENTER, THIN_BORDER)
            tier_score_cells[current_tier_id] = {"row": row, "name": display_name}
            enrich_tier_score_cells[current_tier_id] = {"row": row, "name": display_name}
            ws.row_dimensions[row].height = 28
            row += 1

        # Criterion row
        crit_id = tc["crit_id"]
        rows_for_crit = crit_to_rows.get(crit_id, [])

        ws.cell(row=row, column=3, value=f"    {tc['crit_name']}")
        style_cell(ws.cell(row=row, column=3), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)

        if rows_for_crit:
            score_refs = [f"Assessment!F{qr['row']}" for qr in rows_for_crit]
            refs_str = ",".join(score_refs)
            ws.cell(row=row, column=4, value=f'=IF(COUNT({refs_str})=0,"",AVERAGE({refs_str}))')
        ws.cell(row=row, column=4).number_format = "0.00"
        style_cell(ws.cell(row=row, column=4), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=5, value=level_formula(f"D{row}"))
        style_cell(ws.cell(row=row, column=5), FONT_BODY, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=6, value=status_formula(f"D{row}"))
        style_cell(ws.cell(row=row, column=6), FONT_BODY, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        criterion_score_cells.append(f"D{row}")
        tier_crit_cells.append(f"D{row}")
        row += 1

    _finalize_tier(current_tier_id)

    # Conditional formatting on enrichment section
    enrich_end_row = row - 1
    apply_conditional_formatting(ws, f"D{enrich_hdr_row + 1}:D{enrich_end_row}")
    ws.conditional_formatting.add(
        f"F{enrich_hdr_row + 1}:F{enrich_end_row}",
        CellIsRule(operator="equal", formula=['"✓ Pass"'], fill=FILL_SCORE_GREEN),
    )
    ws.conditional_formatting.add(
        f"F{enrich_hdr_row + 1}:F{enrich_end_row}",
        CellIsRule(operator="equal", formula=['"✗ Below Target"'], fill=FILL_SCORE_RED),
    )

    # ── Overall score formula ─────────────────────────────────────────
    if criterion_score_cells:
        all_refs = ",".join(criterion_score_cells)
        ws.cell(row=overall_row, column=2,
                value=f'=IF(COUNT({all_refs})=0,"",ROUND(AVERAGE({all_refs}),2)&" / 5.0")')
        ws.cell(row=overall_row, column=2).number_format = "@"

    # ── Achieved tier formula ─────────────────────────────────────────
    core_tiers_ordered = ["tier_0", "tier_1", "tier_2", "tier_3", "tier_4"]
    tier_display_labels = {
        "tier_0": "Tier 0: Foundation",
        "tier_1": "Tier 1: Basic",
        "tier_2": "Tier 2: Operational",
        "tier_3": "Tier 3: Advanced",
        "tier_4": "Tier 4: Leading",
    }

    tier_check_parts = {tid: [] for tid in core_tiers_ordered}
    crit_cell_idx = 0
    current_tier_scan = None
    for tc_item in core_criteria:
        if tc_item["tier_id"] != current_tier_scan:
            current_tier_scan = tc_item["tier_id"]
        if crit_cell_idx < len(core_criterion_cells):
            tier_check_parts[current_tier_scan].append(core_criterion_cells[crit_cell_idx])
        crit_cell_idx += 1

    def tier_check(tid):
        cells = tier_check_parts.get(tid, [])
        if not cells:
            return "TRUE"
        return f"AND({','.join(f'ISNUMBER({c}),{c}>=3' for c in cells)})"

    cumul = {}
    for i, tid in enumerate(core_tiers_ordered):
        checks = [tier_check(core_tiers_ordered[j]) for j in range(i + 1)]
        cumul[tid] = f"AND({','.join(checks)})"

    tier_formula = f'=IF({cumul["tier_4"]},"{tier_display_labels["tier_4"]}",'
    tier_formula += f'IF({cumul["tier_3"]},"{tier_display_labels["tier_3"]}",'
    tier_formula += f'IF({cumul["tier_2"]},"{tier_display_labels["tier_2"]}",'
    tier_formula += f'IF({cumul["tier_1"]},"{tier_display_labels["tier_1"]}",'
    tier_formula += f'IF({cumul["tier_0"]},"{tier_display_labels["tier_0"]}",'
    tier_formula += '"Below Foundation")))))'
    ws.cell(row=overall_row, column=4, value=tier_formula)

    # Conditional formatting on overall score
    apply_conditional_formatting(ws, f"B{overall_row}:C{overall_row}")

    ws.freeze_panes = "A3"

    # Collect chart data for separate chart tabs
    core_tier_row_list = [
        (ts["name"], ts["row"])
        for tid, ts in tier_score_cells.items()
        if tid in core_tier_ids
    ]
    enrich_row_list = [
        (ts["name"], ts["row"])
        for ts in enrich_tier_score_cells.values()
    ]
    return ws, core_tier_row_list, enrich_row_list


# ── Tab 4: DEBMM Tier Scores Chart ───────────────────────────────────────────


def build_core_chart_tab(wb: Workbook, core_tier_row_list: list):
    """Dedicated sheet for the TIBMM Core Tier Scores bar chart."""
    if not core_tier_row_list:
        return None
    ws = wb.create_sheet("Tier Scores Chart")
    ws.sheet_properties.tabColor = MED_BLUE

    col_widths = {"A": 3, "B": 22, "C": 12, "D": 3}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Title banner
    ws.merge_cells("A1:D1")
    ws.row_dimensions[1].height = 48
    c = ws["A1"]
    c.value = "  TIBMM Tier Scores"
    style_cell(c, FONT_TITLE, FILL_DARK_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCD":
        ws[f"{col}1"].fill = FILL_DARK_NAVY

    # Subtitle
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 22
    c = ws["A2"]
    c.value = "  Average score per TIBMM core tier (target: 3.0)"
    style_cell(c, FONT_SUBTITLE, FILL_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCD":
        ws[f"{col}2"].fill = FILL_NAVY

    # Data table header
    row = 4
    ws.cell(row=row, column=2, value="Tier")
    style_cell(ws.cell(row=row, column=2), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=3, value="Score")
    style_cell(ws.cell(row=row, column=3), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    ws.row_dimensions[row].height = 24
    row += 1

    data_start = row
    for name, trow in core_tier_row_list:
        ws.cell(row=row, column=2, value=name)
        style_cell(ws.cell(row=row, column=2), FONT_BODY_BOLD, FILL_LIGHT_BLUE, ALIGN_LEFT, THIN_BORDER)
        ws.cell(row=row, column=3, value=f"='Results Dashboard'!D{trow}")
        ws.cell(row=row, column=3).number_format = "0.00"
        style_cell(ws.cell(row=row, column=3), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)
        apply_conditional_formatting(ws, f"C{row}:C{row}")
        ws.row_dimensions[row].height = 22
        row += 1
    data_end = row - 1

    # Build chart
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "TIBMM Tier Scores"
    chart.y_axis.title = "Score (1\u20145)"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 5
    chart.x_axis.title = None
    chart.legend = None

    data_ref = Reference(ws, min_col=3, min_row=4, max_row=data_end)
    cats_ref = Reference(ws, min_col=2, min_row=data_start, max_row=data_end)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    if chart.series:
        for pt_idx in range(len(core_tier_row_list)):
            pt = DataPoint(idx=pt_idx)
            pt.graphicalProperties.solidFill = MED_BLUE
            chart.series[0].data_points.append(pt)

    chart.width = 22
    chart.height = 14
    row += 1
    ws.add_chart(chart, f"B{row}")

    ws.freeze_panes = "A3"
    return ws


# ── Tab 5: Organizational Readiness Chart ────────────────────────────────────


def build_enrichment_chart_tab(wb: Workbook, enrich_row_list: list):
    """Dedicated sheet for the Organizational Readiness bar chart."""
    if not enrich_row_list:
        return None
    ws = wb.create_sheet("Readiness Chart")
    ws.sheet_properties.tabColor = TEAL

    col_widths = {"A": 3, "B": 30, "C": 12, "D": 3}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Title banner
    ws.merge_cells("A1:D1")
    ws.row_dimensions[1].height = 48
    c = ws["A1"]
    c.value = "  Tradecraft & Governance Readiness"
    style_cell(c, FONT_TITLE, FILL_DARK_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCD":
        ws[f"{col}1"].fill = FILL_DARK_NAVY

    # Subtitle
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 22
    c = ws["A2"]
    c.value = "  Enrichment dimensions \u2014 do not affect TIBMM tier determination"
    style_cell(c, FONT_SUBTITLE, FILL_TEAL, Alignment(horizontal="left", vertical="center"))
    for col in "ABCD":
        ws[f"{col}2"].fill = FILL_TEAL

    # Data table header
    row = 4
    ws.cell(row=row, column=2, value="Dimension")
    style_cell(ws.cell(row=row, column=2), FONT_COL_HEADER, FILL_TEAL, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=3, value="Score")
    style_cell(ws.cell(row=row, column=3), FONT_COL_HEADER, FILL_TEAL, ALIGN_CENTER, THIN_BORDER)
    ws.row_dimensions[row].height = 24
    row += 1

    data_start = row
    for name, trow in enrich_row_list:
        ws.cell(row=row, column=2, value=name)
        style_cell(ws.cell(row=row, column=2), FONT_BODY_BOLD, FILL_LIGHT_TEAL, ALIGN_LEFT, THIN_BORDER)
        ws.cell(row=row, column=3, value=f"='Results Dashboard'!D{trow}")
        ws.cell(row=row, column=3).number_format = "0.00"
        style_cell(ws.cell(row=row, column=3), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)
        apply_conditional_formatting(ws, f"C{row}:C{row}")
        ws.row_dimensions[row].height = 22
        row += 1
    data_end = row - 1

    # Build chart
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Tradecraft & Governance Readiness"
    chart.y_axis.title = "Score (1\u20145)"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 5
    chart.x_axis.title = None
    chart.legend = None

    data_ref = Reference(ws, min_col=3, min_row=4, max_row=data_end)
    cats_ref = Reference(ws, min_col=2, min_row=data_start, max_row=data_end)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)

    if chart.series:
        for pt_idx in range(len(enrich_row_list)):
            pt = DataPoint(idx=pt_idx)
            pt.graphicalProperties.solidFill = TEAL
            chart.series[0].data_points.append(pt)

    chart.width = 18
    chart.height = 14
    row += 1
    ws.add_chart(chart, f"B{row}")

    ws.freeze_panes = "A3"
    return ws


# ── Tab 6: Rubric Reference ─────────────────────────────────────────────────


def build_rubric_tab(wb: Workbook, rubric: dict):
    ws = wb.create_sheet("Rubric Reference")
    ws.sheet_properties.tabColor = STEEL

    col_widths = {"A": 14, "B": 60, "C": 40}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Title
    ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 48
    c = ws["A1"]
    c.value = "  TIBMM Rubric Reference"
    style_cell(c, FONT_TITLE, FILL_DARK_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABC":
        ws[f"{col}1"].fill = FILL_DARK_NAVY

    row = 3
    level_names = {1: "Initial", 2: "Repeatable", 3: "Defined", 4: "Managed", 5: "Optimized"}

    core_tiers = [t for t in rubric["tiers"] if not is_enrichment(t["id"])]
    enrichment_tiers = [t for t in rubric["tiers"] if is_enrichment(t["id"])]

    def render_tiers(tiers, section_label, banner_fill, section_fill, accent_border, section_font):
        nonlocal row

        # Section header
        ws.merge_cells(f"A{row}:C{row}")
        ws.cell(row=row, column=1, value=f"  {section_label}")
        style_cell(ws.cell(row=row, column=1), FONT_TIER_BANNER, banner_fill,
                   Alignment(horizontal="left", vertical="center"))
        for c in range(2, 4):
            ws.cell(row=row, column=c).fill = banner_fill
        ws.row_dimensions[row].height = 28
        row += 1

        for tier in tiers:
            # Tier banner
            ws.merge_cells(f"A{row}:C{row}")
            tier_name = tier["name"]
            if tier["id"].startswith("tier_"):
                num = tier["id"].replace("tier_", "")
                tier_name = f"Tier {num}: {tier['name']}"
            else:
                tier_name = tier["name"].replace("Enrichment: ", "")
            ws.cell(row=row, column=1, value=f"  {tier_name}")
            style_cell(ws.cell(row=row, column=1), FONT_TIER_BANNER, banner_fill,
                       Alignment(horizontal="left", vertical="center"))
            for c in range(2, 4):
                ws.cell(row=row, column=c).fill = banner_fill
            ws.row_dimensions[row].height = 30
            row += 1

            # Tier description
            desc = tier.get("description", "").strip()
            if desc:
                ws.merge_cells(f"A{row}:C{row}")
                ws.cell(row=row, column=1, value=desc)
                style_cell(ws.cell(row=row, column=1), FONT_SMALL_ITALIC, section_fill, ALIGN_WRAP)
                for c in range(2, 4):
                    ws.cell(row=row, column=c).fill = section_fill
                ws.row_dimensions[row].height = 36
                row += 1

            for crit in tier["criteria"]:
                # Criterion name
                ws.merge_cells(f"A{row}:C{row}")
                ws.cell(row=row, column=1, value=crit["name"])
                style_cell(ws.cell(row=row, column=1), section_font, section_fill, ALIGN_LEFT, accent_border)
                for c in range(2, 4):
                    ws.cell(row=row, column=c).fill = section_fill
                ws.row_dimensions[row].height = 26
                row += 1

                # Sub-headers
                for col_idx, hdr in enumerate(["Level", "Description", "Quantitative Measure"], 1):
                    ws.cell(row=row, column=col_idx, value=hdr)
                    style_cell(ws.cell(row=row, column=col_idx), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
                ws.row_dimensions[row].height = 22
                row += 1

                for level_num in sorted(crit["levels"].keys()):
                    level_data = crit["levels"][level_num]
                    row_fill = FILL_LEVEL.get(level_num, FILL_WHITE)

                    ws.cell(row=row, column=1, value=f"{level_num} \u2014 {level_names[level_num]}")
                    style_cell(ws.cell(row=row, column=1), FONT_LEVEL_BOLD, row_fill, ALIGN_CENTER, THIN_BORDER)

                    ws.cell(row=row, column=2, value=level_data["qualitative"].strip())
                    style_cell(ws.cell(row=row, column=2), FONT_BODY, row_fill, ALIGN_WRAP, THIN_BORDER)

                    ws.cell(row=row, column=3, value=level_data.get("quantitative", "").strip())
                    style_cell(ws.cell(row=row, column=3), FONT_BODY, row_fill, ALIGN_WRAP, THIN_BORDER)

                    ws.row_dimensions[row].height = 42
                    row += 1

                row += 1  # Space between criteria

    # Render core tiers
    render_tiers(core_tiers, "TIBMM CORE TIERS", FILL_NAVY, FILL_LIGHT_BLUE,
                 BLUE_ACCENT_LEFT, FONT_SECTION)

    # Spacer
    ws.row_dimensions[row].height = 16
    row += 1

    # Render enrichment tiers
    render_tiers(enrichment_tiers, "ENRICHMENT DIMENSIONS", FILL_TEAL, FILL_LIGHT_TEAL,
                 TEAL_ACCENT_LEFT, FONT_SECTION_TEAL)

    ws.freeze_panes = "A2"
    return ws


# ── Tab 7: Glossary ──────────────────────────────────────────────────────────

GLOSSARY = [
    # (Term, Category, Definition)
    ("PIR",
     "Direction",
     "Priority Intelligence Requirement. A documented question that a stakeholder "
     "needs CTI to answer on a recurring basis. PIRs drive what CTI collects, "
     "analyzes, and produces. A mature CTI function has 6-12 active PIRs each "
     "tied to a named stakeholder and a business question."),

    ("SIR",
     "Direction",
     "Specific Intelligence Requirement. A narrower decomposition of a PIR into "
     "one or more concrete, collectable questions. Example: PIR \"What ransomware "
     "groups target Canadian telecom?\" decomposes into SIRs like \"Which ransomware "
     "groups have publicly claimed Canadian telecom victims in the last 12 months?\""),

    ("RFI",
     "Direction",
     "Request for Information. A stakeholder-initiated ad-hoc question outside "
     "the standing PIR set (e.g., \"what do we know about actor X?\")."),

    ("BLUF",
     "Product Structure",
     "Bottom Line Up Front. The intelligence-writing convention that places the "
     "key analytic judgment and recommendation in the first paragraph, before any "
     "background. Codified in the SANS Intelligence Analyst's Playbook. Distinct "
     "from an executive summary: BLUF is the judgment itself, not a preview."),

    ("ICD 203",
     "Tradecraft",
     "Intelligence Community Directive 203 (ODNI, Jan 2015, amended 2022). "
     "Defines nine analytic tradecraft standards governing US IC finished "
     "intelligence: (1) describe source quality, (2) express uncertainty with "
     "calibrated language, (3) distinguish information from assumption, (4) "
     "incorporate analysis of alternatives, (5) demonstrate customer relevance, "
     "(6) use clear argumentation, (7) explain change or consistency, (8) make "
     "accurate judgments, (9) use effective visuals. TIBMM adopts these as the "
     "tradecraft floor for product-producing activity."),

    ("Calibrated Probability Language",
     "Tradecraft",
     "A fixed probability scale used to express confidence. The ODNI 8-level scale: "
     "almost certain (93-99%), very likely (85-92%), likely (70-84%), probable "
     "(55-69%), roughly even (45-54%), unlikely (15-44%), very unlikely (3-14%), "
     "remote (1-2%). The point is to prevent \"likely\" meaning 55% to one reader "
     "and 85% to another. Never use 100% or 0%."),

    ("Admiralty Code / NATO Source Rating",
     "Tradecraft",
     "A two-dimensional source and information evaluation system. Source "
     "reliability is rated A (completely reliable) through F (cannot be judged); "
     "information credibility is rated 1 (confirmed) through 6 (cannot be judged). "
     "Source and information are rated independently; a piece of intelligence might "
     "be \"B2\" (usually reliable source, probably true)."),

    ("CRAAP",
     "Tradecraft",
     "Currency, Relevance, Authority, Accuracy, Purpose. A lightweight framework "
     "for evaluating a source of intelligence information. Currency: how recent? "
     "Relevance: does it address the question? Authority: who produced it? "
     "Accuracy: is it verifiable? Purpose: why was it published?"),

    ("SAT",
     "Tradecraft",
     "Structured Analytic Technique. Formal methods for combating cognitive bias "
     "in analysis. Common SATs: Analysis of Competing Hypotheses (ACH), Key "
     "Assumptions Check, Pre-Mortem, Devil's Advocacy, Red Team Analysis. TIBMM "
     "expects at least one SAT applied to every major product at Tier 4."),

    ("ACH",
     "Tradecraft",
     "Analysis of Competing Hypotheses. Richards Heuer's 8-step SAT that evaluates "
     "multiple explanations against the same evidence simultaneously, looking for "
     "the hypothesis least inconsistent with the evidence (rather than the one most "
     "consistent). Designed to combat confirmation bias."),

    ("CMF",
     "Collection",
     "Collection Management Framework. The document that maps each intelligence "
     "source (OSINT feed, commercial feed, ISAC, vendor portal, internal telemetry) "
     "to the PIRs and SIRs it is intended to support. The CMF makes collection "
     "gaps visible and justifies source spend. At a minimum, a CMF is a spreadsheet "
     "with columns: Source, Type, Owner, Cost, PIRs Supported, Coverage Notes."),

    ("TLP",
     "Sharing",
     "Traffic Light Protocol. The standard marking convention for sensitivity of "
     "shared intelligence: RED (named recipients only), AMBER (limited distribution "
     "within organization), AMBER+STRICT (org only, not for vendors/partners), "
     "GREEN (community/sector), CLEAR (public, formerly WHITE)."),

    ("STIX / TAXII",
     "Tooling",
     "STIX (Structured Threat Information Expression) is the standard data format "
     "for sharing CTI. TAXII (Trusted Automated Exchange of Indicator Information) "
     "is the transport protocol for pulling/pushing STIX. Most TIPs, SIEMs, and "
     "ISAC feeds support STIX/TAXII natively."),

    ("TIP",
     "Tooling",
     "Threat Intelligence Platform. The category of tool that aggregates, enriches, "
     "stores, scores, and distributes threat intelligence across feeds and "
     "internal systems. Examples: OpenCTI (open-source), MISP (open-source), "
     "ThreatConnect, Anomali, Recorded Future. A TIP acts as the single source of "
     "truth for indicators across SIEM/EDR/SOAR."),

    ("SOAR",
     "Tooling",
     "Security Orchestration, Automation, and Response. Playbook-driven automation "
     "that ties alerts, enrichment lookups, and response actions together. CTI "
     "integrates with SOAR so high-confidence indicators can trigger automated "
     "blocking and SOC alerts can auto-enrich with TIP context."),

    ("EPSS",
     "Vulnerability",
     "Exploit Prediction Scoring System. A probabilistic scoring system from FIRST "
     "(0-1 scale) indicating the likelihood a CVE will be exploited in the wild in "
     "the next 30 days. Complements CVSS (which measures severity, not "
     "exploitation probability). CTI-feeding EPSS-augmented prioritization to "
     "Vulnerability Management is a common T3 practice."),

    ("IOC",
     "Tactical",
     "Indicator of Compromise. Atomic artifact signaling potentially malicious "
     "activity: IP address, domain, URL, file hash, email address, registry key. "
     "At the bottom of Bianco's Pyramid of Pain; easy for adversaries to change."),

    ("TTP",
     "Tactical / Behavioral",
     "Tactic, Technique, Procedure. The behavioral description of adversary actions, "
     "higher on the Pyramid of Pain than IOCs because TTPs are harder for "
     "adversaries to change. MITRE ATT&CK catalogs TTPs at the technique and "
     "sub-technique level."),

    ("ATT&CK",
     "Framework",
     "MITRE's knowledge base of adversary tactics, techniques, and sub-techniques "
     "based on real-world observations. Currently v18 with ~216 techniques and "
     "~475 sub-techniques. Organized by 14 tactics (Initial Access through Impact). "
     "ATT&CK IDs are the standard shorthand in CTI, detection, and hunt products."),

    ("Diamond Model",
     "Framework",
     "Intrusion analysis framework (Caltagirone, Pendergast, Betz 2013) with four "
     "vertices: Adversary, Capability, Infrastructure, Victim. Used to describe the "
     "relational structure of an intrusion and to pivot investigations. "
     "Complementary to ATT&CK, which describes what adversaries do; Diamond "
     "describes the relationships between actors, tools, victims, and infrastructure."),

    ("Cyber Kill Chain",
     "Framework",
     "Lockheed Martin's 7-phase model of an intrusion: Reconnaissance, Weaponization, "
     "Delivery, Exploitation, Installation, Command & Control, Actions on Objectives. "
     "Largely superseded by ATT&CK for granular analysis but still useful as a "
     "high-level campaign narrative structure."),

    ("Pyramid of Pain",
     "Framework",
     "David Bianco's concept ranking indicator types by how much pain they cause "
     "adversaries to change. Bottom (easy to change): hash values, IP addresses, "
     "domain names. Middle: network artifacts, host artifacts. Top (painful to "
     "change): tools, TTPs. Hunts and detections should aim for the upper tiers."),

    ("TaHiTI",
     "Hunt Methodology",
     "Targeted Hunting integrating Threat Intelligence. A threat-hunting methodology "
     "from FI-ISAC / Dutch Payments Association. Three phases: Initialize (generate "
     "hunt hypothesis abstract from CTI, IR, red team, or MaGMa gap), Hunt (refine "
     "hypothesis with CTI-enriched view, execute), Finalize (document findings; "
     "hand off to DE, IR, or back to CTI). The framework explicitly places CTI as "
     "the upstream producer of hunt hypotheses."),

    ("MaGMa",
     "Hunt Methodology",
     "Use Case Framework / tool that accompanies TaHiTI, used to structure hunt "
     "outcomes and measure hunt team performance over time."),

    ("NPS",
     "Metrics",
     "Net Promoter Score. \"How likely are you to recommend working with CTI on "
     "intelligence questions in your domain?\" on a 0-10 scale. Used in TIBMM to "
     "capture stakeholder satisfaction per domain on an annual cycle."),

    ("Brier Score",
     "Metrics",
     "A scoring rule for probabilistic forecasts. Lower is better (0 = perfect, "
     "1 = worst). Used to measure forecast calibration across many predictions: "
     "does a CTI judgment tagged \"likely (70-84%)\" come true roughly 70-84% of "
     "the time when aggregated? Feeds ICD 203 standard 8 (accurate judgments)."),

    ("ENISA CTL",
     "Reporting",
     "European Union Agency for Cybersecurity's Cyber Threat Landscape methodology. "
     "Standard structure for recurring threat landscape reporting: identify prime "
     "threats, map actors and motivations, describe techniques, summarize events, "
     "project trends. TIBMM references this as the methodology to follow for the "
     "quarterly CTL product."),

    ("PIPEDA",
     "Canadian Regulatory",
     "Personal Information Protection and Electronic Documents Act. Canada's "
     "federal privacy law governing how private-sector organizations collect, "
     "use, and disclose personal information. Relevant to CTI when collecting or "
     "storing indicators that contain personal data (emails, phone numbers, "
     "names), particularly from credential leaks or underground-forum observation."),

    ("CASL",
     "Canadian Regulatory",
     "Canada's Anti-Spam Legislation. Restricts commercial electronic messages; "
     "relevant to CTI when collecting phishing/smishing samples and when sharing "
     "threat samples externally."),

    ("CRTC",
     "Canadian Regulatory",
     "Canadian Radio-television and Telecommunications Commission. Federal "
     "telecom regulator. Issues cybersecurity expectations for carriers; requires "
     "reporting of certain network-security incidents. Strategic CTI products to "
     "leadership should synthesize CRTC expectations as part of regulatory context."),

    ("Bill C-26 / CCSPA",
     "Canadian Regulatory",
     "Critical Cyber Systems Protection Act (passed as part of Bill C-26, 2024). "
     "Imposes cybersecurity program, incident-reporting, and direction-compliance "
     "obligations on designated operators including Canadian telecommunications "
     "providers. CTI is likely to be a named input to the CCSPA-required "
     "cybersecurity program."),

    ("CCCS / CSE",
     "Canadian Regulatory",
     "Canadian Centre for Cyber Security, part of the Communications Security "
     "Establishment (CSE). Federal body issuing cyber advisories, threat "
     "bulletins, and sector guidance. Telco CTI should ingest CCCS advisories as "
     "a named source and maintain a CCCS liaison path."),

    ("CCTX",
     "Sharing",
     "Canadian Cyber Threat Exchange. National not-for-profit cross-sector threat "
     "intelligence sharing community. Most Canadian critical-infrastructure "
     "organizations, including telcos, participate. Membership is expected at "
     "TIBMM Tier-3+."),

    ("ISAC",
     "Sharing",
     "Information Sharing and Analysis Center. Industry-specific threat-sharing "
     "communities (e.g., FS-ISAC for finance, H-ISAC for health). The telco sector "
     "participates in the Communications ISAC (Comm-ISAC) in the US and similar "
     "fora elsewhere; CCTX is the Canadian cross-sector equivalent."),

    ("Smishing",
     "Telco Fraud",
     "SMS phishing. A fraud category particularly acute for telcos because the "
     "carrier's own infrastructure is the delivery channel. Smishing kits are "
     "collected, analyzed, and shared; indicators (sender numbers, lure URLs, "
     "kit infrastructure) feed SIEM/EDR and anti-fraud detection."),

    ("SIM Swap",
     "Telco Fraud",
     "Fraud technique where an attacker socially engineers or insider-assists a "
     "carrier to port a victim's phone number to an attacker-controlled SIM, "
     "intercepting SMS-based 2FA. CTI on SIM-swap actor TTPs (recruitment of "
     "insiders, tooling, target sectors) supports Fraud and IAM teams."),

    ("Number Porting Abuse",
     "Telco Fraud",
     "Closely related to SIM swap: abuse of Local Number Portability processes "
     "across carriers. Canadian telcos face this threat directly given CRTC "
     "porting rules; CTI tracks cross-carrier porting-abuse campaigns."),

    ("OSS / BSS",
     "Telco Stack",
     "Operational Support Systems / Business Support Systems. The carrier's "
     "back-end stack: network management, provisioning, billing, CRM, mediation. "
     "Vulnerability intelligence on OSS/BSS platforms (Amdocs, Netcracker, Oracle "
     "BRM, etc.) is a scored practice at T3."),

    ("RAN",
     "Telco Stack",
     "Radio Access Network. The wireless edge (cell sites, base stations, NodeB/"
     "eNodeB/gNodeB). RAN vendors include Ericsson, Nokia, Samsung, Huawei. "
     "Vulnerability and supply-chain intelligence on RAN vendors is in scope for "
     "TIBMM T3 Vuln and TPRM practices."),

    ("SS7 / Diameter / GTP",
     "Telco Stack",
     "Telecommunications signaling protocols. SS7 (legacy 2G/3G signaling), "
     "Diameter (4G/LTE signaling), GTP (GPRS/4G/5G user-plane and signaling). "
     "Each has well-documented abuse cases (location tracking, SMS interception, "
     "billing fraud, DoS). Vulnerability intelligence on these protocols, and "
     "changes in their threat landscape, are in scope for telco CTI."),

    ("5G / SBA",
     "Telco Stack",
     "5G's Service-Based Architecture introduces HTTP/2-based NF-to-NF signaling "
     "(Network Function service-based interfaces). New attack surface (NRF "
     "impersonation, AMF/SMF abuse). CTI on 5G-specific research and disclosed "
     "vulnerabilities feeds the T3 VM practice."),

    ("CPE",
     "Telco Stack",
     "Customer Premises Equipment. Consumer/business modems, routers, ONTs. CPE "
     "compromise is a telco-visible fraud and abuse vector (botnet recruitment, "
     "DDoS amplification). CPE-vendor supply-chain intelligence is a TPRM topic."),
]


def build_glossary_tab(wb: Workbook):
    """Alphabetical glossary of terms used throughout the TIBMM workbook."""
    ws = wb.create_sheet("Glossary")
    ws.sheet_properties.tabColor = STEEL

    col_widths = {"A": 3, "B": 26, "C": 22, "D": 85, "E": 3}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Title
    ws.merge_cells("A1:E1")
    ws.row_dimensions[1].height = 48
    c = ws["A1"]
    c.value = "  TIBMM Glossary"
    style_cell(c, FONT_TITLE, FILL_DARK_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCDE":
        ws[f"{col}1"].fill = FILL_DARK_NAVY

    # Subtitle
    ws.merge_cells("A2:E2")
    ws.row_dimensions[2].height = 24
    c = ws["A2"]
    c.value = ("  Definitions of terms referenced in the rubric and assessment. "
               "Use this tab if any term in a question is unfamiliar.")
    style_cell(c, FONT_SUBTITLE, FILL_NAVY, Alignment(horizontal="left", vertical="center"))
    for col in "ABCDE":
        ws[f"{col}2"].fill = FILL_NAVY

    # Headers
    row = 4
    ws.cell(row=row, column=2, value="Term")
    style_cell(ws.cell(row=row, column=2), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=3, value="Category")
    style_cell(ws.cell(row=row, column=3), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    ws.cell(row=row, column=4, value="Definition")
    style_cell(ws.cell(row=row, column=4), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    ws.row_dimensions[row].height = 26
    row += 1

    # Group by category, preserving order of first appearance
    category_order = []
    categorized = {}
    for term, cat, defn in GLOSSARY:
        if cat not in categorized:
            categorized[cat] = []
            category_order.append(cat)
        categorized[cat].append((term, defn))

    for cat in category_order:
        # Category band
        ws.merge_cells(f"B{row}:D{row}")
        ws.cell(row=row, column=2, value=f"  {cat}")
        style_cell(ws.cell(row=row, column=2), FONT_TIER_BANNER, FILL_STEEL,
                   Alignment(horizontal="left", vertical="center"))
        for c in range(3, 5):
            ws.cell(row=row, column=c).fill = FILL_STEEL
        ws.row_dimensions[row].height = 24
        row += 1

        # Entries, alphabetical within category
        for term, defn in sorted(categorized[cat], key=lambda t: t[0].lower()):
            ws.cell(row=row, column=2, value=term)
            style_cell(ws.cell(row=row, column=2), FONT_BODY_BOLD, FILL_WHITE, ALIGN_LEFT_TOP, HAIRLINE_BOTTOM)
            ws.cell(row=row, column=3, value=cat)
            style_cell(ws.cell(row=row, column=3), FONT_SMALL, FILL_WHITE, ALIGN_LEFT_TOP, HAIRLINE_BOTTOM)
            ws.cell(row=row, column=4, value=defn)
            style_cell(ws.cell(row=row, column=4), FONT_BODY, FILL_WHITE, ALIGN_WRAP, HAIRLINE_BOTTOM)
            # Row height scales loosely with definition length
            approx_lines = max(2, (len(defn) // 95) + 1)
            ws.row_dimensions[row].height = approx_lines * 15
            row += 1

    ws.freeze_panes = "A5"
    return ws


# ── Tab 8: Report Data ────────────────────────────────────────────────────────


def build_report_data_tab(wb: Workbook, rubric: dict, question_rows: list):
    """Build a flat data sheet optimized for Power BI / reporting consumption.

    Three tables:
      1. Summary (A1:B6) — key metrics
      2. Tier Progression (A9:F14) — tier status with Complete/Current/In Progress/Not Started
      3. Criterion Breakdown (A17:Gxx) — every criterion with section, category, score, level, status
    """
    ws = wb.create_sheet("Report Data")
    ws.sheet_properties.tabColor = STEEL

    col_widths = {"A": 18, "B": 30, "C": 22, "D": 34, "E": 12, "F": 16, "G": 16}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # ── Build criterion → question row mappings ───────────────────────
    crit_to_rows = {}
    for qr in question_rows:
        crit_to_rows.setdefault(qr["criterion"], []).append(qr)

    core_tier_ids_ordered = ["tier_0", "tier_1", "tier_2", "tier_3", "tier_4"]
    core_tier_names = {
        "tier_0": "Foundation", "tier_1": "Basic", "tier_2": "Operational",
        "tier_3": "Advanced", "tier_4": "Leading",
    }

    # Collect all tier → criteria structure
    all_tiers = []
    for tier in rubric["tiers"]:
        for crit in tier["criteria"]:
            all_tiers.append({
                "tier_id": tier["id"],
                "tier_name": tier["name"],
                "crit_id": crit["id"],
                "crit_name": crit["name"],
            })

    # ── Table 1: Summary ──────────────────────────────────────────────
    row = 1
    ws.cell(row=row, column=1, value="SUMMARY")
    style_cell(ws.cell(row=row, column=1), FONT_COL_HEADER, FILL_STEEL, ALIGN_LEFT, THIN_BORDER)
    ws.cell(row=row, column=2)
    style_cell(ws.cell(row=row, column=2), FONT_COL_HEADER, FILL_STEEL, ALIGN_LEFT, THIN_BORDER)
    row += 1

    summary_items = [
        ("Organization", "=Assessment!D4"),
        ("Assessor", "=Assessment!D5"),
        ("Date", "=Assessment!D7"),
        ("Assessment Type", "=Assessment!D8"),
    ]
    for label, formula in summary_items:
        ws.cell(row=row, column=1, value=label)
        style_cell(ws.cell(row=row, column=1), FONT_BODY_BOLD, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
        ws.cell(row=row, column=2, value=formula)
        style_cell(ws.cell(row=row, column=2), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
        row += 1

    # Overall score and tier will be filled after we compute criterion score cells
    overall_score_row = row
    ws.cell(row=row, column=1, value="Overall Score")
    style_cell(ws.cell(row=row, column=1), FONT_BODY_BOLD, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
    style_cell(ws.cell(row=row, column=2), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
    ws.cell(row=row, column=2).number_format = "0.00"
    row += 1

    achieved_tier_row = row
    ws.cell(row=row, column=1, value="Achieved Tier")
    style_cell(ws.cell(row=row, column=1), FONT_BODY_BOLD, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
    style_cell(ws.cell(row=row, column=2), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
    row += 1

    completion_row = row
    ws.cell(row=row, column=1, value="Completion")
    style_cell(ws.cell(row=row, column=1), FONT_BODY_BOLD, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
    total_q = len(question_rows)
    q_cells = ",".join(f"Assessment!E{qr['row']}" for qr in question_rows)
    ws.cell(row=row, column=2, value=f'=COUNTA({q_cells})&" / {total_q}"')
    style_cell(ws.cell(row=row, column=2), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)
    row += 2

    # ── Table 2: Tier Progression ─────────────────────────────────────
    tier_table_start = row
    tier_headers = ["Tier", "Tier Name", "Score", "Level", "Status", "Progression"]
    for col_idx, hdr in enumerate(tier_headers, 1):
        ws.cell(row=row, column=col_idx, value=hdr)
        style_cell(ws.cell(row=row, column=col_idx), FONT_COL_HEADER, FILL_NAVY, ALIGN_CENTER, THIN_BORDER)
    row += 1

    # Build tier score formulas from question data
    tier_score_rows = {}  # tier_id -> row in this table
    tier_score_cells = {}  # tier_id -> cell ref for the score

    for i, tid in enumerate(core_tier_ids_ordered):
        tier_name = core_tier_names[tid]
        tier_crits = [tc for tc in all_tiers if tc["tier_id"] == tid]

        ws.cell(row=row, column=1, value=f"T{i}")
        style_cell(ws.cell(row=row, column=1), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)

        ws.cell(row=row, column=2, value=tier_name)
        style_cell(ws.cell(row=row, column=2), FONT_BODY, FILL_WHITE, ALIGN_LEFT, THIN_BORDER)

        # Score = average of all question scores in this tier
        tier_q_refs = []
        for tc in tier_crits:
            for qr in crit_to_rows.get(tc["crit_id"], []):
                tier_q_refs.append(f"Assessment!F{qr['row']}")

        if tier_q_refs:
            refs_str = ",".join(tier_q_refs)
            ws.cell(row=row, column=3,
                    value=f'=IF(COUNT({refs_str})=0,"",AVERAGE({refs_str}))')
        ws.cell(row=row, column=3).number_format = "0.00"
        style_cell(ws.cell(row=row, column=3), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)

        # Level
        ws.cell(row=row, column=4, value=level_formula(f"C{row}"))
        style_cell(ws.cell(row=row, column=4), FONT_BODY, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)

        # Status (Pass / Below Target)
        ws.cell(row=row, column=5, value=status_formula(f"C{row}"))
        style_cell(ws.cell(row=row, column=5), FONT_BODY, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)

        # Progression will be filled after all tier scores exist
        style_cell(ws.cell(row=row, column=6), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, THIN_BORDER)

        tier_score_rows[tid] = row
        tier_score_cells[tid] = f"C{row}"
        row += 1

    # Progression and Achieved Tier formulas are deferred until after
    # Table 3 so we can use per-criterion checks (matching Dashboard logic).

    # Conditional formatting for progression column (applied now, formulas later)
    prog_range = f"F{tier_table_start + 1}:F{row - 1}"
    ws.conditional_formatting.add(
        prog_range,
        CellIsRule(operator="equal", formula=['"Complete"'], fill=FILL_SCORE_GREEN),
    )
    ws.conditional_formatting.add(
        prog_range,
        CellIsRule(operator="equal", formula=['"Current"'], fill=FILL_SCORE_YELLOW),
    )
    ws.conditional_formatting.add(
        prog_range,
        CellIsRule(operator="equal", formula=['"In Progress"'], fill=FILL_SCORE_ORANGE),
    )

    # Conditional formatting on tier scores
    apply_conditional_formatting(ws, f"C{tier_table_start + 1}:C{row - 1}")

    row += 1

    # ── Table 3: Criterion Breakdown ──────────────────────────────────
    crit_headers = ["Section", "Category", "Criterion", "Score", "Level", "Status"]
    for col_idx, hdr in enumerate(crit_headers, 1):
        ws.cell(row=row, column=col_idx, value=hdr)
        style_cell(ws.cell(row=row, column=col_idx), FONT_COL_HEADER, FILL_STEEL, ALIGN_CENTER, THIN_BORDER)
    crit_table_hdr = row
    row += 1

    all_crit_score_cells = []
    core_crit_cells_by_tier = {tid: [] for tid in core_tier_ids_ordered}

    for tc in all_tiers:
        crit_id = tc["crit_id"]
        tier_id = tc["tier_id"]
        rows_for_crit = crit_to_rows.get(crit_id, [])

        section = "TIBMM Core" if not is_enrichment(tier_id) else "Enrichment"
        category = tc["tier_name"].replace("Enrichment: ", "")

        ws.cell(row=row, column=1, value=section)
        style_cell(ws.cell(row=row, column=1), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=2, value=category)
        style_cell(ws.cell(row=row, column=2), FONT_BODY, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=3, value=tc["crit_name"])
        style_cell(ws.cell(row=row, column=3), FONT_BODY_BOLD, FILL_WHITE, ALIGN_LEFT, HAIRLINE_BOTTOM)

        if rows_for_crit:
            score_refs = [f"Assessment!F{qr['row']}" for qr in rows_for_crit]
            refs_str = ",".join(score_refs)
            ws.cell(row=row, column=4, value=f'=IF(COUNT({refs_str})=0,"",AVERAGE({refs_str}))')
        ws.cell(row=row, column=4).number_format = "0.00"
        style_cell(ws.cell(row=row, column=4), FONT_BODY_BOLD, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=5, value=level_formula(f"D{row}"))
        style_cell(ws.cell(row=row, column=5), FONT_BODY, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        ws.cell(row=row, column=6, value=status_formula(f"D{row}"))
        style_cell(ws.cell(row=row, column=6), FONT_BODY, FILL_WHITE, ALIGN_CENTER, HAIRLINE_BOTTOM)

        all_crit_score_cells.append(f"D{row}")
        if tier_id in core_crit_cells_by_tier:
            core_crit_cells_by_tier[tier_id].append(f"D{row}")
        row += 1

    # Conditional formatting on criterion scores
    apply_conditional_formatting(ws, f"D{crit_table_hdr + 1}:D{row - 1}")
    ws.conditional_formatting.add(
        f"F{crit_table_hdr + 1}:F{row - 1}",
        CellIsRule(operator="equal", formula=['"✓ Pass"'], fill=FILL_SCORE_GREEN),
    )
    ws.conditional_formatting.add(
        f"F{crit_table_hdr + 1}:F{row - 1}",
        CellIsRule(operator="equal", formula=['"✗ Below Target"'], fill=FILL_SCORE_RED),
    )

    # ── Fill in summary overall score and tier ────────────────────────
    if all_crit_score_cells:
        all_refs = ",".join(all_crit_score_cells)
        ws.cell(row=overall_score_row, column=2,
                value=f'=IF(COUNT({all_refs})=0,"",ROUND(AVERAGE({all_refs}),2))')
        ws.cell(row=overall_score_row, column=2).number_format = "0.00"

    # ── Per-criterion tier checks (matches Dashboard logic) ───────────
    # Tier passes only when ALL individual criteria in that tier score >= 3.0
    def tier_crit_check(tid):
        cells = core_crit_cells_by_tier.get(tid, [])
        if not cells:
            return "TRUE"
        return f"AND({','.join(f'ISNUMBER({c}),{c}>=3' for c in cells)})"

    # Fill in Progression column using per-criterion checks
    for i, tid in enumerate(core_tier_ids_ordered):
        trow = tier_score_rows[tid]

        # All tiers through this one pass? (per-criterion)
        cumul_checks = [tier_crit_check(core_tier_ids_ordered[j]) for j in range(i + 1)]
        all_pass = f"AND({','.join(cumul_checks)})"

        # All tiers through previous one pass?
        if i > 0:
            prev_checks = [tier_crit_check(core_tier_ids_ordered[j]) for j in range(i)]
            prev_pass = f"AND({','.join(prev_checks)})"
        else:
            prev_pass = "TRUE"

        # Progression: Complete / Current / In Progress / Not Started
        formula = (
            f'=IF({tier_score_cells[tid]}="","Not Started",'
            f'IF({all_pass},"Complete",'
            f'IF({prev_pass},"Current","In Progress")))'
        )
        ws.cell(row=trow, column=6, value=formula)

    # Achieved tier — cumulative per-criterion logic
    tier_display = {
        "tier_0": "Tier 0: Foundation", "tier_1": "Tier 1: Basic",
        "tier_2": "Tier 2: Operational", "tier_3": "Tier 3: Advanced",
        "tier_4": "Tier 4: Leading",
    }

    cumul_report = {}
    for i, tid in enumerate(core_tier_ids_ordered):
        checks = [tier_crit_check(core_tier_ids_ordered[j]) for j in range(i + 1)]
        cumul_report[tid] = f"AND({','.join(checks)})"

    tier_f = f'=IF({cumul_report["tier_4"]},"{tier_display["tier_4"]}",'
    tier_f += f'IF({cumul_report["tier_3"]},"{tier_display["tier_3"]}",'
    tier_f += f'IF({cumul_report["tier_2"]},"{tier_display["tier_2"]}",'
    tier_f += f'IF({cumul_report["tier_1"]},"{tier_display["tier_1"]}",'
    tier_f += f'IF({cumul_report["tier_0"]},"{tier_display["tier_0"]}",'
    tier_f += '"Below Foundation")))))'
    ws.cell(row=achieved_tier_row, column=2, value=tier_f)

    ws.freeze_panes = "A2"
    return ws


# ── Main ──────────────────────────────────────────────────────────────────────


def generate_spreadsheet(
    output_path: Path,
    rubric_path: Path = DEFAULT_RUBRIC,
    questionnaire_path: Path = DEFAULT_QUESTIONNAIRE,
    mode: str = "self",
):
    rubric = load_yaml(rubric_path)
    questionnaire = load_yaml(questionnaire_path)

    wb = Workbook()
    build_instructions_tab(wb, mode)
    _, question_rows, header_row = build_assessment_tab(wb, questionnaire, rubric, mode)
    _, core_tier_row_list, enrich_row_list = build_dashboard_tab(
        wb, rubric, questionnaire, question_rows, header_row)
    build_core_chart_tab(wb, core_tier_row_list)
    build_enrichment_chart_tab(wb, enrich_row_list)
    build_rubric_tab(wb, rubric)
    build_glossary_tab(wb)
    build_report_data_tab(wb, rubric, question_rows)

    wb.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Generate TIBMM assessment workbook.")
    parser.add_argument(
        "--output", "-o", type=Path,
        default=PROJECT_ROOT / "templates" / "TIBMM-Assessment-Workbook.xlsx",
        help="Output Excel file path (default: templates/TIBMM-Assessment-Workbook.xlsx)",
    )
    parser.add_argument(
        "--mode", choices=["self", "audit"], default="self",
        help="Assessment mode: self-assessment or audit (default: self)",
    )
    parser.add_argument("--rubric", type=Path, default=DEFAULT_RUBRIC)
    parser.add_argument("--questionnaire", type=Path, default=DEFAULT_QUESTIONNAIRE)

    args = parser.parse_args()
    output = generate_spreadsheet(args.output, args.rubric, args.questionnaire, args.mode)
    print(f"Spreadsheet generated: {output}")


if __name__ == "__main__":
    main()
