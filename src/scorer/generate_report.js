const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// ── Read JSON data (extracted from Excel via Python) ──
const jsonPath = process.argv[2];
const outPath = process.argv[3] || "tibmm-report.pptx";

if (!jsonPath) {
  console.error("Usage: node scorer/generate_report.js <data.json> [output.pptx]");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));

const org = data.org;
const assessor = data.assessor;
const assessDate = data.date || new Date().toISOString().split("T")[0];
const assessType = data.type;
const overallScore = data.overallScore;
const achievedTier = data.achievedTier;
const completion = data.completion;
const tiers = data.tiers;
const criteria = data.criteria;

const tibmmCore = criteria.filter((c) => c.section === "TIBMM Core");
const enrichment = criteria.filter((c) => c.section === "Enrichment");
const passing = criteria.filter((c) => c.status.includes("Pass")).length;
const failing = criteria.filter((c) => c.status.includes("Below")).length;

// ── Colors ──
const C = {
  bgDark: "0B1120",
  bgCard: "111A2E",
  bgCardAlt: "162038",
  border: "1E2D4A",
  textPrimary: "E2E8F0",
  textSecondary: "8B9DC3",
  textMuted: "5A6F94",
  green: "10B981",
  blue: "3B82F6",
  yellow: "F59E0B",
  orange: "F97316",
  cyan: "06B6D4",
  purple: "8B5CF6",
  red: "EF4444",
  white: "FFFFFF",
};

const tierColors = {
  T0: C.green,
  T1: C.blue,
  T2: C.yellow,
  T3: C.orange,
  T4: C.purple,
};

const levelColors = {
  Optimized: C.green,
  Managed: C.blue,
  Defined: C.yellow,
  Repeatable: C.orange,
  Initial: C.red,
};

function getLevelColor(level) {
  return levelColors[level] || C.textMuted;
}

function getScoreColor(score) {
  if (score >= 4.5) return C.green;
  if (score >= 3.5) return C.cyan;
  if (score >= 3.0) return C.yellow;
  if (score >= 2.0) return C.orange;
  return C.red;
}

// Fresh shadow factory (pptxgenjs mutates objects)
const mkShadow = () => ({
  type: "outer",
  blur: 4,
  offset: 2,
  angle: 135,
  color: "000000",
  opacity: 0.3,
});

// ── Build Presentation ──
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10" x 5.625"
pres.author = assessor;
pres.title = "TIBMM Assessment Report";

// ================================================================
// SLIDE 1: Title Slide
// ================================================================
const s1 = pres.addSlide();
s1.background = { color: C.bgDark };

// Top accent bar
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: 0,
  w: 10,
  h: 0.06,
  fill: { color: C.cyan },
});

// Section label
s1.addText("THREAT INTELLIGENCE BEHAVIOR MATURITY MODEL", {
  x: 0.6,
  y: 1.2,
  w: 8.8,
  h: 0.4,
  fontSize: 11,
  fontFace: "Consolas",
  color: C.cyan,
  charSpacing: 3,
  margin: 0,
});

// Main title
s1.addText("TIBMM Assessment Report", {
  x: 0.6,
  y: 1.65,
  w: 8.8,
  h: 0.8,
  fontSize: 36,
  fontFace: "Calibri",
  bold: true,
  color: C.white,
  margin: 0,
});

// Subtitle
s1.addText(
  `${assessType} — ${typeof assessDate === "number" ? "Current Period" : assessDate}`,
  {
    x: 0.6,
    y: 2.5,
    w: 8.8,
    h: 0.4,
    fontSize: 16,
    fontFace: "Calibri",
    color: C.textSecondary,
    margin: 0,
  }
);

// Bottom info bar
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: 4.6,
  w: 10,
  h: 1.025,
  fill: { color: C.bgCard },
});

const infoItems = [
  { label: "OVERALL SCORE", value: overallScore.toFixed(2) + " / 5.00" },
  { label: "ACHIEVED TIER", value: achievedTier },
  { label: "COMPLETION", value: completion },
  {
    label: "PASS / FAIL",
    value: `${passing} Pass  ·  ${failing} Below Target`,
  },
];

infoItems.forEach((item, i) => {
  const ix = 0.6 + i * 2.35;
  s1.addText(item.label, {
    x: ix,
    y: 4.72,
    w: 2.2,
    h: 0.25,
    fontSize: 9,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });
  s1.addText(item.value, {
    x: ix,
    y: 4.98,
    w: 2.2,
    h: 0.35,
    fontSize: 14,
    fontFace: "Calibri",
    bold: true,
    color: C.textPrimary,
    margin: 0,
  });
});

// ================================================================
// SLIDE 2: Tier Progression + KPI Summary
// ================================================================
const s2 = pres.addSlide();
s2.background = { color: C.bgDark };

// Top accent
s2.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: 0,
  w: 10,
  h: 0.06,
  fill: { color: C.cyan },
});

// Section label
s2.addText("01  TIER PROGRESSION", {
  x: 0.6,
  y: 0.25,
  w: 8,
  h: 0.3,
  fontSize: 10,
  fontFace: "Consolas",
  color: C.textMuted,
  charSpacing: 2,
  margin: 0,
});

// Title
s2.addText("TIBMM Tier Overview", {
  x: 0.6,
  y: 0.55,
  w: 8,
  h: 0.5,
  fontSize: 24,
  fontFace: "Calibri",
  bold: true,
  color: C.white,
  margin: 0,
});

// 5 Tier KPI cards across the top
const cardW = 1.68;
const cardGap = 0.15;
const cardStartX = 0.6;
const cardY = 1.25;
const cardH = 1.5;

tiers.forEach((t, i) => {
  const cx = cardStartX + i * (cardW + cardGap);
  const tc = tierColors[t.id] || C.textMuted;

  // Card background
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx,
    y: cardY,
    w: cardW,
    h: cardH,
    fill: { color: C.bgCard },
    shadow: mkShadow(),
  });

  // Top color accent
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx,
    y: cardY,
    w: cardW,
    h: 0.05,
    fill: { color: tc },
  });

  // Tier label
  s2.addText(`${t.id} — ${t.name}`, {
    x: cx + 0.12,
    y: cardY + 0.15,
    w: cardW - 0.24,
    h: 0.22,
    fontSize: 8,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });

  // Score
  s2.addText(t.score.toFixed(2), {
    x: cx + 0.12,
    y: cardY + 0.38,
    w: cardW - 0.24,
    h: 0.42,
    fontSize: 28,
    fontFace: "Consolas",
    bold: true,
    color: tc,
    margin: 0,
  });

  // Level
  s2.addText(t.level, {
    x: cx + 0.12,
    y: cardY + 0.78,
    w: cardW - 0.24,
    h: 0.2,
    fontSize: 11,
    fontFace: "Calibri",
    color: C.textSecondary,
    margin: 0,
  });

  // Progress bar background
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx + 0.12,
    y: cardY + 1.1,
    w: cardW - 0.24,
    h: 0.08,
    fill: { color: C.bgDark },
  });

  // Progress bar fill
  const pct = Math.min(t.score / 5, 1);
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx + 0.12,
    y: cardY + 1.1,
    w: (cardW - 0.24) * pct,
    h: 0.08,
    fill: { color: tc },
  });

  // Progression status
  s2.addText(t.progression, {
    x: cx + 0.12,
    y: cardY + 1.24,
    w: cardW - 0.24,
    h: 0.18,
    fontSize: 8,
    fontFace: "Calibri",
    color: t.progression === "Complete" ? C.green : C.textMuted,
    margin: 0,
  });
});

// ── Bottom section: Summary stats ──
const statY = 3.05;
s2.addText("ASSESSMENT SUMMARY", {
  x: 0.6,
  y: statY,
  w: 8,
  h: 0.25,
  fontSize: 10,
  fontFace: "Consolas",
  color: C.textMuted,
  charSpacing: 2,
  margin: 0,
});

// Summary cards row
const summCards = [
  {
    label: "Overall Score",
    value: overallScore.toFixed(2),
    sub: "out of 5.00",
    color: C.cyan,
  },
  {
    label: "Achieved Tier",
    value: achievedTier.replace("Tier ", "T"),
    sub: "",
    color: C.cyan,
  },
  { label: "Passing", value: String(passing), sub: "criteria", color: C.green },
  {
    label: "Below Target",
    value: String(failing),
    sub: "criteria",
    color: failing > 0 ? C.red : C.green,
  },
  {
    label: "TIBMM Core Avg",
    value: (
      tibmmCore.reduce((a, c) => a + c.score, 0) / tibmmCore.length
    ).toFixed(2),
    sub: "",
    color: C.cyan,
  },
  {
    label: "Enrichment Avg",
    value: (
      enrichment.reduce((a, c) => a + c.score, 0) / enrichment.length
    ).toFixed(2),
    sub: "",
    color: C.green,
  },
];

const scW = 1.35;
const scGap = 0.12;
summCards.forEach((sc, i) => {
  const sx = 0.6 + i * (scW + scGap);
  const sy = statY + 0.35;

  s2.addShape(pres.shapes.RECTANGLE, {
    x: sx,
    y: sy,
    w: scW,
    h: 1.1,
    fill: { color: C.bgCard },
    shadow: mkShadow(),
  });

  s2.addText(sc.label, {
    x: sx + 0.1,
    y: sy + 0.1,
    w: scW - 0.2,
    h: 0.18,
    fontSize: 8,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });

  // Auto-size: shrink font if value text is long
  const valFontSize = sc.value.length > 10 ? 13 : sc.value.length > 6 ? 18 : 24;
  s2.addText(sc.value, {
    x: sx + 0.1,
    y: sy + 0.32,
    w: scW - 0.2,
    h: 0.42,
    fontSize: valFontSize,
    fontFace: "Consolas",
    bold: true,
    color: sc.color,
    valign: "middle",
    margin: 0,
  });

  if (sc.sub) {
    s2.addText(sc.sub, {
      x: sx + 0.1,
      y: sy + 0.75,
      w: scW - 0.2,
      h: 0.18,
      fontSize: 9,
      fontFace: "Calibri",
      color: C.textSecondary,
      margin: 0,
    });
  }
});

// Bottom bar - tier progression visual
const barY = 4.8;
s2.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: barY,
  w: 10,
  h: 0.825,
  fill: { color: C.bgCard },
});

s2.addText("TIER PROGRESSION", {
  x: 0.6,
  y: barY + 0.08,
  w: 3,
  h: 0.18,
  fontSize: 8,
  fontFace: "Consolas",
  color: C.textMuted,
  charSpacing: 2,
  margin: 0,
});

const segW = 1.68;
tiers.forEach((t, i) => {
  const sx = 0.6 + i * (segW + 0.08);
  const sy = barY + 0.32;

  let segColor = C.bgCardAlt;
  let borderColor = C.border;
  let labelColor = C.textMuted;

  if (t.progression === "Complete") {
    segColor = "0D2818";
    borderColor = C.green;
    labelColor = C.green;
  } else if (t.progression === "Current") {
    segColor = "0B1F2E";
    borderColor = C.cyan;
    labelColor = C.cyan;
  } else if (t.progression === "In Progress") {
    segColor = "1A1708";
    borderColor = C.yellow;
    labelColor = C.yellow;
  }

  s2.addShape(pres.shapes.RECTANGLE, {
    x: sx,
    y: sy,
    w: segW,
    h: 0.4,
    fill: { color: segColor },
    line: { color: borderColor, width: 1 },
  });

  s2.addText(`${t.id}  ${t.name}`, {
    x: sx + 0.1,
    y: sy + 0.02,
    w: segW - 0.2,
    h: 0.2,
    fontSize: 9,
    fontFace: "Consolas",
    bold: true,
    color: labelColor,
    margin: 0,
  });

  s2.addText(t.progression, {
    x: sx + 0.1,
    y: sy + 0.2,
    w: segW - 0.2,
    h: 0.16,
    fontSize: 8,
    fontFace: "Calibri",
    color: labelColor,
    margin: 0,
  });
});

// ================================================================
// SLIDE 3: TIBMM Core Criteria Detail
// ================================================================
const s3 = pres.addSlide();
s3.background = { color: C.bgDark };

s3.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: 0,
  w: 10,
  h: 0.06,
  fill: { color: C.cyan },
});

s3.addText("02  TIBMM CORE", {
  x: 0.6,
  y: 0.2,
  w: 8,
  h: 0.25,
  fontSize: 10,
  fontFace: "Consolas",
  color: C.textMuted,
  charSpacing: 2,
  margin: 0,
});

s3.addText("Core Criteria Breakdown", {
  x: 0.6,
  y: 0.45,
  w: 8,
  h: 0.4,
  fontSize: 22,
  fontFace: "Calibri",
  bold: true,
  color: C.white,
  margin: 0,
});

// Table header
const tblHeaderY = 0.95;
const colX = [0.6, 3.2, 5.6, 6.7, 7.65, 8.6];
const colLabels = ["Criterion", "Category", "Level", "Score", "Bar", "Status"];
const colW = [2.6, 2.4, 1.1, 0.95, 0.95, 0.8];

// Header background
s3.addShape(pres.shapes.RECTANGLE, {
  x: 0.5,
  y: tblHeaderY,
  w: 9.0,
  h: 0.28,
  fill: { color: C.bgCard },
});

colLabels.forEach((lbl, i) => {
  s3.addText(lbl.toUpperCase(), {
    x: colX[i],
    y: tblHeaderY + 0.02,
    w: colW[i],
    h: 0.24,
    fontSize: 7,
    fontFace: "Consolas",
    color: C.textMuted,
    charSpacing: 1,
    margin: 0,
  });
});

// Data rows
const rowH = 0.21;
const startY = tblHeaderY + 0.3;

tibmmCore.forEach((c, i) => {
  const ry = startY + i * (rowH + 0.01);
  const isBelow = c.status.includes("Below");
  const sc = getScoreColor(c.score);

  // Alternating row bg
  if (i % 2 === 0) {
    s3.addShape(pres.shapes.RECTANGLE, {
      x: 0.5,
      y: ry - 0.02,
      w: 9.0,
      h: rowH + 0.02,
      fill: { color: C.bgCard },
    });
  }

  // Criterion name
  s3.addText(c.criterion, {
    x: colX[0],
    y: ry,
    w: colW[0],
    h: rowH,
    fontSize: 9,
    fontFace: "Calibri",
    color: isBelow ? C.red : C.textPrimary,
    margin: 0,
  });

  // Category
  s3.addText(c.category, {
    x: colX[1],
    y: ry,
    w: colW[1],
    h: rowH,
    fontSize: 9,
    fontFace: "Calibri",
    color: C.textSecondary,
    margin: 0,
  });

  // Level badge
  s3.addShape(pres.shapes.RECTANGLE, {
    x: colX[2],
    y: ry + 0.02,
    w: 0.95,
    h: 0.18,
    fill: { color: C.bgCardAlt },
  });
  s3.addText(c.level, {
    x: colX[2],
    y: ry + 0.02,
    w: 0.95,
    h: 0.18,
    fontSize: 7,
    fontFace: "Consolas",
    bold: true,
    color: getLevelColor(c.level),
    align: "center",
    valign: "middle",
    margin: 0,
  });

  // Score
  s3.addText(c.score.toFixed(2), {
    x: colX[3],
    y: ry,
    w: colW[3],
    h: rowH,
    fontSize: 10,
    fontFace: "Consolas",
    bold: true,
    color: sc,
    align: "center",
    margin: 0,
  });

  // Score bar
  const barW = 0.85;
  s3.addShape(pres.shapes.RECTANGLE, {
    x: colX[4],
    y: ry + 0.07,
    w: barW,
    h: 0.08,
    fill: { color: C.bgDark },
  });
  s3.addShape(pres.shapes.RECTANGLE, {
    x: colX[4],
    y: ry + 0.07,
    w: barW * Math.min(c.score / 5, 1),
    h: 0.08,
    fill: { color: sc },
  });

  // Status icon
  s3.addText(isBelow ? "✗" : "✓", {
    x: colX[5],
    y: ry,
    w: colW[5],
    h: rowH,
    fontSize: 11,
    fontFace: "Calibri",
    bold: true,
    color: isBelow ? C.red : C.green,
    align: "center",
    margin: 0,
  });
});

// ================================================================
// SLIDE 4: Enrichment Criteria
// ================================================================
const s4 = pres.addSlide();
s4.background = { color: C.bgDark };

s4.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: 0,
  w: 10,
  h: 0.06,
  fill: { color: C.cyan },
});

s4.addText("03  ENRICHMENT", {
  x: 0.6,
  y: 0.2,
  w: 8,
  h: 0.25,
  fontSize: 10,
  fontFace: "Consolas",
  color: C.textMuted,
  charSpacing: 2,
  margin: 0,
});

s4.addText("Tradecraft & Governance", {
  x: 0.6,
  y: 0.45,
  w: 8,
  h: 0.4,
  fontSize: 22,
  fontFace: "Calibri",
  bold: true,
  color: C.white,
  margin: 0,
});

// Enrichment as cards (2 columns x 3 rows)
const enrCardW = 4.1;
const enrCardH = 1.15;
const enrGap = 0.2;
const enrStartY = 1.1;

enrichment.forEach((c, i) => {
  const col = i % 2;
  const row = Math.floor(i / 2);
  const cx = 0.6 + col * (enrCardW + enrGap);
  const cy = enrStartY + row * (enrCardH + enrGap);
  const sc = getScoreColor(c.score);
  const lc = getLevelColor(c.level);

  // Card bg
  s4.addShape(pres.shapes.RECTANGLE, {
    x: cx,
    y: cy,
    w: enrCardW,
    h: enrCardH,
    fill: { color: C.bgCard },
    shadow: mkShadow(),
  });

  // Left accent
  s4.addShape(pres.shapes.RECTANGLE, {
    x: cx,
    y: cy,
    w: 0.06,
    h: enrCardH,
    fill: { color: lc },
  });

  // Category label
  s4.addText(c.category.toUpperCase(), {
    x: cx + 0.2,
    y: cy + 0.1,
    w: enrCardW - 1.2,
    h: 0.18,
    fontSize: 7,
    fontFace: "Consolas",
    color: C.textMuted,
    charSpacing: 1,
    margin: 0,
  });

  // Criterion name
  s4.addText(c.criterion, {
    x: cx + 0.2,
    y: cy + 0.3,
    w: enrCardW - 1.2,
    h: 0.25,
    fontSize: 12,
    fontFace: "Calibri",
    bold: true,
    color: C.textPrimary,
    margin: 0,
  });

  // Level badge
  s4.addShape(pres.shapes.RECTANGLE, {
    x: cx + 0.2,
    y: cy + 0.65,
    w: 0.8,
    h: 0.22,
    fill: { color: C.bgCardAlt },
  });
  s4.addText(c.level, {
    x: cx + 0.2,
    y: cy + 0.65,
    w: 0.8,
    h: 0.22,
    fontSize: 8,
    fontFace: "Consolas",
    bold: true,
    color: lc,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  // Status text
  s4.addText(c.status, {
    x: cx + 1.1,
    y: cy + 0.65,
    w: 1.2,
    h: 0.22,
    fontSize: 8,
    fontFace: "Calibri",
    color: c.status.includes("Pass") ? C.green : C.red,
    valign: "middle",
    margin: 0,
  });

  // Score (big, right side)
  s4.addText(c.score.toFixed(2), {
    x: cx + enrCardW - 1.1,
    y: cy + 0.15,
    w: 0.95,
    h: 0.55,
    fontSize: 28,
    fontFace: "Consolas",
    bold: true,
    color: sc,
    align: "center",
    valign: "middle",
    margin: 0,
  });

  // Score bar
  s4.addShape(pres.shapes.RECTANGLE, {
    x: cx + enrCardW - 1.1,
    y: cy + 0.78,
    w: 0.9,
    h: 0.08,
    fill: { color: C.bgDark },
  });
  s4.addShape(pres.shapes.RECTANGLE, {
    x: cx + enrCardW - 1.1,
    y: cy + 0.78,
    w: 0.9 * Math.min(c.score / 5, 1),
    h: 0.08,
    fill: { color: sc },
  });
});

// Bottom summary bar
const enrBarY = 4.6;
s4.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: enrBarY,
  w: 10,
  h: 1.025,
  fill: { color: C.bgCard },
});

// Split enrichment by category (Tradecraft vs Governance)
const tradecraft = enrichment.filter((c) =>
  c.category.includes("Tradecraft")
);
const governance = enrichment.filter((c) =>
  c.category.includes("Governance")
);
const tcAvg =
  tradecraft.length
    ? tradecraft.reduce((a, c) => a + c.score, 0) / tradecraft.length
    : 0;
const gvAvg =
  governance.length
    ? governance.reduce((a, c) => a + c.score, 0) / governance.length
    : 0;
const enrAvg =
  enrichment.reduce((a, c) => a + c.score, 0) / enrichment.length;

const enrStats = [
  { label: "ENRICHMENT AVG", value: enrAvg.toFixed(2), color: C.cyan },
  { label: "TRADECRAFT", value: tcAvg.toFixed(2), color: C.green },
  { label: "GOVERNANCE", value: gvAvg.toFixed(2), color: C.blue },
  {
    label: "ALL PASSING",
    value: enrichment.every((c) => c.status.includes("Pass")) ? "YES" : "NO",
    color: enrichment.every((c) => c.status.includes("Pass"))
      ? C.green
      : C.red,
  },
];

enrStats.forEach((st, i) => {
  const sx = 0.6 + i * 2.35;
  s4.addText(st.label, {
    x: sx,
    y: enrBarY + 0.15,
    w: 2.1,
    h: 0.2,
    fontSize: 9,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });
  s4.addText(st.value, {
    x: sx,
    y: enrBarY + 0.4,
    w: 2.1,
    h: 0.4,
    fontSize: 22,
    fontFace: "Consolas",
    bold: true,
    color: st.color,
    margin: 0,
  });
});

// ================================================================
// Write file
// ================================================================
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("TIBMM report generated: " + outPath);
});
