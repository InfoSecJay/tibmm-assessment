const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ── CLI ──
const historyPath = process.argv[2];
const outPath = process.argv[3] || "tibmm-trend.pptx";

if (!historyPath) {
  console.error(
    "Usage: node scorer/generate_trend.js <history.json> [output.pptx]"
  );
  process.exit(1);
}

const history = JSON.parse(fs.readFileSync(historyPath, "utf-8"));
if (!Array.isArray(history) || history.length === 0) {
  console.error("Error: history.json is empty or not an array.");
  process.exit(1);
}

// Sort chronologically
history.sort((a, b) => (a.date || "").localeCompare(b.date || ""));

const latest = history[history.length - 1];
const prev = history.length >= 2 ? history[history.length - 2] : null;
const hasMultiple = history.length >= 2;

// ── Colors (matching generate_report.js) ──
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

function getScoreColor(score) {
  if (score >= 4.5) return C.green;
  if (score >= 3.5) return C.cyan;
  if (score >= 3.0) return C.yellow;
  if (score >= 2.0) return C.orange;
  return C.red;
}

function getDeltaColor(delta) {
  if (delta > 0.1) return C.green;
  if (delta < -0.1) return C.red;
  return C.textMuted;
}

function getDeltaArrow(delta) {
  if (delta > 0.1) return "\u25B2"; // ▲
  if (delta < -0.1) return "\u25BC"; // ▼
  return "\u2014"; // —
}

function fmtScore(s) {
  return s != null ? s.toFixed(2) : "N/A";
}

function fmtDelta(d) {
  if (d == null) return "N/A";
  const sign = d > 0 ? "+" : "";
  return sign + d.toFixed(2);
}

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
pres.layout = "LAYOUT_16x9";
pres.author = latest.assessor || "TIBMM";
pres.title = "TIBMM Trend Report";

// ================================================================
// SLIDE 1: Score Trajectory
// ================================================================
const s1 = pres.addSlide();
s1.background = { color: C.bgDark };

// Top accent
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: 0,
  w: 10,
  h: 0.06,
  fill: { color: C.cyan },
});

s1.addText("MATURITY TREND", {
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

s1.addText("Score Trajectory", {
  x: 0.6,
  y: 0.45,
  w: 6,
  h: 0.45,
  fontSize: 24,
  fontFace: "Calibri",
  bold: true,
  color: C.white,
  margin: 0,
});

// Period range
const rangeText =
  history.length === 1
    ? `Baseline: ${latest.date}`
    : `${history[0].date} \u2014 ${latest.date}  (${history.length} assessments)`;
s1.addText(rangeText, {
  x: 0.6,
  y: 0.85,
  w: 8,
  h: 0.25,
  fontSize: 11,
  fontFace: "Calibri",
  color: C.textSecondary,
  margin: 0,
});

if (!hasMultiple) {
  // Single entry — baseline card
  s1.addShape(pres.shapes.RECTANGLE, {
    x: 1.5,
    y: 1.8,
    w: 7,
    h: 2.5,
    fill: { color: C.bgCard },
    shadow: mkShadow(),
  });

  s1.addText("BASELINE ESTABLISHED", {
    x: 1.5,
    y: 2.0,
    w: 7,
    h: 0.4,
    fontSize: 14,
    fontFace: "Consolas",
    bold: true,
    color: C.cyan,
    align: "center",
    margin: 0,
  });

  s1.addText(
    `Overall Score: ${fmtScore(latest.overallScore)} / 5.00\nAchieved Tier: ${latest.achievedTier}\nCompletion: ${latest.completion}`,
    {
      x: 2,
      y: 2.6,
      w: 6,
      h: 1.2,
      fontSize: 16,
      fontFace: "Calibri",
      color: C.textPrimary,
      align: "center",
      lineSpacingMultiple: 1.5,
      margin: 0,
    }
  );

  s1.addText(
    "Complete your next assessment to start tracking trends.",
    {
      x: 2,
      y: 3.8,
      w: 6,
      h: 0.3,
      fontSize: 11,
      fontFace: "Calibri",
      color: C.textMuted,
      align: "center",
      margin: 0,
    }
  );
} else {
  // Multi-entry — chart area
  const chartX = 0.6;
  const chartY = 1.3;
  const chartW = 8.8;
  const chartH = 3.0;

  // Chart background
  s1.addShape(pres.shapes.RECTANGLE, {
    x: chartX,
    y: chartY,
    w: chartW,
    h: chartH,
    fill: { color: C.bgCard },
    shadow: mkShadow(),
  });

  // Y-axis labels and grid lines
  const yMin = 0;
  const yMax = 5;
  for (let y = 0; y <= 5; y++) {
    const yPos = chartY + chartH - (y / yMax) * chartH;
    // Grid line
    s1.addShape(pres.shapes.LINE, {
      x: chartX + 0.4,
      y: yPos,
      w: chartW - 0.6,
      h: 0,
      line: { color: y === 3 ? C.yellow : C.border, width: y === 3 ? 1.5 : 0.5, dashType: y === 3 ? "dash" : "solid" },
    });
    // Label
    s1.addText(String(y), {
      x: chartX + 0.05,
      y: yPos - 0.1,
      w: 0.3,
      h: 0.2,
      fontSize: 8,
      fontFace: "Consolas",
      color: C.textMuted,
      align: "right",
      margin: 0,
    });
  }

  // 3.0 threshold label
  const thresholdY = chartY + chartH - (3.0 / yMax) * chartH;
  s1.addText("3.0 threshold", {
    x: chartX + chartW - 1.5,
    y: thresholdY - 0.22,
    w: 1.3,
    h: 0.18,
    fontSize: 7,
    fontFace: "Consolas",
    color: C.yellow,
    align: "right",
    margin: 0,
  });

  // Plot points and lines
  const pointSpacing = (chartW - 0.8) / Math.max(history.length - 1, 1);
  const points = history.map((entry, i) => {
    const x = chartX + 0.5 + i * pointSpacing;
    const yFrac = (entry.overallScore || 0) / yMax;
    const y = chartY + chartH - yFrac * chartH;
    return { x, y, score: entry.overallScore || 0, date: entry.date, tier: entry.achievedTier };
  });

  // Connect points with lines
  for (let i = 1; i < points.length; i++) {
    const p1 = points[i - 1];
    const p2 = points[i];
    // Use a thin rectangle as a line approximation (pptxgenjs LINE is tricky for angled)
    s1.addShape(pres.shapes.LINE, {
      x: p1.x,
      y: p1.y,
      w: p2.x - p1.x,
      h: p2.y - p1.y,
      line: { color: C.cyan, width: 2.5 },
    });
  }

  // Plot data points as circles and labels
  points.forEach((p, i) => {
    const dotSize = 0.16;
    s1.addShape(pres.shapes.OVAL, {
      x: p.x - dotSize / 2,
      y: p.y - dotSize / 2,
      w: dotSize,
      h: dotSize,
      fill: { color: C.cyan },
    });

    // Score label above dot
    s1.addText(fmtScore(p.score), {
      x: p.x - 0.3,
      y: p.y - 0.28,
      w: 0.6,
      h: 0.18,
      fontSize: 8,
      fontFace: "Consolas",
      bold: true,
      color: C.white,
      align: "center",
      margin: 0,
    });

    // Date label below chart
    s1.addText(p.date, {
      x: p.x - 0.35,
      y: chartY + chartH + 0.05,
      w: 0.7,
      h: 0.18,
      fontSize: 7,
      fontFace: "Consolas",
      color: C.textMuted,
      align: "center",
      margin: 0,
    });
  });

  // Tier achievement badges
  let lastTier = null;
  points.forEach((p) => {
    if (p.tier !== lastTier && p.tier !== "Below Foundation") {
      s1.addText(p.tier.replace("Tier ", "T"), {
        x: p.x - 0.4,
        y: p.y + 0.14,
        w: 0.8,
        h: 0.16,
        fontSize: 6,
        fontFace: "Consolas",
        bold: true,
        color: C.green,
        align: "center",
        margin: 0,
      });
    }
    lastTier = p.tier;
  });
}

// Bottom bar — current vs previous
const s1BarY = 4.6;
s1.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: s1BarY,
  w: 10,
  h: 1.025,
  fill: { color: C.bgCard },
});

const overallDelta = prev
  ? latest.overallScore - prev.overallScore
  : null;
const tierChanged = prev && latest.achievedTier !== prev.achievedTier;

const s1Stats = [
  { label: "CURRENT SCORE", value: fmtScore(latest.overallScore), color: C.cyan },
  {
    label: "CHANGE",
    value: overallDelta != null ? `${getDeltaArrow(overallDelta)} ${fmtDelta(overallDelta)}` : "Baseline",
    color: overallDelta != null ? getDeltaColor(overallDelta) : C.textMuted,
  },
  { label: "ACHIEVED TIER", value: latest.achievedTier.replace("Tier ", "T"), color: C.cyan },
  {
    label: tierChanged ? "TIER CHANGED" : "ASSESSMENTS",
    value: tierChanged ? `from ${prev.achievedTier.replace("Tier ", "T")}` : String(history.length),
    color: tierChanged ? C.green : C.textSecondary,
  },
];

s1Stats.forEach((st, i) => {
  const sx = 0.6 + i * 2.35;
  s1.addText(st.label, {
    x: sx,
    y: s1BarY + 0.15,
    w: 2.1,
    h: 0.2,
    fontSize: 9,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });
  const valFontSize = st.value.length > 12 ? 12 : st.value.length > 8 ? 16 : 22;
  s1.addText(st.value, {
    x: sx,
    y: s1BarY + 0.4,
    w: 2.1,
    h: 0.4,
    fontSize: valFontSize,
    fontFace: "Consolas",
    bold: true,
    color: st.color,
    margin: 0,
  });
});

// ================================================================
// SLIDE 2: Per-Tier Trends
// ================================================================
const s2 = pres.addSlide();
s2.background = { color: C.bgDark };

s2.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: 0,
  w: 10,
  h: 0.06,
  fill: { color: C.cyan },
});

s2.addText("TIER BREAKDOWN", {
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

s2.addText("Per-Tier Score Trends", {
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

// Assessor change warning
if (prev && latest.assessor !== prev.assessor) {
  s2.addText(
    `\u26A0 Assessor changed: ${prev.assessor || "N/A"} \u2192 ${latest.assessor || "N/A"}`,
    {
      x: 0.6,
      y: 0.85,
      w: 8,
      h: 0.22,
      fontSize: 9,
      fontFace: "Calibri",
      color: C.yellow,
      margin: 0,
    }
  );
}

// 5 tier cards
const tierCardW = 1.68;
const tierCardGap = 0.15;
const tierCardY = 1.15;
const tierCardH = 2.8;

const latestTiers = latest.tiers || [];
const prevTiers = prev ? prev.tiers || [] : [];

latestTiers.forEach((t, i) => {
  const cx = 0.6 + i * (tierCardW + tierCardGap);
  const tc = tierColors[t.id] || C.textMuted;
  const prevTier = prevTiers.find((pt) => pt.id === t.id);
  const delta = prevTier ? t.score - prevTier.score : null;

  // Card bg
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx,
    y: tierCardY,
    w: tierCardW,
    h: tierCardH,
    fill: { color: C.bgCard },
    shadow: mkShadow(),
  });

  // Top accent
  s2.addShape(pres.shapes.RECTANGLE, {
    x: cx,
    y: tierCardY,
    w: tierCardW,
    h: 0.05,
    fill: { color: tc },
  });

  // Tier label
  s2.addText(`${t.id} \u2014 ${t.name}`, {
    x: cx + 0.12,
    y: tierCardY + 0.15,
    w: tierCardW - 0.24,
    h: 0.22,
    fontSize: 8,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });

  // Current score
  s2.addText(fmtScore(t.score), {
    x: cx + 0.12,
    y: tierCardY + 0.4,
    w: tierCardW - 0.24,
    h: 0.45,
    fontSize: 30,
    fontFace: "Consolas",
    bold: true,
    color: tc,
    margin: 0,
  });

  // Delta
  if (delta != null) {
    const dc = getDeltaColor(delta);
    s2.addText(`${getDeltaArrow(delta)} ${fmtDelta(delta)}`, {
      x: cx + 0.12,
      y: tierCardY + 0.88,
      w: tierCardW - 0.24,
      h: 0.25,
      fontSize: 12,
      fontFace: "Consolas",
      bold: true,
      color: dc,
      margin: 0,
    });
  } else {
    s2.addText("Baseline", {
      x: cx + 0.12,
      y: tierCardY + 0.88,
      w: tierCardW - 0.24,
      h: 0.25,
      fontSize: 10,
      fontFace: "Consolas",
      color: C.textMuted,
      margin: 0,
    });
  }

  // Status and progression
  s2.addText(t.status, {
    x: cx + 0.12,
    y: tierCardY + 1.2,
    w: tierCardW - 0.24,
    h: 0.2,
    fontSize: 9,
    fontFace: "Calibri",
    color: t.status.includes("Pass") ? C.green : C.red,
    margin: 0,
  });

  s2.addText(t.progression, {
    x: cx + 0.12,
    y: tierCardY + 1.42,
    w: tierCardW - 0.24,
    h: 0.2,
    fontSize: 9,
    fontFace: "Calibri",
    color:
      t.progression === "Complete"
        ? C.green
        : t.progression === "Current"
        ? C.cyan
        : C.textMuted,
    margin: 0,
  });

  // Sparkline — score history for this tier
  if (history.length >= 2) {
    const sparkY = tierCardY + 1.8;
    const sparkW = tierCardW - 0.24;
    const sparkH = 0.7;

    // Background
    s2.addShape(pres.shapes.RECTANGLE, {
      x: cx + 0.12,
      y: sparkY,
      w: sparkW,
      h: sparkH,
      fill: { color: C.bgDark },
    });

    // 3.0 reference line
    const ref3Y = sparkY + sparkH - (3.0 / 5) * sparkH;
    s2.addShape(pres.shapes.LINE, {
      x: cx + 0.12,
      y: ref3Y,
      w: sparkW,
      h: 0,
      line: { color: C.yellow, width: 0.5, dashType: "dash" },
    });

    // Plot bars for each month
    const barW = sparkW / history.length - 0.02;
    history.forEach((entry, mi) => {
      const tierEntry = (entry.tiers || []).find((et) => et.id === t.id);
      const s = tierEntry ? tierEntry.score : 0;
      const barH = Math.max((s / 5) * sparkH, 0.02);
      const bx = cx + 0.12 + mi * (sparkW / history.length) + 0.01;
      const by = sparkY + sparkH - barH;

      s2.addShape(pres.shapes.RECTANGLE, {
        x: bx,
        y: by,
        w: barW,
        h: barH,
        fill: { color: mi === history.length - 1 ? tc : C.bgCardAlt },
      });
    });
  }
});

// Bottom summary
const s2BarY = 4.6;
s2.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: s2BarY,
  w: 10,
  h: 1.025,
  fill: { color: C.bgCard },
});

const improving = latestTiers.filter((t) => {
  const pt = prevTiers.find((p) => p.id === t.id);
  return pt && t.score - pt.score > 0.1;
}).length;
const regressing = latestTiers.filter((t) => {
  const pt = prevTiers.find((p) => p.id === t.id);
  return pt && t.score - pt.score < -0.1;
}).length;
const passing = latestTiers.filter((t) => t.status.includes("Pass")).length;

const s2Stats = [
  { label: "TIERS PASSING", value: `${passing} / ${latestTiers.length}`, color: C.green },
  { label: "TIERS IMPROVING", value: hasMultiple ? String(improving) : "N/A", color: C.green },
  { label: "TIERS REGRESSING", value: hasMultiple ? String(regressing) : "N/A", color: regressing > 0 ? C.red : C.green },
  { label: "PERIOD", value: latest.date, color: C.cyan },
];

s2Stats.forEach((st, i) => {
  const sx = 0.6 + i * 2.35;
  s2.addText(st.label, {
    x: sx,
    y: s2BarY + 0.15,
    w: 2.1,
    h: 0.2,
    fontSize: 9,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });
  s2.addText(st.value, {
    x: sx,
    y: s2BarY + 0.4,
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
// SLIDE 3: Criteria Delta Table
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

s3.addText("CRITERIA CHANGES", {
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

s3.addText(
  hasMultiple ? "Improvements & Focus Areas" : "Current Criteria Scores",
  {
    x: 0.6,
    y: 0.45,
    w: 8,
    h: 0.4,
    fontSize: 22,
    fontFace: "Calibri",
    bold: true,
    color: C.white,
    margin: 0,
  }
);

const latestCriteria = latest.criteria || [];
const prevCriteria = prev ? prev.criteria || [] : [];

// Build delta list (only criteria present in both periods)
const deltas = latestCriteria.map((c) => {
  const pc = prevCriteria.find((p) => p.criterion === c.criterion);
  return {
    criterion: c.criterion,
    category: c.category,
    section: c.section,
    score: c.score,
    prevScore: pc ? pc.score : null,
    delta: pc ? c.score - pc.score : null,
    status: c.status,
    distTo3: Math.abs(c.score - 3.0),
  };
});

// Sort for two lists
const improvements = deltas
  .filter((d) => d.delta != null && d.delta > 0.05)
  .sort((a, b) => b.delta - a.delta)
  .slice(0, 5);

const needsAttention = hasMultiple
  ? deltas
      .filter((d) => d.delta != null && d.delta < -0.05)
      .sort((a, b) => a.delta - b.delta)
      .slice(0, 5)
  : [];

// If not enough regressions, fill with "closest to threshold"
const focusAreas =
  needsAttention.length < 5
    ? deltas
        .filter(
          (d) =>
            d.score < 3.0 &&
            !needsAttention.find((n) => n.criterion === d.criterion)
        )
        .sort((a, b) => b.score - a.score) // closest to 3.0 first
        .slice(0, 5 - needsAttention.length)
    : [];

const attentionList = [...needsAttention, ...focusAreas];

// Table rendering helper
function renderTable(slide, title, titleColor, items, startY, showDelta) {
  slide.addText(title, {
    x: 0.6,
    y: startY,
    w: 8,
    h: 0.28,
    fontSize: 10,
    fontFace: "Consolas",
    bold: true,
    color: titleColor,
    charSpacing: 1,
    margin: 0,
  });

  // Header
  const hdrY = startY + 0.32;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5,
    y: hdrY,
    w: 9.0,
    h: 0.26,
    fill: { color: C.bgCard },
  });

  const cols = showDelta
    ? [
        { x: 0.6, w: 3.2, label: "CRITERION" },
        { x: 3.8, w: 1.6, label: "CATEGORY" },
        { x: 5.4, w: 1.0, label: "PREVIOUS" },
        { x: 6.4, w: 1.0, label: "CURRENT" },
        { x: 7.4, w: 0.8, label: "DELTA" },
        { x: 8.2, w: 1.2, label: "STATUS" },
      ]
    : [
        { x: 0.6, w: 3.5, label: "CRITERION" },
        { x: 4.1, w: 2.0, label: "CATEGORY" },
        { x: 6.1, w: 1.2, label: "SCORE" },
        { x: 7.3, w: 1.0, label: "LEVEL" },
        { x: 8.3, w: 1.2, label: "STATUS" },
      ];

  cols.forEach((col) => {
    slide.addText(col.label, {
      x: col.x,
      y: hdrY + 0.03,
      w: col.w,
      h: 0.2,
      fontSize: 7,
      fontFace: "Consolas",
      color: C.textMuted,
      charSpacing: 1,
      margin: 0,
    });
  });

  if (items.length === 0) {
    slide.addText(
      showDelta ? "No changes detected." : "No criteria below threshold.",
      {
        x: 0.6,
        y: hdrY + 0.35,
        w: 8,
        h: 0.25,
        fontSize: 10,
        fontFace: "Calibri",
        color: C.textMuted,
        margin: 0,
      }
    );
    return hdrY + 0.7;
  }

  const rowH = 0.24;
  items.forEach((item, idx) => {
    const ry = hdrY + 0.3 + idx * (rowH + 0.02);

    if (idx % 2 === 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.5,
        y: ry - 0.01,
        w: 9.0,
        h: rowH + 0.01,
        fill: { color: C.bgCard },
      });
    }

    if (showDelta) {
      // Criterion
      slide.addText(item.criterion, {
        x: 0.6, y: ry, w: 3.2, h: rowH,
        fontSize: 9, fontFace: "Calibri", color: C.textPrimary, margin: 0,
      });
      // Category
      slide.addText(item.category, {
        x: 3.8, y: ry, w: 1.6, h: rowH,
        fontSize: 9, fontFace: "Calibri", color: C.textSecondary, margin: 0,
      });
      // Previous
      slide.addText(item.prevScore != null ? fmtScore(item.prevScore) : "N/A", {
        x: 5.4, y: ry, w: 1.0, h: rowH,
        fontSize: 9, fontFace: "Consolas", color: C.textMuted, align: "center", margin: 0,
      });
      // Current
      slide.addText(fmtScore(item.score), {
        x: 6.4, y: ry, w: 1.0, h: rowH,
        fontSize: 9, fontFace: "Consolas", bold: true, color: getScoreColor(item.score), align: "center", margin: 0,
      });
      // Delta
      const dc = item.delta != null ? getDeltaColor(item.delta) : C.textMuted;
      slide.addText(item.delta != null ? `${getDeltaArrow(item.delta)} ${fmtDelta(item.delta)}` : "N/A", {
        x: 7.4, y: ry, w: 0.8, h: rowH,
        fontSize: 9, fontFace: "Consolas", bold: true, color: dc, align: "center", margin: 0,
      });
      // Status
      const isBelow = item.status.includes("Below");
      slide.addText(isBelow ? "\u2717" : "\u2713", {
        x: 8.2, y: ry, w: 1.2, h: rowH,
        fontSize: 11, fontFace: "Calibri", bold: true, color: isBelow ? C.red : C.green, align: "center", margin: 0,
      });
    } else {
      slide.addText(item.criterion, {
        x: 0.6, y: ry, w: 3.5, h: rowH,
        fontSize: 9, fontFace: "Calibri", color: C.textPrimary, margin: 0,
      });
      slide.addText(item.category, {
        x: 4.1, y: ry, w: 2.0, h: rowH,
        fontSize: 9, fontFace: "Calibri", color: C.textSecondary, margin: 0,
      });
      slide.addText(fmtScore(item.score), {
        x: 6.1, y: ry, w: 1.2, h: rowH,
        fontSize: 9, fontFace: "Consolas", bold: true, color: getScoreColor(item.score), align: "center", margin: 0,
      });
      slide.addText(item.level || "N/A", {
        x: 7.3, y: ry, w: 1.0, h: rowH,
        fontSize: 8, fontFace: "Consolas", color: C.textSecondary, align: "center", margin: 0,
      });
      const isBelow = item.status.includes("Below");
      slide.addText(isBelow ? "\u2717" : "\u2713", {
        x: 8.3, y: ry, w: 1.2, h: rowH,
        fontSize: 11, fontFace: "Calibri", bold: true, color: isBelow ? C.red : C.green, align: "center", margin: 0,
      });
    }
  });

  return hdrY + 0.3 + items.length * (rowH + 0.02) + 0.15;
}

if (hasMultiple) {
  const afterImprove = renderTable(
    s3,
    "\u25B2 BIGGEST IMPROVEMENTS",
    C.green,
    improvements,
    0.95,
    true
  );
  renderTable(
    s3,
    "\u25BC NEEDS ATTENTION",
    C.red,
    attentionList,
    afterImprove + 0.1,
    true
  );
} else {
  // Single entry — show all below-threshold criteria
  const belowThreshold = deltas
    .filter((d) => d.score < 3.0)
    .sort((a, b) => a.score - b.score);
  renderTable(s3, "CRITERIA BELOW THRESHOLD (< 3.0)", C.red, belowThreshold, 0.95, false);
}

// Bottom summary stats
const s3BarY = 4.6;
s3.addShape(pres.shapes.RECTANGLE, {
  x: 0,
  y: s3BarY,
  w: 10,
  h: 1.025,
  fill: { color: C.bgCard },
});

const totalImproved = deltas.filter((d) => d.delta != null && d.delta > 0.05).length;
const totalRegressed = deltas.filter((d) => d.delta != null && d.delta < -0.05).length;
const totalUnchanged = deltas.filter(
  (d) => d.delta != null && Math.abs(d.delta) <= 0.05
).length;
const belowCount = deltas.filter((d) => d.score < 3.0).length;

const s3Stats = hasMultiple
  ? [
      { label: "IMPROVED", value: String(totalImproved), color: C.green },
      { label: "REGRESSED", value: String(totalRegressed), color: totalRegressed > 0 ? C.red : C.green },
      { label: "UNCHANGED", value: String(totalUnchanged), color: C.textSecondary },
      { label: "BELOW 3.0", value: String(belowCount), color: belowCount > 0 ? C.orange : C.green },
    ]
  : [
      { label: "TOTAL CRITERIA", value: String(deltas.length), color: C.cyan },
      { label: "PASSING", value: String(deltas.length - belowCount), color: C.green },
      { label: "BELOW 3.0", value: String(belowCount), color: belowCount > 0 ? C.orange : C.green },
      { label: "PERIOD", value: latest.date, color: C.cyan },
    ];

s3Stats.forEach((st, i) => {
  const sx = 0.6 + i * 2.35;
  s3.addText(st.label, {
    x: sx,
    y: s3BarY + 0.15,
    w: 2.1,
    h: 0.2,
    fontSize: 9,
    fontFace: "Consolas",
    color: C.textMuted,
    margin: 0,
  });
  s3.addText(st.value, {
    x: sx,
    y: s3BarY + 0.4,
    w: 2.1,
    h: 0.4,
    fontSize: 22,
    fontFace: "Consolas",
    bold: true,
    color: st.color,
    margin: 0,
  });
});

// ── Write file ──
pres.writeFile({ fileName: outPath }).then(() => {
  console.log(`Trend report generated: ${outPath}`);
  console.log(`  Periods: ${history.length} (${history[0].date} \u2014 ${latest.date})`);
  if (hasMultiple) {
    console.log(`  Criteria: ${totalImproved} improved, ${totalRegressed} regressed, ${totalUnchanged} unchanged`);
  }
});
