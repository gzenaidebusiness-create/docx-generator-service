const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageBreak, LevelFormat
} = require('docx');

// ── COLOURS ──────────────────────────────────────────────────
const NAVY   = "0F2044";
const TEAL   = "0D9488";
const AMBER  = "D97706";
const LIGHT  = "EFF6FF";
const WHITE  = "FFFFFF";
const GREY   = "F1F5F9";
const MID    = "475569";

// ── PAGE SETUP ────────────────────────────────────────────────
// Landscape A4: 11906 x 16838 DXA (docx-js swaps internally)
const LANDSCAPE_W = 11906;
const LANDSCAPE_H = 16838;
const MARGIN = 720; // 0.5 inch
const CONTENT_W = LANDSCAPE_H - (MARGIN * 2); // ~15398 DXA landscape

const BORDER = { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

function cell(text, opts = {}) {
  const { bold = false, color = "1E293B", bg = WHITE, w = 1000, shade = false, align = AlignmentType.LEFT, size = 18 } = opts;
  return new TableCell({
    borders: BORDERS,
    width: { size: w, type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.TOP,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text || ''), bold, color, font: "Arial", size })]
    })]
  });
}

function hdrCell(text, w) {
  return cell(text, { bold: true, color: WHITE, bg: NAVY, w, size: 18 });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: TEAL, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: NAVY })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: TEAL })]
  });
}

function para(text, opts = {}) {
  const { bold = false, color = "1E293B", size = 20, before = 60, after = 60, italic = false } = opts;
  return new Paragraph({
    spacing: { before, after },
    children: [new TextRun({ text: String(text || ''), bold, color, font: "Arial", size, italic })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: String(text || ''), font: "Arial", size: 18, color: "1E293B" })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun("")] });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function arr(v) { return Array.isArray(v) ? v : (v ? [v] : []); }
function str(v) { return v ? String(v) : '[NOT AVAILABLE]'; }

// ── COVER PAGE ────────────────────────────────────────────────
function buildCoverPage(categoryName, runId, runLabel, totalReviews, totalProducts) {
  const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  return [
    spacer(), spacer(), spacer(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: "CONSUMER RESEARCH ANALYSIS", font: "Arial", size: 48, bold: true, color: NAVY, allCaps: true })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 4 } },
      children: [new TextRun({ text: `${categoryName} — Prompt 0 Full Segmentation Report`, font: "Arial", size: 32, color: TEAL })]
    }),
    spacer(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 40 },
      children: [new TextRun({ text: `Source: Amazon Reviews Dataset + Product Page Details`, font: "Arial", size: 22, color: MID })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: `${totalReviews || '—'} Reviews  |  ${totalProducts || '—'} Products`, font: "Arial", size: 22, color: MID })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: `Run: ${runLabel || runId}  |  Generated: ${today}`, font: "Arial", size: 20, color: MID })]
    }),
    spacer(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 40 },
      children: [new TextRun({ text: "Five Tasks:", font: "Arial", size: 20, bold: true, color: NAVY })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: "Competitive Profile · Supply-Side Fact Sheet · Buyer Segments · Theme Stratification · Longitudinal Flag", font: "Arial", size: 20, color: MID })]
    }),
    pageBreak()
  ];
}

// ── TASK 1: COMPETITIVE PROFILE ───────────────────────────────
function buildTask1(task1) {
  const products = arr(task1);
  if (!products.length) return [h1("Task 1 — Product-Level Competitive Profile"), para("No data available.")];

  // Column widths summing to CONTENT_W ~15398
  const cols = [1200, 2200, 1400, 1400, 2200, 2200, 2400, 2400]; // 15400
  const headers = ["ASIN", "Product", "Ratings", "Format", "Top 3 Positive Themes", "Top 3 Negative Themes", "POSITIVE Job Signal →", "NEGATIVE Job Signal →"];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => hdrCell(h, cols[i]))
  });

  const dataRows = products.map((p, idx) => {
    const ratings = `n = ${p.total_reviews || p.totalReviews || '?'}\nAvg: ${p.avg_rating || p.avgRating || '?'}★\n1-2★: ${p.pct_low_star || p.pctLowStar || '?'}%\n4-5★: ${p.pct_high_star || p.pctHighStar || '?'}%`;
    const posThemes = arr(p.top3_positive_themes || p.topPositiveThemes).map((t, i) => `${i+1}. ${t}`).join('\n');
    const negThemes = arr(p.top3_negative_themes || p.topNegativeThemes).map((t, i) => `${i+1}. ${t}`).join('\n');
    const bg = idx % 2 === 0 ? WHITE : GREY;
    return new TableRow({
      children: [
        cell(p.asin, { w: cols[0], bg }),
        cell(p.product_name || p.productName || p.asin, { w: cols[1], bg, bold: true }),
        cell(ratings, { w: cols[2], bg }),
        cell(p.dominant_format || p.dominantFormat || '[NOT AVAILABLE]', { w: cols[3], bg }),
        cell(posThemes, { w: cols[4], bg }),
        cell(negThemes, { w: cols[5], bg }),
        cell(p.positive_job_signal || p.positiveJobSignal || '[NOT AVAILABLE]', { w: cols[6], bg, color: "065F46" }),
        cell(p.negative_job_signal || p.negativeJobSignal || '[NOT AVAILABLE]', { w: cols[7], bg, color: "991B1B" }),
      ]
    });
  });

  return [
    h1("Task 1 — Product-Level Competitive Profile"),
    para("Each row represents one ASIN. Ratings derived from the combined review dataset. Positive and negative job signals are in separate columns — direction is explicit.", { color: MID, size: 18 }),
    spacer(),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: cols,
      rows: [headerRow, ...dataRows]
    }),
    spacer(),
    pageBreak()
  ];
}

// ── TASK 2: SUPPLY-SIDE ───────────────────────────────────────
function buildTask2(task2) {
  const products = arr(task2);
  if (!products.length) return [h1("Task 2 — Supply-Side Product Fact Sheet"), para("No data available.")];

  const cols = [2000, 2500, 2200, 2000, 1600, 1600, 3500]; // ~15400
  const headers = ["Product", "Primary Active Ingredients", "Certifications / Endorsements", "Regulatory Classification", "Price per Dose", "Species Positioning", "Distribution Channel Signals"];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => hdrCell(h, cols[i]))
  });

  const dataRows = products.map((p, idx) => {
    const bg = idx % 2 === 0 ? WHITE : GREY;
    return new TableRow({
      children: [
        cell(p.product_name || p.productName || p.asin, { w: cols[0], bg, bold: true }),
        cell(str(p.primary_ingredients || p.primaryIngredients), { w: cols[1], bg }),
        cell(str(p.certifications), { w: cols[2], bg }),
        cell(str(p.regulatory_classification || p.regulatoryClassification), { w: cols[3], bg }),
        cell(str(p.price_per_dose || p.pricePerDose), { w: cols[4], bg }),
        cell(str(p.species_positioning || p.speciesPositioning), { w: cols[5], bg }),
        cell(str(p.distribution_channels || p.distributionChannels), { w: cols[6], bg }),
      ]
    });
  });

  return [
    h1("Task 2 — Supply-Side Product Fact Sheet"),
    para("Primary active ingredients, certifications, regulatory classification, pricing, species positioning and channel data extracted or inferred from listings and review text.", { color: MID, size: 18 }),
    para("Purpose note: This table travels through all subsequent analysis prompts. Innovation concepts generated in later stages can be cross-referenced against these constraints.", { color: AMBER, size: 17, italic: true }),
    spacer(),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: cols,
      rows: [headerRow, ...dataRows]
    }),
    spacer(),
    pageBreak()
  ];
}

// ── TASK 3: BUYER SEGMENTS ────────────────────────────────────
function buildTask3(task3) {
  const segments = arr(task3);
  if (!segments.length) return [h1("Task 3 — Buyer Segment Identification"), para("No data available.")];

  const cols = [2200, 2400, 2400, 2000, 1200, 5200]; // ~15400
  const headers = ["Segment Name", "Hiring Trigger", "Subject Type & Life Stage", "Dominant Emotional Register", "Est. % of Reviews", "Representative Quotes (paraphrased, ≤15 words)"];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => hdrCell(h, cols[i]))
  });

  const dataRows = segments.map((s, idx) => {
    const quotes = arr(s.representative_quotes || s.representativeQuotes).map(q => `"${q}"`).join('\n\n');
    const bg = idx % 2 === 0 ? WHITE : GREY;
    return new TableRow({
      children: [
        cell(s.segment_name || s.segmentName, { w: cols[0], bg, bold: true, color: NAVY }),
        cell(str(s.hiring_trigger || s.hiringTrigger), { w: cols[1], bg }),
        cell(str(s.subject_type_and_life_stage || s.subjectTypeAndLifeStage), { w: cols[2], bg }),
        cell(str(s.emotional_register || s.dominantEmotionalRegister || s.emotionalRegister), { w: cols[3], bg, italic: true }),
        cell(str(s.estimated_pct_of_reviews || s.estimatedPct || s.estimatedPctOfReviews), { w: cols[4], bg, bold: true, color: TEAL, align: AlignmentType.CENTER }),
        cell(quotes || '[No quotes available]', { w: cols[5], bg, italic: true, color: "374151" }),
      ]
    });
  });

  return [
    h1("Task 3 — Buyer Segment Identification"),
    para(`${segments.length} buyer archetypes identified from review text. Segments defined by hiring trigger, subject type, emotional register and estimated share of reviews.`, { color: MID, size: 18 }),
    spacer(),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: cols,
      rows: [headerRow, ...dataRows]
    }),
    spacer(),
    pageBreak()
  ];
}

// ── TASK 4: THEME STRATIFICATION ─────────────────────────────
function buildTask4(task4) {
  const themes = arr(task4);
  if (!themes.length) return [h1("Task 4 — Star-Rating Stratification by Theme"), para("No data available.")];

  const cols = [5000, 2000, 2000, 6400]; // ~15400
  const headers = ["Theme", "% in 1–2★ Reviews", "% in 4–5★ Reviews", "Net Satisfaction Signal"];

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => hdrCell(h, cols[i]))
  });

  const dataRows = themes.map((t, idx) => {
    const low = Number(t.pct_in_low_star || t.pctInLowStar || 0);
    const high = Number(t.pct_in_high_star || t.pctInHighStar || 0);
    const lowColor = low > 50 ? "991B1B" : "374151";
    const highColor = high > 50 ? "065F46" : "374151";
    const bg = idx % 2 === 0 ? WHITE : GREY;
    return new TableRow({
      children: [
        cell(t.theme, { w: cols[0], bg, bold: true }),
        cell(`${low}%`, { w: cols[1], bg, color: lowColor, bold: low > 50, align: AlignmentType.CENTER }),
        cell(`${high}%`, { w: cols[2], bg, color: highColor, bold: high > 50, align: AlignmentType.CENTER }),
        cell(str(t.net_satisfaction_signal || t.netSatisfactionSignal), { w: cols[3], bg }),
      ]
    });
  });

  return [
    h1("Task 4 — Star-Rating Stratification by Theme"),
    para("Theme mention frequencies estimated through systematic reading of all 1–2★ and 4–5★ reviews. Percentages = share of low/high-star reviews mentioning the theme.", { color: MID, size: 18 }),
    spacer(),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: cols,
      rows: [headerRow, ...dataRows]
    }),
    spacer(),
    pageBreak()
  ];
}

// ── TASK 5: LONGITUDINAL FLAG ─────────────────────────────────
function buildTask5(task5) {
  if (!task5) return [h1("Task 5 — Longitudinal Flag"), para("No data available.")];

  const cohorts = ['early', 'mid', 'recent'];
  const labels = ['Cohort 1: EARLY', 'Cohort 2: MID', 'Cohort 3: RECENT'];
  const elements = [h1("Task 5 — Longitudinal Flag")];

  cohorts.forEach((key, i) => {
    const c = task5[key] || task5[`cohort_${key}`] || task5[`${key}Cohort`];
    if (!c) return;
    elements.push(h2(labels[i]));
    if (c.date_range || c.dateRange) {
      const dr = c.date_range || c.dateRange;
      elements.push(para(`Date range: ${dr.from || dr.earliest || '?'} – ${dr.to || dr.latest || '?'}  |  Reviews: ${c.review_count || c.count || c.reviewCount || '?'}`, { color: MID, size: 18, italic: true }));
    }
    if (arr(c.top3_complaints || c.top3Complaints).length) {
      elements.push(para("Top 3 Complaints:", { bold: true, color: "991B1B", size: 19 }));
      arr(c.top3_complaints || c.top3Complaints).forEach(t => elements.push(bullet(t)));
    }
    if (arr(c.top3_praise || c.top3Praise).length) {
      elements.push(para("Top 3 Praise:", { bold: true, color: "065F46", size: 19 }));
      arr(c.top3_praise || c.top3Praise).forEach(t => elements.push(bullet(t)));
    }
    elements.push(spacer());
  });

  // Cross-cohort signals
  const emerging = arr(task5.emerging_signals || task5.emergingSignals);
  const resolving = arr(task5.resolving_signals || task5.resolvingSignals);

  if (emerging.length || resolving.length) {
    elements.push(h2("Cross-Cohort Signal Summary"));
    if (emerging.length) {
      elements.push(para("ESCALATING / EMERGING", { bold: true, color: AMBER, size: 20 }));
      emerging.forEach(s => elements.push(bullet(s)));
    }
    if (resolving.length) {
      elements.push(para("RESOLVING / DECLINING", { bold: true, color: MID, size: 20 }));
      resolving.forEach(s => elements.push(bullet(s)));
    }
  }

  // Temporal warning
  if (task5.temporal_compression_warning || task5.temporalCompressionWarning) {
    elements.push(spacer());
    elements.push(para("⚠ TEMPORAL COMPRESSION WARNING", { bold: true, color: AMBER, size: 20 }));
    elements.push(para(task5.temporal_compression_warning || task5.temporalCompressionWarning, { color: MID, size: 18, italic: true }));
  }

  return elements;
}

// ── MAIN GENERATOR ────────────────────────────────────────────
async function generateP0Docx(p0Output, categoryName, runId, runLabel, totalReviews, totalProducts) {
  const today = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

  const children = [
    ...buildCoverPage(categoryName, runId, runLabel, totalReviews, totalProducts),
    ...buildTask1(p0Output.task1_competitive_profile),
    ...buildTask2(p0Output.task2_supply_side_profile),
    ...buildTask3(p0Output.task3_buyer_segments),
    ...buildTask4(p0Output.task4_theme_stratification),
    ...buildTask5(p0Output.task5_longitudinal_flags),
    spacer(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 80 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } },
      children: [new TextRun({ text: `End of Prompt 0 Report  |  ${categoryName}  |  ${today}  |  ${totalReviews || '—'} reviews across ${totalProducts || '—'} products`, font: "Arial", size: 16, color: MID, italic: true })]
    })
  ];

  const doc = new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "▪",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }]
    },
    styles: {
      default: {
        document: { run: { font: "Arial", size: 20 } }
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "Arial", color: NAVY },
          paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 }
        },
        {
          id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "Arial", color: TEAL },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          size: {
            width: LANDSCAPE_W,
            height: LANDSCAPE_H,
            orientation: "landscape"
          },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
        }
      },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}

module.exports = { generateP0Docx };
