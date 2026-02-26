/**
 * AI Product Manager - PowerPoint Deck Generator
 * Creates executive stakeholder presentations from scored capability data.
 */
const pptxgen = require("pptxgenjs");
const fs = require("fs");

// Color palette — Ocean Gradient (professional, trust-building)
const COLORS = {
  primary: "065A82",     // deep blue
  secondary: "1C7293",   // teal
  accent: "21295C",      // midnight
  white: "FFFFFF",
  offWhite: "F8FAFC",
  lightGray: "E2E8F0",
  darkText: "1E293B",
  mutedText: "64748B",
  high: "059669",        // green for HIGH
  medium: "D97706",      // amber for MEDIUM
  low: "DC2626",         // red for LOW
  chartGreen: "10B981",
  chartBlue: "3B82F6",
  chartAmber: "F59E0B",
};

const FONTS = { header: "Georgia", body: "Calibri" };

function makeShadow() {
  return { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 };
}

function createDeck(data, outputPath) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "AI Product Manager";
  pres.title = "Capability Prioritization — Executive Brief";

  addTitleSlide(pres, data);
  addExecutiveSummarySlide(pres, data);
  addSurveyInsightsSlide(pres, data);
  addPriorityMatrixSlide(pres, data);
  addTop5Slide(pres, data);
  addCategoryBreakdownSlide(pres, data);
  addDemandHeatmapSlide(pres, data);
  addRecommendationsSlide(pres, data);
  addNextStepsSlide(pres, data);

  pres.writeFile({ fileName: outputPath }).then(() => {
    console.log(`Deck saved to ${outputPath}`);
  });
}

function addTitleSlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.accent };

  // Large decorative shape
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.2, fill: { color: COLORS.primary, transparency: 40 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.4, w: 10, h: 1.225, fill: { color: COLORS.primary, transparency: 40 }
  });

  slide.addText("AI PRODUCT MANAGER", {
    x: 0.8, y: 1.5, w: 8.4, h: 0.6, fontSize: 14, fontFace: FONTS.body,
    color: COLORS.secondary, charSpacing: 6, bold: true
  });
  slide.addText("Capability Prioritization\nExecutive Brief", {
    x: 0.8, y: 2.1, w: 8.4, h: 1.6, fontSize: 36, fontFace: FONTS.header,
    color: COLORS.white, bold: true
  });
  slide.addText(`${data.stats.total_capabilities} Capabilities Scored | ${data.survey.total_respondents} Enterprise Customers Surveyed`, {
    x: 0.8, y: 3.8, w: 8.4, h: 0.4, fontSize: 13, fontFace: FONTS.body, color: COLORS.lightGray
  });
  slide.addText(new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" }), {
    x: 0.8, y: 4.6, w: 4, h: 0.3, fontSize: 11, fontFace: FONTS.body, color: COLORS.mutedText
  });
}

function addExecutiveSummarySlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.offWhite };

  slide.addText("Executive Summary", {
    x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 28, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });

  // Stat cards
  const stats = [
    { label: "HIGH PRIORITY", value: String(data.stats.high_priority), color: COLORS.high },
    { label: "MEDIUM PRIORITY", value: String(data.stats.medium_priority), color: COLORS.medium },
    { label: "LOW PRIORITY", value: String(data.stats.low_priority), color: COLORS.low },
    { label: "AVG SCORE", value: String(data.stats.avg_score), color: COLORS.primary },
  ];

  stats.forEach((stat, i) => {
    const x = 0.8 + i * 2.2;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.15, w: 2.0, h: 1.4, fill: { color: COLORS.white },
      shadow: makeShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.15, w: 2.0, h: 0.06, fill: { color: stat.color }
    });
    slide.addText(stat.value, {
      x, y: 1.35, w: 2.0, h: 0.8, fontSize: 36, fontFace: FONTS.header,
      color: stat.color, align: "center", valign: "middle", bold: true
    });
    slide.addText(stat.label, {
      x, y: 2.1, w: 2.0, h: 0.35, fontSize: 9, fontFace: FONTS.body,
      color: COLORS.mutedText, align: "center", charSpacing: 2
    });
  });

  // Top 3 capabilities
  slide.addText("Top 3 Capabilities", {
    x: 0.8, y: 2.85, w: 8.4, h: 0.4, fontSize: 16, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });

  data.top5.slice(0, 3).forEach((cap, i) => {
    const y = 3.35 + i * 0.7;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y, w: 8.4, h: 0.6, fill: { color: COLORS.white }, shadow: makeShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y, w: 0.06, h: 0.6, fill: { color: COLORS.primary }
    });
    slide.addText(`#${cap.rank}`, {
      x: 1.0, y, w: 0.6, h: 0.6, fontSize: 16, fontFace: FONTS.header,
      color: COLORS.primary, bold: true, valign: "middle"
    });
    slide.addText(cap.name, {
      x: 1.6, y, w: 5.5, h: 0.6, fontSize: 14, fontFace: FONTS.body,
      color: COLORS.darkText, valign: "middle"
    });
    slide.addText(cap.score.toFixed(2), {
      x: 7.5, y, w: 1.0, h: 0.6, fontSize: 18, fontFace: FONTS.header,
      color: COLORS.primary, bold: true, align: "center", valign: "middle"
    });
    const prioColor = cap.priority === "HIGH" ? COLORS.high : cap.priority === "MEDIUM" ? COLORS.medium : COLORS.low;
    slide.addText(cap.priority, {
      x: 8.5, y: y + 0.12, w: 0.65, h: 0.35, fontSize: 8, fontFace: FONTS.body,
      color: COLORS.white, fill: { color: prioColor }, align: "center", valign: "middle",
      bold: true, charSpacing: 1
    });
  });
}

function addSurveyInsightsSlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.offWhite };

  slide.addText("Customer Survey Insights", {
    x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 28, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });
  slide.addText(`${data.survey.total_respondents} enterprise respondents across diverse industries`, {
    x: 0.8, y: 0.85, w: 8.4, h: 0.3, fontSize: 12, fontFace: FONTS.body, color: COLORS.mutedText
  });

  // Companies list
  slide.addText("Participating Companies", {
    x: 0.8, y: 1.35, w: 4, h: 0.35, fontSize: 14, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });

  const companies = data.survey.companies || [];
  companies.slice(0, 9).forEach((company, i) => {
    const col = Math.floor(i / 5);
    const row = i % 5;
    slide.addText([
      { text: "\u2022 ", options: { color: COLORS.primary, bold: true } },
      { text: company, options: { color: COLORS.darkText } }
    ], {
      x: 0.8 + col * 2.5, y: 1.75 + row * 0.35, w: 2.3, h: 0.3,
      fontSize: 11, fontFace: FONTS.body
    });
  });

  // Top demand chart
  if (data.demand && data.demand.length > 0) {
    const topDemand = data.demand.slice(0, 5);
    const chartData = [{
      name: "Demand",
      labels: topDemand.map(d => d.capability),
      values: topDemand.map(d => d.count)
    }];

    slide.addChart(pres.charts.BAR, chartData, {
      x: 5.5, y: 1.35, w: 4.0, h: 3.8, barDir: "bar",
      showTitle: true, title: "Top Capability Demand",
      titleColor: COLORS.accent, titleFontFace: FONTS.header, titleFontSize: 12,
      chartColors: [COLORS.primary],
      catAxisLabelColor: COLORS.darkText, catAxisLabelFontSize: 9,
      valAxisLabelColor: COLORS.mutedText, valAxisLabelFontSize: 8,
      valGridLine: { color: COLORS.lightGray, size: 0.5 },
      catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: COLORS.darkText,
      chartArea: { fill: { color: COLORS.white }, roundedCorners: true },
    });
  }
}

function addPriorityMatrixSlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.offWhite };

  slide.addText("Priority Matrix", {
    x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 28, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });

  // Table
  const header = [
    { text: "Rank", options: { fill: { color: COLORS.primary }, color: COLORS.white, bold: true, fontSize: 9 } },
    { text: "Capability", options: { fill: { color: COLORS.primary }, color: COLORS.white, bold: true, fontSize: 9 } },
    { text: "Tier", options: { fill: { color: COLORS.primary }, color: COLORS.white, bold: true, fontSize: 9 } },
    { text: "Priority", options: { fill: { color: COLORS.primary }, color: COLORS.white, bold: true, fontSize: 9 } },
    { text: "Score", options: { fill: { color: COLORS.primary }, color: COLORS.white, bold: true, fontSize: 9 } },
    { text: "Customer", options: { fill: { color: COLORS.primary }, color: COLORS.white, bold: true, fontSize: 9 } },
    { text: "Business", options: { fill: { color: COLORS.primary }, color: COLORS.white, bold: true, fontSize: 9 } },
  ];

  const rows = [header];
  data.all_capabilities.slice(0, 15).forEach((cap, i) => {
    const bgColor = i % 2 === 0 ? COLORS.white : COLORS.offWhite;
    const prioColor = cap.priority === "HIGH" ? COLORS.high : cap.priority === "MEDIUM" ? COLORS.medium : COLORS.low;
    rows.push([
      { text: String(cap.rank), options: { fill: { color: bgColor }, fontSize: 8, align: "center" } },
      { text: cap.name.substring(0, 45), options: { fill: { color: bgColor }, fontSize: 8 } },
      { text: cap.tier, options: { fill: { color: bgColor }, fontSize: 8 } },
      { text: cap.priority, options: { fill: { color: bgColor }, color: prioColor, bold: true, fontSize: 8 } },
      { text: cap.score.toFixed(1), options: { fill: { color: bgColor }, fontSize: 8, align: "center" } },
      { text: cap.customer_impact.toFixed(1), options: { fill: { color: bgColor }, fontSize: 8, align: "center" } },
      { text: cap.business_impact.toFixed(1), options: { fill: { color: bgColor }, fontSize: 8, align: "center" } },
    ]);
  });

  slide.addTable(rows, {
    x: 0.3, y: 1.0, w: 9.4,
    colW: [0.45, 3.6, 1.0, 0.75, 0.65, 0.85, 0.85],
    border: { pt: 0.5, color: COLORS.lightGray },
    autoPage: false,
  });
}

function addTop5Slide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.offWhite };

  slide.addText("Top 5 Capabilities — Detail", {
    x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 28, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });

  data.top5.forEach((cap, i) => {
    const y = 1.05 + i * 0.88;
    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 9.0, h: 0.78, fill: { color: COLORS.white }, shadow: makeShadow()
    });
    // Left accent
    const prioColor = cap.priority === "HIGH" ? COLORS.high : cap.priority === "MEDIUM" ? COLORS.medium : COLORS.low;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y, w: 0.06, h: 0.78, fill: { color: prioColor }
    });
    // Rank circle
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: y + 0.14, w: 0.5, h: 0.5, fill: { color: COLORS.primary }
    });
    slide.addText(String(cap.rank), {
      x: 0.7, y: y + 0.14, w: 0.5, h: 0.5, fontSize: 16, fontFace: FONTS.header,
      color: COLORS.white, bold: true, align: "center", valign: "middle"
    });
    // Name and tier
    slide.addText(cap.name, {
      x: 1.35, y: y + 0.05, w: 5.0, h: 0.35, fontSize: 13, fontFace: FONTS.body,
      color: COLORS.darkText, bold: true, valign: "middle"
    });
    slide.addText(`${cap.tier} | Customer: ${cap.customer_impact.toFixed(1)} | Business: ${cap.business_impact.toFixed(1)} | Cost: ${cap.cost_to_implement.toFixed(1)}`, {
      x: 1.35, y: y + 0.38, w: 5.5, h: 0.3, fontSize: 9, fontFace: FONTS.body,
      color: COLORS.mutedText, valign: "middle"
    });
    // Score
    slide.addText(cap.score.toFixed(2), {
      x: 7.8, y: y + 0.05, w: 1.2, h: 0.45, fontSize: 22, fontFace: FONTS.header,
      color: COLORS.primary, bold: true, align: "center", valign: "middle"
    });
    slide.addText(cap.priority, {
      x: 8.1, y: y + 0.5, w: 0.65, h: 0.22, fontSize: 7, fontFace: FONTS.body,
      color: COLORS.white, fill: { color: prioColor }, align: "center", valign: "middle",
      bold: true, charSpacing: 1
    });
  });
}

function addCategoryBreakdownSlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.offWhite };

  slide.addText("Category Breakdown", {
    x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 28, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });

  if (data.categories && data.categories.length > 0) {
    const chartData = [{
      name: "Avg Score",
      labels: data.categories.map(c => c.name),
      values: data.categories.map(c => c.avg_score)
    }];

    slide.addChart(pres.charts.BAR, chartData, {
      x: 0.5, y: 1.0, w: 5.5, h: 4.0, barDir: "col",
      showTitle: false,
      chartColors: [COLORS.primary, COLORS.secondary, COLORS.accent, COLORS.chartGreen],
      catAxisLabelColor: COLORS.darkText, catAxisLabelFontSize: 10,
      valAxisLabelColor: COLORS.mutedText,
      valGridLine: { color: COLORS.lightGray, size: 0.5 },
      catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: COLORS.darkText,
      chartArea: { fill: { color: COLORS.white }, roundedCorners: true },
    });

    // Category cards on right
    data.categories.forEach((cat, i) => {
      const y = 1.0 + i * 1.1;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 6.3, y, w: 3.2, h: 0.95, fill: { color: COLORS.white }, shadow: makeShadow()
      });
      slide.addText(cat.name, {
        x: 6.5, y: y + 0.05, w: 2.8, h: 0.3, fontSize: 11, fontFace: FONTS.body,
        color: COLORS.accent, bold: true
      });
      slide.addText(`${cat.count} capabilities | Avg: ${cat.avg_score.toFixed(1)} | High: ${cat.high_count}`, {
        x: 6.5, y: y + 0.35, w: 2.8, h: 0.25, fontSize: 9, fontFace: FONTS.body,
        color: COLORS.mutedText
      });
    });
  }
}

function addDemandHeatmapSlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.offWhite };

  slide.addText("Customer Demand Heatmap", {
    x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 28, fontFace: FONTS.header,
    color: COLORS.accent, bold: true, margin: 0
  });
  slide.addText("Capabilities ranked by number of enterprise customers requesting them", {
    x: 0.8, y: 0.85, w: 8.4, h: 0.3, fontSize: 12, fontFace: FONTS.body, color: COLORS.mutedText
  });

  if (data.demand && data.demand.length > 0) {
    data.demand.forEach((d, i) => {
      const y = 1.35 + i * 0.42;
      const barWidth = Math.max(0.5, (d.count / data.survey.total_respondents) * 6.5);
      const intensity = Math.min(1, d.count / data.survey.total_respondents);
      const barColor = intensity > 0.7 ? COLORS.high : intensity > 0.4 ? COLORS.chartAmber : COLORS.primary;

      slide.addText(d.capability, {
        x: 0.5, y, w: 2.5, h: 0.35, fontSize: 10, fontFace: FONTS.body,
        color: COLORS.darkText, align: "right", valign: "middle"
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 3.1, y: y + 0.05, w: barWidth, h: 0.25, fill: { color: barColor }
      });
      slide.addText(`${d.count} (${d.pct}%)`, {
        x: 3.1 + barWidth + 0.1, y, w: 1.5, h: 0.35, fontSize: 9, fontFace: FONTS.body,
        color: COLORS.mutedText, valign: "middle"
      });
    });
  }
}

function addRecommendationsSlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.accent };

  slide.addText("AI Recommendations", {
    x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 28, fontFace: FONTS.header,
    color: COLORS.white, bold: true, margin: 0
  });

  const recs = [
    { title: "Immediate Focus", desc: data.top5[0] ? `Prioritize ${data.top5[0].name} — highest composite score combining business value and customer demand` : "Run scoring pipeline with updated data" },
    { title: "Quick Wins", desc: "Bundle related high-priority catalog and merchandising capabilities into a single release for maximum market impact" },
    { title: "Strategic Investment", desc: "AI-powered pricing and promotion optimization shows strong strategic value — invest ahead of competitors" },
    { title: "Customer Alignment", desc: `${data.survey.total_respondents} surveyed enterprises validate demand — use this data to align sales messaging with roadmap` },
  ];

  recs.forEach((rec, i) => {
    const y = 1.1 + i * 1.1;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y, w: 8.4, h: 0.95, fill: { color: COLORS.primary, transparency: 30 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y, w: 0.06, h: 0.95, fill: { color: COLORS.secondary }
    });
    slide.addText(rec.title, {
      x: 1.15, y: y + 0.08, w: 7.8, h: 0.3, fontSize: 14, fontFace: FONTS.header,
      color: COLORS.white, bold: true
    });
    slide.addText(rec.desc, {
      x: 1.15, y: y + 0.42, w: 7.8, h: 0.45, fontSize: 11, fontFace: FONTS.body,
      color: COLORS.lightGray
    });
  });
}

function addNextStepsSlide(pres, data) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.accent };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.8, fill: { color: COLORS.primary, transparency: 40 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.825, w: 10, h: 0.8, fill: { color: COLORS.primary, transparency: 40 }
  });

  slide.addText("Next Steps", {
    x: 0.8, y: 1.2, w: 8.4, h: 0.6, fontSize: 32, fontFace: FONTS.header,
    color: COLORS.white, bold: true, margin: 0
  });

  const steps = [
    "Review and validate scoring weights with leadership team",
    "Schedule 1-hour roadmap interviews with willing survey participants",
    "Align Q1-Q2 engineering sprints to top-5 priority capabilities",
    "Set up n8n automation for continuous scoring updates",
  ];

  steps.forEach((step, i) => {
    slide.addText([
      { text: `${i + 1}. `, options: { color: COLORS.secondary, bold: true, fontSize: 14 } },
      { text: step, options: { color: COLORS.lightGray, fontSize: 13 } }
    ], {
      x: 1.2, y: 2.1 + i * 0.55, w: 7.6, h: 0.45, fontFace: FONTS.body
    });
  });

  slide.addText("AI Product Manager v1.0 — Powered by Intelligent Prioritization", {
    x: 0.8, y: 5.0, w: 8.4, h: 0.3, fontSize: 10, fontFace: FONTS.body,
    color: COLORS.mutedText, align: "center"
  });
}

// Main execution — reads JSON data from stdin or file
const args = process.argv.slice(2);
if (args.length < 2) {
  console.error("Usage: node deck_generator.js <data.json> <output.pptx>");
  process.exit(1);
}

const dataPath = args[0];
const outputPath = args[1];
const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
createDeck(data, outputPath);
