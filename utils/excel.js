import ExcelJS from "exceljs";

/**
 * divisions input expected:
 * [
 *   { division: "DIV 2", items: [{ title: "...", scope: "..." }, ...] },
 *   ...
 * ]
 */

const DIV_NAMES = {
  "DIV 1": "GENERAL REQUIREMENTS",
  "DIV 2": "SITEWORK",
  "DIV 3": "CONCRETE",
  "DIV 4": "MASONRY",
  "DIV 5": "METALS",
  "DIV 6": "WOODS & PLASTICS",
  "DIV 7": "THERMAL & MOISTURE PROTECTION",
  "DIV 8": "OPENINGS",
  "DIV 9": "FINISHES",
  "DIV 10": "SPECIALTIES",
  "DIV 11": "EQUIPMENT",
  "DIV 12": "FURNISHINGS",
  "DIV 13": "SPECIAL CONSTRUCTION",
  "DIV 14": "CONVEYING SYSTEMS",
  "DIV 21": "FIRE SUPPRESSION",
  "DIV 22": "PLUMBING",
  "DIV 23": "HVAC",
  "DIV 26": "ELECTRICAL",
  "DIV 27": "TELECOMMUNICATIONS",
  "DIV 28": "ELECTRONIC SAFETY AND SECURITY",
  "DIV 31": "EARTHWORK",
  "DIV 32": "EXTERIOR IMPROVEMENTS",
  "DIV 33": "UTILITIES",
};

const GEN_REQ_LINES = [
  { desc: "Master Permit Fees", comment: "" },
  { desc: "Trade Permits", comment: "Included with trade estimates" },
  { desc: "Permit Processing", comment: "" },
  { desc: "Miscellaneous Tools", comment: "" },
  { desc: "Messenger Service", comment: "" },
  { desc: "Blueprints", comment: "" },
  { desc: "Parking", comment: "" },
  { desc: "Job Office Supplies", comment: "" },
  { desc: "Telephone", comment: "" },
  { desc: "Project Management Personnel", comment: "Weekly rates" },
  { desc: "Supervision Personnel", comment: "Weekly rates" },
  { desc: "General Site Labor", comment: "" },
  { desc: "Carpenters", comment: "General Temporary Structures" },
  { desc: "Security", comment: "" },
  { desc: "XRAY", comment: "IF REQUIRED" },
  { desc: "Temporary Power & Lighting", comment: "" },
  { desc: "Temporary Toilets", comment: "" },
  { desc: "Temporary Fire Protection", comment: "" },
  { desc: "First Aid and Safety", comment: "" },
  { desc: "Rubbish Removal (Dumpsters)", comment: "" },
  { desc: "Temporary Field Offices", comment: "" },
  { desc: "Temporary Protection", comment: "" },
  { desc: "Insurance", comment: "General Liability" },
  { desc: "Project Cleaning", comment: "Rough & Final" },
];

function moneyFmt() {
  return '"$"#,##0;[Red]-"$"#,##0;"$"-;@';
}

function thinBorder() {
  return {
    top: { style: "thin", color: { argb: "FFBFBFBF" } },
    left: { style: "thin", color: { argb: "FFBFBFBF" } },
    bottom: { style: "thin", color: { argb: "FFBFBFBF" } },
    right: { style: "thin", color: { argb: "FFBFBFBF" } },
  };
}

function baseCell(cell) {
  cell.font = { name: "Calibri", size: 10 };
  cell.alignment = { vertical: "middle", wrapText: true };
}

function applyBorderRow(ws, rowNum, fromCol, toCol) {
  const row = ws.getRow(rowNum);
  for (let c = fromCol; c <= toCol; c++) {
    const cell = row.getCell(c);
    baseCell(cell);
    cell.border = thinBorder();
  }
}

function setupPrint(ws) {
  ws.pageSetup = {
    paperSize: 1, // Letter
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    showGridLines: false,
    margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.2, footer: 0.2 },
  };
  ws.views = [{ showGridLines: false }];
}

function headerBar(ws, rowNum, fromCol, toCol, text) {
  ws.mergeCells(rowNum, fromCol, rowNum, toCol);
  const cell = ws.getCell(rowNum, fromCol);
  cell.value = text;
  cell.font = { name: "Calibri", size: 14, bold: true };
  cell.alignment = { horizontal: "left", vertical: "middle" };
  ws.getRow(rowNum).height = 24;
}

function darkHeaderRow(ws, rowNum, labels) {
  const row = ws.getRow(rowNum);
  row.values = labels;
  row.height = 18;

  for (let c = 1; c <= labels.length; c++) {
    const cell = row.getCell(c);
    baseCell(cell);
    cell.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF6E6E6E" } };
    cell.border = thinBorder();
  }
}

function divisionBandRow(ws, rowNum, fromCol, toCol, leftText, rightText = "") {
  ws.mergeCells(rowNum, fromCol, rowNum, toCol - 2); // big left band
  const left = ws.getCell(rowNum, fromCol);
  left.value = leftText;
  left.font = { name: "Calibri", size: 10, bold: true };
  left.alignment = { horizontal: "left", vertical: "middle" };
  left.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } };
  left.border = thinBorder();

  // right side cells (for $ / div total if you want it visible on band)
  const r1 = ws.getCell(rowNum, toCol - 1);
  const r2 = ws.getCell(rowNum, toCol);
  [r1, r2].forEach((c) => {
    baseCell(c);
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } };
    c.border = thinBorder();
  });

  if (rightText) r2.value = rightText;

  ws.getRow(rowNum).height = 16;
}

function fillRange(ws, startRow, startCol, endRow, endCol, argb) {
  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r);
    for (let c = startCol; c <= endCol; c++) {
      const cell = row.getCell(c);
      baseCell(cell);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
    }
  }
}

export async function generateExcel(project, divisions) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "ANDCON AI Estimator";

  // ─────────────────────────────────────────────
  // Create 5 sheets (exact names)
  // ─────────────────────────────────────────────
  const cover = wb.addWorksheet("COVER");
  const summary = wb.addWorksheet("SUMMARY");
  const genReq = wb.addWorksheet("GEN REQ");
  const detailed = wb.addWorksheet("DETAILED SOV");
  const alt = wb.addWorksheet("ALT EXCLUSIONS");

  [cover, summary, genReq, detailed, alt].forEach(setupPrint);

  // ─────────────────────────────────────────────
  // COVER (light gray full background like your image)
  // ─────────────────────────────────────────────
  // Set columns A-L widths so it fills like a cover page block
  for (let i = 1; i <= 12; i++) cover.getColumn(i).width = i === 12 ? 12 : 10;
  for (let r = 1; r <= 55; r++) cover.getRow(r).height = 15;

  // fill big area with light gray
  fillRange(cover, 1, 1, 55, 11, "FFCFCFCF");

  // Simple cover text placement (adjust as you want)
  cover.mergeCells("B10:K10");
  cover.getCell("B10").value = project?.name || "";
  cover.getCell("B10").font = { name: "Calibri", size: 20, bold: true };
  cover.getCell("B10").alignment = { horizontal: "center", vertical: "middle" };

  cover.mergeCells("B13:K13");
  cover.getCell("B13").value = project?.address || "";
  cover.getCell("B13").font = { name: "Calibri", size: 12 };
  cover.getCell("B13").alignment = { horizontal: "center", vertical: "middle" };

  cover.mergeCells("B18:K18");
  cover.getCell("B18").value = `Estimate: ${project?.version || ""}`;
  cover.getCell("B18").font = { name: "Calibri", size: 11, italic: true };
  cover.getCell("B18").alignment = { horizontal: "center", vertical: "middle" };

  cover.mergeCells("B45:K45");
  cover.getCell("B45").value = project?.date || "";
  cover.getCell("B45").font = { name: "Calibri", size: 11, italic: true, color: { argb: "FFFFFFFF" } };
  cover.getCell("B45").alignment = { horizontal: "center", vertical: "middle" };

  // Remove borders on cover (PDF cover is clean)
  cover.views = [{ showGridLines: false }];

  // ─────────────────────────────────────────────
  // SUMMARY (PDF-like)
  // ─────────────────────────────────────────────
  summary.columns = [
    { width: 10 }, // A div
    { width: 38 }, // B name
    { width: 30 }, // C comments
    { width: 10 }, // D qty
    { width: 12 }, // E rate
    { width: 4 },  // F $
    { width: 14 }, // G subtotal
  ];

  headerBar(summary, 1, 1, 7, "SUMMARY SOV");
  darkHeaderRow(summary, 4, ["DIVISION", "TRADE CODE DESCRIPTIONS", "COMMENTS", "QTY", "RATE", "SUB TOTAL", ""]);

  const summaryStart = 5;
  const summaryDivRows = {}; // div -> row

  const divOrder = Object.keys(DIV_NAMES);
  let sr = summaryStart;
  for (const div of divOrder) {
    summary.getCell(sr, 1).value = div;
    summary.getCell(sr, 2).value = DIV_NAMES[div];
    summary.getCell(sr, 6).value = "$";
    summary.getCell(sr, 7).numFmt = moneyFmt();
    applyBorderRow(summary, sr, 1, 7);
    summaryDivRows[div] = sr;
    sr++;
  }

  // Totals block
  sr += 1;
  summary.mergeCells(sr, 1, sr, 5);
  summary.getCell(sr, 1).value = "TOTAL COST OF CONSTRUCTION";
  summary.getCell(sr, 6).value = "$";
  summary.getCell(sr, 7).numFmt = moneyFmt();
  applyBorderRow(summary, sr, 1, 7);
  const totalCostRow = sr;

  sr++;
  summary.mergeCells(sr, 1, sr, 4);
  summary.getCell(sr, 1).value = "Contractor Fee";
  summary.getCell(sr, 5).value = "0.00%";
  summary.getCell(sr, 6).value = "$";
  summary.getCell(sr, 7).numFmt = moneyFmt();
  applyBorderRow(summary, sr, 1, 7);
  const feeRow = sr;

  sr++;
  summary.mergeCells(sr, 1, sr, 5);
  summary.getCell(sr, 1).value = "TOTAL BUDGET";
  summary.getCell(sr, 1).font = { name: "Calibri", size: 11, bold: true };
  summary.getCell(sr, 6).value = "$";
  summary.getCell(sr, 7).numFmt = moneyFmt();
  for (let c = 1; c <= 7; c++) {
    const cell = summary.getCell(sr, c);
    baseCell(cell);
    cell.border = thinBorder();
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB7CFDD" } };
  }
  const totalBudgetRow = sr;

  // ─────────────────────────────────────────────
  // GEN REQ (build like your screenshot, fully populated labels)
  // ─────────────────────────────────────────────
  genReq.columns = [
    { width: 44 }, // A description
    { width: 28 }, // B comments
    { width: 10 }, // C qty
    { width: 12 }, // D rate
    { width: 4 },  // E $
    { width: 14 }, // F subtotal
  ];

  headerBar(genReq, 1, 1, 6, "GENERAL CONDITIONS SUMMARY");
  darkHeaderRow(genReq, 4, ["DESCRIPTION", "COMMENTS", "QTY", "RATE", "SUB", ""]);

  let gr = 5;
  for (const line of GEN_REQ_LINES) {
    genReq.getCell(gr, 1).value = line.desc;
    genReq.getCell(gr, 2).value = line.comment || "";
    genReq.getCell(gr, 5).value = "$";
    genReq.getCell(gr, 6).numFmt = moneyFmt();
    applyBorderRow(genReq, gr, 1, 6);
    gr++;
  }

  // General Conditions Total
  gr += 1;
  genReq.mergeCells(gr, 1, gr, 4);
  genReq.getCell(gr, 1).value = "GENERAL CONDITIONS TOTAL";
  genReq.getCell(gr, 1).font = { name: "Calibri", size: 10, bold: true };
  genReq.getCell(gr, 5).value = "$";
  genReq.getCell(gr, 6).numFmt = moneyFmt();
  for (let c = 1; c <= 6; c++) {
    const cell = genReq.getCell(gr, c);
    baseCell(cell);
    cell.border = thinBorder();
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } };
  }
  // Sum gen req subtotal column F
  genReq.getCell(gr, 6).value = { formula: `SUM(F5:F${gr - 2})` };

  // ─────────────────────────────────────────────
  // DETAILED SOV (your requested layout)
  // Division band row + item rows, no blank rows, title + description next to it
  // ─────────────────────────────────────────────
  detailed.columns = [
    { width: 8 },  // A DIV
    { width: 28 }, // B Line Item Title
    { width: 66 }, // C Description (AI scope)
    { width: 10 }, // D Qty
    { width: 12 }, // E Rate
    { width: 4 },  // F $
    { width: 14 }, // G Subtotal
    { width: 4 },  // H $
    { width: 14 }, // I Div Total
  ];

  detailed.getCell("A1").value = "DETAILED SOV";
  detailed.getCell("A1").font = { name: "Calibri", size: 10, bold: true, color: { argb: "FF2F75B5" } };

  darkHeaderRow(detailed, 4, ["DIVISION", "LINE ITEM", "DESCRIPTION", "QTY", "RATE", "SUB TOTAL", "", "DIV TOTAL", ""]);
  // Make header look like PDF (merge $ labels visually)
  detailed.mergeCells("F4:G4");
  detailed.mergeCells("H4:I4");
  detailed.getCell("F4").value = "SUB TOTAL";
  detailed.getCell("H4").value = "DIV TOTAL";

  let dr = 5;

  // Track division total cells for SUMMARY formulas
  const detailedDivTotalCellByDiv = {};

  for (const divData of divisions || []) {
    const div = (divData.division || "").trim();
    if (!div) continue;

    const divName = DIV_NAMES[div] || "";
    // Division header band row (light gray)
    divisionBandRow(detailed, dr, 1, 9, `${div}   ${divName}`);
    // also show "$" columns on band
    detailed.getCell(dr, 8).value = "$"; // H
    detailed.getCell(dr, 9).numFmt = moneyFmt();
    detailed.getCell(dr, 9).value = ""; // div total will be set after items
    detailedDivTotalCellByDiv[div] = `I${dr}`;
    dr++;

    const items = Array.isArray(divData.items) ? divData.items : [];
    if (items.length === 0) {
      // If no items, still keep one blank line to preserve structure
      detailed.getCell(dr, 1).value = div;
      detailed.getCell(dr, 6).value = "$";
      detailed.getCell(dr, 7).numFmt = moneyFmt();
      detailed.getCell(dr, 8).value = "$";
      detailed.getCell(dr, 9).numFmt = moneyFmt();
      applyBorderRow(detailed, dr, 1, 9);
      dr++;
      continue;
    }

    const startItemRow = dr;

    for (const item of items) {
      const title = (item?.title || "").trim();
      const scope = (item?.scope || "").trim();

      detailed.getCell(dr, 1).value = div;
      detailed.getCell(dr, 2).value = title || "";
      detailed.getCell(dr, 3).value = scope || "";

      // qty/rate (optional if you don’t have them yet)
      if (typeof item?.qty === "number") detailed.getCell(dr, 4).value = item.qty;
      if (typeof item?.rate === "number") detailed.getCell(dr, 5).value = item.rate;

      // subtotal split
      detailed.getCell(dr, 6).value = "$";
      detailed.getCell(dr, 7).numFmt = moneyFmt();
      detailed.getCell(dr, 7).value = { formula: `IF(OR(D${dr}="",E${dr}=""),"",D${dr}*E${dr})` };

      // leave div total columns blank on item lines
      detailed.getCell(dr, 8).value = "";
      detailed.getCell(dr, 9).value = "";

      applyBorderRow(detailed, dr, 1, 9);

      // better row height for description
      detailed.getRow(dr).height = 28;
      detailed.getCell(dr, 3).alignment = { wrapText: true, vertical: "top" };

      dr++;
    }

    const endItemRow = dr - 1;

    // Set division total on the DIV BAND row (same row as header band)
    const bandRow = startItemRow - 1;
    detailed.getCell(bandRow, 8).value = "$";
    detailed.getCell(bandRow, 9).numFmt = moneyFmt();
    detailed.getCell(bandRow, 9).value = { formula: `SUM(G${startItemRow}:G${endItemRow})` };
  }

  // ─────────────────────────────────────────────
  // SUMMARY totals: point each division row to the DETAILED SOV band row totals
  // ─────────────────────────────────────────────
  for (const [div, bandCell] of Object.entries(detailedDivTotalCellByDiv)) {
    const row = summaryDivRows[div];
    if (!row) continue;

    summary.getCell(row, 6).value = "$";
    summary.getCell(row, 7).numFmt = moneyFmt();
    summary.getCell(row, 7).value = { formula: `'DETAILED SOV'!${bandCell}` };
  }

  // Total cost of construction = sum all summary subtotals
  summary.getCell(totalCostRow, 7).value = { formula: `SUM(G${summaryStart}:G${summaryStart + divOrder.length - 1})` };
  // Fee (if you want it calculated later you can modify this)
  summary.getCell(feeRow, 7).value = 0;
  // Total Budget = cost + fee
  summary.getCell(totalBudgetRow, 7).value = { formula: `G${totalCostRow}+G${feeRow}` };

  // ─────────────────────────────────────────────
  // ALT EXCLUSIONS
  // ─────────────────────────────────────────────
  alt.getColumn(1).width = 110;
  headerBar(alt, 1, 1, 1, "ALT / NOTES / EXCLUSIONS");
  alt.getCell("A4").value =
    "• Change Orders that reduce scope will be credited at the supplier/vendor credited amount.\n" +
    "• All applicable sales taxes are included unless noted.\n" +
    "• Pricing valid for 30 days.\n" +
    "• Contractor fee applies to accepted alternates unless noted.\n" +
    "• Hazardous materials/asbestos survey or abatement not included.\n";
  alt.getCell("A4").alignment = { wrapText: true, vertical: "top" };
  alt.getRow(4).height = 160;

  // Return file
  return await wb.xlsx.writeBuffer();
}
