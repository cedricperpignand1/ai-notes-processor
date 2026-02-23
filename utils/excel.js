import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * divisions input expected:
 * [
 *   { division: "DIV 2", items: [{ title: "...", scope: "..." , qty?: number, rate?: number }, ...] },
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

/**
 * Division band row (gray) + keep borders
 */
function divisionBandRow(ws, rowNum, fromCol, toCol, text) {
  ws.mergeCells(rowNum, fromCol, rowNum, toCol);
  const cell = ws.getCell(rowNum, fromCol);
  cell.value = text;
  cell.font = { name: "Calibri", size: 10, bold: true };
  cell.alignment = { horizontal: "left", vertical: "middle" };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } };

  // Border entire row range
  for (let c = fromCol; c <= toCol; c++) {
    const cc = ws.getCell(rowNum, c);
    baseCell(cc);
    cc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } };
    cc.border = thinBorder();
  }

  ws.getRow(rowNum).height = 16;
}

export async function generateExcel(project, divisions) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "ANDCON AI Estimator";

  // Load logo (best-effort — skip if file is missing)
  let logoId = null;
  try {
    const logoBuffer = fs.readFileSync(path.join(__dirname, '..', 'assets', 'logo.png'));
    logoId = wb.addImage({ buffer: logoBuffer, extension: 'png' });
  } catch (_) {}

  // 5 Sheets
  const cover = wb.addWorksheet("COVER");
  const summary = wb.addWorksheet("SUMMARY");
  const genReq = wb.addWorksheet("GEN REQ");
  const detailed = wb.addWorksheet("DETAILED SOV");
  const alt = wb.addWorksheet("ALT EXCLUSIONS");

  [cover, summary, genReq, detailed, alt].forEach(setupPrint);

  // ───────────────────────── COVER ─────────────────────────
  for (let i = 1; i <= 12; i++) cover.getColumn(i).width = i === 12 ? 12 : 10;
  for (let r = 1; r <= 55; r++) cover.getRow(r).height = 15;

  // light gray page block like your template
  fillRange(cover, 1, 1, 55, 11, "FFCFCFCF");

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

  // Logo — bottom right of cover
  if (logoId !== null) {
    cover.addImage(logoId, { tl: { col: 8, row: 48 }, ext: { width: 180, height: 56 }, editAs: 'oneCell' });
  }

  // ───────────────────────── SUMMARY ─────────────────────────
  summary.columns = [
    { width: 10 }, // A div
    { width: 38 }, // B name
    { width: 30 }, // C comments
    { width: 10 }, // D qty
    { width: 12 }, // E rate
    { width: 16 }, // F subtotal (currency)
  ];

  headerBar(summary, 1, 1, 6, "SUMMARY SOV");
  if (logoId !== null) {
    summary.addImage(logoId, { tl: { col: 4, row: 0 }, ext: { width: 160, height: 50 }, editAs: 'oneCell' });
  }
  darkHeaderRow(summary, 4, ["DIVISION", "TRADE CODE DESCRIPTIONS", "COMMENTS", "QTY", "RATE", "SUB TOTAL"]);

  const summaryStart = 5;
  const summaryDivRows = {}; // div -> row
  const divOrder = Object.keys(DIV_NAMES);

  let sr = summaryStart;
  divOrder.forEach((div) => {
    summary.getCell(sr, 1).value = div;
    summary.getCell(sr, 2).value = DIV_NAMES[div];
    summary.getCell(sr, 6).numFmt = moneyFmt();
    applyBorderRow(summary, sr, 1, 6);
    summaryDivRows[div] = sr;
    sr++;
  });

  // Totals block (like your screenshot structure)
  sr += 1;
  summary.mergeCells(sr, 1, sr, 5);
  summary.getCell(sr, 1).value = "COST OF CONSTRUCTION";
  summary.getCell(sr, 1).font = { name: "Calibri", size: 10, bold: true };
  summary.getCell(sr, 6).numFmt = moneyFmt();
  applyBorderRow(summary, sr, 1, 6);
  const summaryCostRow = sr;

  sr++;
  summary.mergeCells(sr, 1, sr, 3);
  summary.getCell(sr, 1).value = "ANDCON FEE";
  summary.getCell(sr, 1).font = { name: "Calibri", size: 10 };
  summary.getCell(sr, 4).value = 0.05; // 5% default
  summary.getCell(sr, 4).numFmt = "0.00%";
  summary.getCell(sr, 6).numFmt = moneyFmt();
  applyBorderRow(summary, sr, 1, 6);
  const summaryFeeRow = sr;

  sr++;
  summary.mergeCells(sr, 1, sr, 5);
  summary.getCell(sr, 1).value = "TOTAL";
  summary.getCell(sr, 1).font = { name: "Calibri", size: 11, bold: true };
  summary.getCell(sr, 6).numFmt = moneyFmt();
  for (let c = 1; c <= 6; c++) {
    const cell = summary.getCell(sr, c);
    baseCell(cell);
    cell.border = thinBorder();
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB7CFDD" } };
  }
  const summaryTotalRow = sr;

  // ───────────────────────── GEN REQ ─────────────────────────
  genReq.columns = [
    { width: 44 }, // A description
    { width: 28 }, // B comments
    { width: 10 }, // C qty
    { width: 12 }, // D rate
    { width: 16 }, // E subtotal
  ];

  headerBar(genReq, 1, 1, 5, "GENERAL CONDITIONS SUMMARY");
  if (logoId !== null) {
    genReq.addImage(logoId, { tl: { col: 3, row: 0 }, ext: { width: 160, height: 50 }, editAs: 'oneCell' });
  }
  darkHeaderRow(genReq, 4, ["DESCRIPTION", "COMMENTS", "QTY", "RATE", "SUB"]);

  let gr = 5;
  GEN_REQ_LINES.forEach((line) => {
    genReq.getCell(gr, 1).value = line.desc;
    genReq.getCell(gr, 2).value = line.comment || "";
    genReq.getCell(gr, 5).numFmt = moneyFmt();
    genReq.getCell(gr, 5).value = { formula: `IF(OR(C${gr}="",D${gr}=""),"",C${gr}*D${gr})` };
    applyBorderRow(genReq, gr, 1, 5);
    gr++;
  });

  gr += 1;
  genReq.mergeCells(gr, 1, gr, 4);
  genReq.getCell(gr, 1).value = "GENERAL CONDITIONS TOTAL";
  genReq.getCell(gr, 1).font = { name: "Calibri", size: 10, bold: true };
  genReq.getCell(gr, 5).numFmt = moneyFmt();
  genReq.getCell(gr, 5).value = { formula: `SUM(E5:E${gr - 2})` };
  for (let c = 1; c <= 5; c++) {
    const cell = genReq.getCell(gr, c);
    baseCell(cell);
    cell.border = thinBorder();
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } };
  }

  // ───────────────────────── DETAILED SOV ─────────────────────────
  // ✅ No separate $ cells; currency formatting only
  detailed.columns = [
    { width: 8 },  // A DIV
    { width: 28 }, // B Line Item Title
    { width: 66 }, // C Description (AI scope)
    { width: 12 }, // D Qty
    { width: 14 }, // E Rate
    { width: 18 }, // F Subtotal
    { width: 18 }, // G Div Total
  ];

  detailed.getCell("A1").value = "DETAILED SOV";
  detailed.getCell("A1").font = { name: "Calibri", size: 10, bold: true, color: { argb: "FF2F75B5" } };

  if (logoId !== null) {
    detailed.addImage(logoId, { tl: { col: 5, row: 0 }, ext: { width: 160, height: 50 }, editAs: 'oneCell' });
  }

  darkHeaderRow(detailed, 4, ["DIVISION", "LINE ITEM", "DESCRIPTION", "QTY", "RATE", "SUB TOTAL", "DIV TOTAL"]);

  let dr = 5;
  const detailedDivTotalCellByDiv = {}; // div -> cell address (G band row)
  const divisionTotalCells = []; // for cost of construction sum

  for (const divData of divisions || []) {
    const div = (divData.division || "").trim();
    if (!div) continue;

    const divName = DIV_NAMES[div] || "";

    // Division band row
    divisionBandRow(detailed, dr, 1, 6, `${div}   ${divName}`);
    // Div total cell (G on band)
    detailed.getCell(dr, 7).numFmt = moneyFmt();
    detailed.getCell(dr, 7).border = thinBorder();
    detailed.getCell(dr, 7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } };

    const bandRow = dr;
    dr++;

    const startRow = dr;

    const items = Array.isArray(divData.items) ? divData.items : [];
    for (const item of items) {
      detailed.getCell(dr, 2).value = (item?.title || "").trim();
      detailed.getCell(dr, 3).value = (item?.scope || "").trim();

      if (typeof item?.qty === "number") detailed.getCell(dr, 4).value = item.qty;
      if (typeof item?.rate === "number") detailed.getCell(dr, 5).value = item.rate;

      detailed.getCell(dr, 6).numFmt = moneyFmt();
      detailed.getCell(dr, 6).value = { formula: `IF(OR(D${dr}="",E${dr}=""),"",D${dr}*E${dr})` };

      // Div total column blank on line items (no borders — keeps the column visually open)
      detailed.getCell(dr, 7).value = "";

      applyBorderRow(detailed, dr, 1, 6);

      // Light blue fill on the DIVISION column (A) for line item rows
      const divCell = detailed.getCell(dr, 1);
      baseCell(divCell);
      divCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6E4F0" } };
      divCell.border = thinBorder();
      detailed.getRow(dr).height = 28;
      detailed.getCell(dr, 3).alignment = { wrapText: true, vertical: "top" };

      dr++;
    }

    const endRow = dr - 1;

    // Division total formula on band row
    detailed.getCell(bandRow, 7).value = { formula: `SUM(F${startRow}:F${endRow})` };

    const bandCell = `G${bandRow}`;
    detailedDivTotalCellByDiv[div] = bandCell;
    divisionTotalCells.push(bandCell);
  }

  // Bottom totals block on DETAILED SOV (like your screenshot)
  dr += 1;

  // COST OF CONSTRUCTION
  detailed.mergeCells(dr, 1, dr, 5);
  detailed.getCell(dr, 1).value = "COST OF CONSTRUCTION";
  detailed.getCell(dr, 1).font = { name: "Calibri", size: 10, bold: true };
  detailed.getCell(dr, 6).numFmt = moneyFmt();
  detailed.getCell(dr, 6).value = {
    formula: divisionTotalCells.length ? `SUM(${divisionTotalCells.join(",")})` : "0",
  };
  applyBorderRow(detailed, dr, 1, 7);
  const costRow = dr;
  dr++;

  // ANDCON FEE (with %)
  detailed.mergeCells(dr, 1, dr, 3);
  detailed.getCell(dr, 1).value = "ANDCON FEE";
  detailed.getCell(dr, 4).value = 0.05;
  detailed.getCell(dr, 4).numFmt = "0.00%";

  detailed.getCell(dr, 6).numFmt = moneyFmt();
  detailed.getCell(dr, 6).value = { formula: `F${costRow}*D${dr}` };

  applyBorderRow(detailed, dr, 1, 7);
  const feeRow = dr;
  dr++;

  // TOTAL
  detailed.mergeCells(dr, 1, dr, 5);
  detailed.getCell(dr, 1).value = "TOTAL";
  detailed.getCell(dr, 1).font = { name: "Calibri", size: 11, bold: true };
  detailed.getCell(dr, 6).numFmt = moneyFmt();
  detailed.getCell(dr, 6).value = { formula: `F${costRow}+F${feeRow}` };

  for (let c = 1; c <= 7; c++) {
    const cell = detailed.getCell(dr, c);
    baseCell(cell);
    cell.border = thinBorder();
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF9EB6C4" } };
  }

  // ───────────────────────── SUMMARY FORMULAS LINKED TO DETAILED ─────────────────────────
  // Each division subtotal in SUMMARY pulls from band totals in Detailed SOV
  for (const [div, bandCell] of Object.entries(detailedDivTotalCellByDiv)) {
    const row = summaryDivRows[div];
    if (!row) continue;
    summary.getCell(row, 6).value = { formula: `'DETAILED SOV'!${bandCell}` };
  }

  // Cost of construction on SUMMARY = sum all division rows
  summary.getCell(summaryCostRow, 6).value = { formula: `SUM(F${summaryStart}:F${summaryStart + divOrder.length - 1})` };
  // Summary fee = cost * %
  summary.getCell(summaryFeeRow, 6).value = { formula: `F${summaryCostRow}*D${summaryFeeRow}` };
  // Total = cost + fee
  summary.getCell(summaryTotalRow, 6).value = { formula: `F${summaryCostRow}+F${summaryFeeRow}` };

  // ───────────────────────── ALT EXCLUSIONS ─────────────────────────
  alt.getColumn(1).width = 88;
  alt.getColumn(2).width = 22;
  headerBar(alt, 1, 1, 2, "ALT / NOTES / EXCLUSIONS");
  if (logoId !== null) {
    alt.addImage(logoId, { tl: { col: 1, row: 0 }, ext: { width: 160, height: 50 }, editAs: 'oneCell' });
  }

  alt.getCell("A4").value =
    "• Change Orders that reduce scope will be credited at the supplier/vendor credited amount.\n" +
    "• All applicable sales taxes are included unless noted.\n" +
    "• Pricing valid for 30 days.\n" +
    "• Contractor fee applies to accepted alternates unless noted.\n" +
    "• Hazardous materials/asbestos survey or abatement not included.\n";
  alt.getCell("A4").alignment = { wrapText: true, vertical: "top" };
  alt.getRow(4).height = 160;

  // Return
  return await wb.xlsx.writeBuffer();
}
