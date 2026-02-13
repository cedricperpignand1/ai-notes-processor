import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const TEMPLATE_PATH = path.join(process.cwd(), "templates", "SOV_template.xlsx");

// Your existing mapping (rows where line items go)
const DIVISION_RANGES = {
  "DIV 2": { startRow: 12, endRow: 13 },
  "DIV 3": { startRow: 16, endRow: 19 },
  "DIV 4": { startRow: 21, endRow: 22 },
  "DIV 5": { startRow: 24, endRow: 29 },
  "DIV 6": { startRow: 31, endRow: 39 },
  "DIV 7": { startRow: 41, endRow: 49 },
  "DIV 8": { startRow: 51, endRow: 51 },
  "DIV 9": { startRow: 55, endRow: 58 },
  "DIV 10": { startRow: 61, endRow: 61 },
  "DIV 11": { startRow: 63, endRow: 64 },
  "DIV 12": { startRow: 66, endRow: 67 },
  "DIV 13": { startRow: 69, endRow: 70 },
  "DIV 14": { startRow: 72, endRow: 73 },
  "DIV 21": { startRow: 75, endRow: 76 },
  "DIV 22": { startRow: 78, endRow: 79 },
  "DIV 23": { startRow: 81, endRow: 81 },
  "DIV 26": { startRow: 82, endRow: 84 },
  "DIV 27": { startRow: 86, endRow: 87 },
  "DIV 28": { startRow: 90, endRow: 90 },
  "DIV 32": { startRow: 92, endRow: 94 },
};

// Division names (like PDF)
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

function setBase(cell) {
  cell.font = { name: "Calibri", size: 10 };
  cell.alignment = { vertical: "middle", wrapText: true };
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

function ensureSheet(workbook, name) {
  let ws = workbook.getWorksheet(name);
  if (!ws) ws = workbook.addWorksheet(name);
  setupPrint(ws);
  return ws;
}

function headerBar(ws, rangeA1, text) {
  ws.mergeCells(rangeA1);
  const c = ws.getCell(rangeA1.split(":")[0]);
  c.value = text;
  c.font = { name: "Calibri", size: 14, bold: true };
  c.alignment = { horizontal: "left", vertical: "middle" };
}

function styleTableHeader(ws, rowNum, fromCol, toCol) {
  const row = ws.getRow(rowNum);
  row.height = 18;
  for (let col = fromCol; col <= toCol; col++) {
    const cell = row.getCell(col);
    setBase(cell);
    cell.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF7A7A7A" } }; // dark gray
    cell.border = thinBorder();
  }
}

function styleRow(ws, rowNum, fromCol, toCol) {
  const row = ws.getRow(rowNum);
  row.height = 16;
  for (let col = fromCol; col <= toCol; col++) {
    const cell = row.getCell(col);
    setBase(cell);
    cell.border = thinBorder();
  }
}

function styleDivisionRow(ws, rowNum, fromCol, toCol) {
  const row = ws.getRow(rowNum);
  row.height = 16;
  for (let col = fromCol; col <= toCol; col++) {
    const cell = row.getCell(col);
    setBase(cell);
    cell.font = { name: "Calibri", size: 10, bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD0D0D0" } }; // medium gray
    cell.border = thinBorder();
  }
}

function buildPDFLikeLayout(workbook) {
  const cover = ensureSheet(workbook, "COVER");
  const summary = ensureSheet(workbook, "SUMMARY");
  const genReq = ensureSheet(workbook, "GEN REQ");
  const detailed = ensureSheet(workbook, "DETAILED SOV");
  const alt = ensureSheet(workbook, "ALT EXCLUSIONS");

  // COVER (simple – your template can do the fancy blocks)
  cover.getColumn("A").width = 2;
  cover.getColumn("B").width = 14;
  cover.getColumn("G").width = 45;

  // SUMMARY (keep as PDF: has comments)
  summary.columns = [
    { width: 10 }, // A DIV
    { width: 38 }, // B Desc
    { width: 32 }, // C Comments
    { width: 10 }, // D Qty
    { width: 12 }, // E Rate
    { width: 4 },  // F $
    { width: 14 }, // G Amount
  ];
  headerBar(summary, "A1:G1", "SUMMARY SOV");
  summary.getRow(1).height = 26;

  summary.getRow(4).values = ["", "TRADE CODE DESCRIPTIONS", "COMMENTS", "QTY", "RATE", "SUB TOTAL", ""];
  styleTableHeader(summary, 4, 1, 7);

  // Rows like PDF
  const divList = ["DIV 1","DIV 2","DIV 3","DIV 4","DIV 5","DIV 6","DIV 7","DIV 8","DIV 9","DIV 10","DIV 11","DIV 12","DIV 13","DIV 14","DIV 21","DIV 22","DIV 23","DIV 26","DIV 27","DIV 28","DIV 31","DIV 32","DIV 33"];
  let r = 5;
  divList.forEach((div) => {
    summary.getCell(`A${r}`).value = div;
    summary.getCell(`B${r}`).value = DIV_NAMES[div] || "";
    summary.getCell(`F${r}`).value = "$";
    summary.getCell(`G${r}`).value = { formula: "0" }; // will be overwritten by formulas later
    summary.getCell(`G${r}`).numFmt = moneyFmt();
    styleRow(summary, r, 1, 7);

    // NIC shading like PDF for some divisions
    if (["DIV 27","DIV 28","DIV 31","DIV 32","DIV 33"].includes(div)) {
      summary.getCell(`G${r}`).value = "NIC";
      summary.getCell(`G${r}`).alignment = { horizontal: "right", vertical: "middle" };
      summary.getCell(`G${r}`).font = { name: "Calibri", size: 10, bold: true };
    }

    r++;
  });

  // Totals block like PDF
  const costRow = r + 2;
  summary.mergeCells(`A${costRow}:E${costRow}`);
  summary.getCell(`A${costRow}`).value = "TOTAL COST OF CONSTRUCTION";
  summary.getCell(`F${costRow}`).value = "$";
  summary.getCell(`G${costRow}`).numFmt = moneyFmt();
  styleRow(summary, costRow, 1, 7);

  const feeRow = costRow + 1;
  summary.mergeCells(`A${feeRow}:D${feeRow}`);
  summary.getCell(`A${feeRow}`).value = "Contractor Fee";
  summary.getCell(`E${feeRow}`).value = "0.00%";
  summary.getCell(`F${feeRow}`).value = "$";
  summary.getCell(`G${feeRow}`).numFmt = moneyFmt();
  styleRow(summary, feeRow, 1, 7);

  const totalRow = feeRow + 1;
  summary.mergeCells(`A${totalRow}:E${totalRow}`);
  summary.getCell(`A${totalRow}`).value = "TOTAL BUDGET";
  summary.getCell(`A${totalRow}`).font = { name: "Calibri", size: 11, bold: true };
  summary.getCell(`F${totalRow}`).value = "$";
  summary.getCell(`G${totalRow}`).numFmt = moneyFmt();
  for (let c = 1; c <= 7; c++) {
    const cell = summary.getRow(totalRow).getCell(c);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB7CFDD" } }; // light blue
    cell.border = thinBorder();
  }

  // GEN REQ (PDF keeps comments)
  genReq.columns = [
    { width: 6 },  // A (blank like PDF)
    { width: 44 }, // B Description
    { width: 34 }, // C Comments
    { width: 10 }, // D Qty
    { width: 12 }, // E Rate
    { width: 4 },  // F $
    { width: 14 }, // G Amount
  ];
  headerBar(genReq, "A1:G1", "GENERAL CONDITIONS SUMMARY");
  genReq.getRow(1).height = 26;

  genReq.getRow(4).values = ["", "DESCRIPTION", "COMMENTS", "QTY", "RATE", "SUB TOTAL", ""];
  styleTableHeader(genReq, 4, 1, 7);

  // DETAILED SOV — ✅ REMOVE COMMENTS COLUMN (per your request)
  // Columns like PDF BUT without COMMENTS:
  // A = DIV, B = DESCRIPTION, C = QTY, D = RATE, E = $, F = SUBTOTAL, G = $, H = DIV TOTAL
  detailed.columns = [
    { width: 8 },  // A DIV
    { width: 72 }, // B Description (this is where AI scope goes)
    { width: 10 }, // C Qty
    { width: 12 }, // D Rate
    { width: 4 },  // E $
    { width: 14 }, // F Subtotal
    { width: 4 },  // G $
    { width: 14 }, // H Div Total
  ];

  // Top small title like PDF (left)
  detailed.getCell("A1").value = "DETAILED SOV";
  detailed.getCell("A1").font = { name: "Calibri", size: 10, bold: true, color: { argb: "FF2F75B5" } };

  detailed.getRow(4).values = ["DIVISION", " ", "QTY", "RATE", "SUB TOTAL", "", "DIV TOTAL", ""];
  styleTableHeader(detailed, 4, 1, 8);
  // Make header labels match PDF spacing
  detailed.mergeCells("A4:B4");
  detailed.mergeCells("E4:F4");
  detailed.mergeCells("G4:H4");
  detailed.getCell("A4").value = "DIVISION";
  detailed.getCell("C4").value = "QTY";
  detailed.getCell("D4").value = "RATE";
  detailed.getCell("E4").value = "SUB TOTAL";
  detailed.getCell("G4").value = "DIV TOTAL";

  // ALT EXCLUSIONS
  headerBar(alt, "A1:H1", "ALT / NOTES / EXCLUSIONS");
  alt.getColumn("A").width = 110;
}

function sumRangeFormula(wsName, colLetter, startRow, endRow) {
  return { formula: `SUM('${wsName}'!${colLetter}${startRow}:${colLetter}${endRow})` };
}

export async function generateExcel(project, divisions) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ANDCON AI Estimator";

  if (fs.existsSync(TEMPLATE_PATH)) {
    await workbook.xlsx.readFile(TEMPLATE_PATH);
  } else {
    buildPDFLikeLayout(workbook);
  }

  // Ensure all 5 sheets exist (exact names)
  const cover = ensureSheet(workbook, "COVER");
  const summary = ensureSheet(workbook, "SUMMARY");
  const detailed = ensureSheet(workbook, "DETAILED SOV");
  const alt = ensureSheet(workbook, "ALT EXCLUSIONS");

  // COVER fields
  cover.getCell("G26").value = project?.name || "";
  cover.getCell("G27").value = project?.address || "";
  cover.getCell("G28").value = project?.version || "";
  cover.getCell("G29").value = project?.date || "";

  // ✅ Populate DETAILED SOV line items under each division with AI description
  // Put AI scope in Description column (B). If you want title + scope, we do two lines.
  for (const divData of divisions || []) {
    const range = DIVISION_RANGES[divData.division];
    if (!range) continue;

    const capacity = range.endRow - range.startRow + 1;
    const items = (divData.items || []).slice(0, capacity);

    items.forEach((item, idx) => {
      const rowNum = range.startRow + idx;

      // A = DIV code
      detailed.getCell(`A${rowNum}`).value = divData.division;

      // B = DESCRIPTION (AI output)
      const aiDesc = (item?.scope || "").trim();
      const title = (item?.title || "").trim();

      detailed.getCell(`B${rowNum}`).value = aiDesc
        ? (title ? `${title}\n${aiDesc}` : aiDesc)
        : (title || "");

      // Qty/Rate if you have them (optional)
      if (typeof item?.qty === "number") detailed.getCell(`C${rowNum}`).value = item.qty;
      if (typeof item?.rate === "number") detailed.getCell(`D${rowNum}`).value = item.rate;

      // Subtotal split: E = "$", F = amount
      detailed.getCell(`E${rowNum}`).value = "$";
      const sub = detailed.getCell(`F${rowNum}`);
      sub.numFmt = moneyFmt();
      sub.value = { formula: `IF(OR(C${rowNum}="",D${rowNum}=""),"",C${rowNum}*D${rowNum})` };

      // Div total split lives at the *end row* of each division block (like PDF)
      // We'll set it after the loop.

      // Style
      styleRow(detailed, rowNum, 1, 8);
      detailed.getCell(`B${rowNum}`).alignment = { wrapText: true, vertical: "top" };
      detailed.getRow(rowNum).height = 34; // room for 2 lines
    });

    // Division total at end row: G="$" H=sum(subtotals)
    detailed.getCell(`G${range.endRow}`).value = "$";
    const divTot = detailed.getCell(`H${range.endRow}`);
    divTot.numFmt = moneyFmt();
    divTot.font = { name: "Calibri", size: 10, bold: true };
    divTot.value = sumRangeFormula("DETAILED SOV", "F", range.startRow, range.endRow);
    styleRow(detailed, range.endRow, 1, 8);
  }

  // ✅ SUMMARY: Pull totals from DETAILED SOV per division
  // Find each DIV row in SUMMARY and set amount in column G.
  // (This is simple: scan SUMMARY column A for "DIV X")
  summary.eachRow((row, rowNumber) => {
    const div = String(row.getCell(1).value || "").trim();
    if (!div.startsWith("DIV")) return;
    const range = DIVISION_RANGES[div];
    if (!range) return;

    // Column F = "$", Column G = amount
    summary.getCell(`F${rowNumber}`).value = "$";
    const amt = summary.getCell(`G${rowNumber}`);
    amt.numFmt = moneyFmt();
    amt.value = sumRangeFormula("DETAILED SOV", "F", range.startRow, range.endRow);
  });

  // ALT EXCLUSIONS block text
  alt.getCell("A4").value =
    "• Change Orders that reduce scope will be credited at the supplier/vendor credited amount.\n" +
    "• All applicable sales taxes are included unless noted.\n" +
    "• Pricing valid for 30 days.\n" +
    "• Contractor fee applies to accepted alternates unless noted.\n" +
    "• Hazardous materials/asbestos survey or abatement not included.\n";
  alt.getCell("A4").alignment = { wrapText: true, vertical: "top" };
  alt.getRow(4).height = 160;

  return await workbook.xlsx.writeBuffer();
}
