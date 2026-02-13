import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

const TEMPLATE_PATH = path.join(process.cwd(), "templates", "SOV_template.xlsx");

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

const SUMMARY_ROWS = {
  "DIV 1": 7,
  "DIV 2": 8,
  "DIV 3": 9,
  "DIV 4": 10,
  "DIV 5": 11,
  "DIV 6": 12,
  "DIV 7": 13,
  "DIV 8": 14,
  "DIV 9": 15,
  "DIV 10": 16,
  "DIV 11": 17,
  "DIV 12": 18,
  "DIV 13": 19,
  "DIV 14": 20,
  "DIV 21": 21,
  "DIV 22": 22,
  "DIV 23": 23,
  "DIV 26": 24,
  "DIV 27": 25,
  "DIV 28": 26,
  "DIV 31": 27,
  "DIV 32": 28,
  "DIV 33": 29,
};

function moneyFmt() {
  return '"$"#,##0;[Red]-"$"#,##0;"$"-;@';
}

function setCellBase(cell) {
  cell.font = { name: "Calibri", size: 10 };
  cell.alignment = { vertical: "middle", wrapText: true };
}

function styleHeaderRow(ws, rowNum, fromCol, toCol) {
  const row = ws.getRow(rowNum);
  row.height = 18;

  for (let c = fromCol; c <= toCol; c++) {
    const cell = row.getCell(c);
    setCellBase(cell);
    cell.font = { name: "Calibri", size: 10, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFEFEF" } };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  }
}

function borderRow(ws, rowNum, fromCol, toCol) {
  const row = ws.getRow(rowNum);
  for (let c = fromCol; c <= toCol; c++) {
    const cell = row.getCell(c);
    setCellBase(cell);
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  }
}

function setupPrint(ws) {
  ws.pageSetup = {
    paperSize: 1, // 1 = Letter
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    showGridLines: false,
    margins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
  };
}

function ensureSheet(workbook, name) {
  let ws = workbook.getWorksheet(name);
  if (!ws) ws = workbook.addWorksheet(name);
  setupPrint(ws);
  return ws;
}

function buildMinimalLayout(workbook) {
  const cover = ensureSheet(workbook, "COVER");
  const summary = ensureSheet(workbook, "SUMMARY");
  const genReq = ensureSheet(workbook, "GEN REQ");
  const detailed = ensureSheet(workbook, "DETAILED SOV");
  const alt = ensureSheet(workbook, "ALT EXCLUSIONS");

  // COVER
  cover.getColumn("A").width = 2;
  cover.getColumn("B").width = 14;
  cover.getColumn("G").width = 45;

  // SUMMARY
  summary.columns = [
    { width: 12 }, // A
    { width: 36 }, // B
    { width: 18 }, // C
    { width: 10 }, // D
    { width: 12 }, // E
    { width: 14 }, // F
  ];

  summary.mergeCells("A2:F2");
  summary.getCell("A2").value = "SUMMARY SOV";
  summary.getCell("A2").font = { name: "Calibri", size: 14, bold: true };
  summary.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  summary.getRow(2).height = 24;

  summary.getRow(5).values = ["", "", "COMMENTS", "QTY", "RATE", "SUB TOTAL"];
  styleHeaderRow(summary, 5, 1, 6);

  Object.entries(SUMMARY_ROWS).forEach(([div, r]) => {
    summary.getCell(`A${r}`).value = div;
    summary.getCell(`F${r}`).numFmt = moneyFmt();
    borderRow(summary, r, 1, 6);
  });

  const totalRow = 33;
  summary.getCell(`E${totalRow}`).value = "TOTAL BUDGET";
  summary.getCell(`E${totalRow}`).font = { name: "Calibri", size: 11, bold: true };
  summary.getCell(`F${totalRow}`).numFmt = moneyFmt();
  summary.getCell(`F${totalRow}`).font = { name: "Calibri", size: 11, bold: true };
  borderRow(summary, totalRow, 1, 6);

  // GEN REQ
  genReq.columns = [
    { width: 45 }, // A
    { width: 18 }, // B
    { width: 10 }, // C
    { width: 12 }, // D
    { width: 14 }, // E
  ];

  genReq.mergeCells("A2:E2");
  genReq.getCell("A2").value = "GENERAL CONDITIONS SUMMARY";
  genReq.getCell("A2").font = { name: "Calibri", size: 14, bold: true };
  genReq.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  genReq.getRow(2).height = 24;

  genReq.getRow(5).values = ["DESCRIPTION", "COMMENTS", "QTY", "RATE", "SUB TOTAL"];
  styleHeaderRow(genReq, 5, 1, 5);

  // DETAILED SOV
  detailed.columns = [
    { width: 12 }, // A DIV
    { width: 30 }, // B DESC
    { width: 18 }, // C COMMENTS
    { width: 10 }, // D QTY
    { width: 12 }, // E RATE
    { width: 14 }, // F SUB TOTAL
    { width: 14 }, // G DIV TOTAL
  ];

  detailed.mergeCells("A2:G2");
  detailed.getCell("A2").value = "DETAILED SOV";
  detailed.getCell("A2").font = { name: "Calibri", size: 14, bold: true };
  detailed.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  detailed.getRow(2).height = 24;

  detailed.getRow(6).values = ["DIV", "DESCRIPTION", "COMMENTS", "QTY", "RATE", "SUB TOTAL", "DIV TOTAL"];
  styleHeaderRow(detailed, 6, 1, 7);

  // ALT EXCLUSIONS
  alt.getColumn("A").width = 110;
  alt.getCell("A2").value = "NOTES / ADD ALTERNATES / EXCLUSIONS";
  alt.getCell("A2").font = { name: "Calibri", size: 14, bold: true };
  alt.getRow(2).height = 24;
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
    buildMinimalLayout(workbook);
  }

  const cover = ensureSheet(workbook, "COVER");
  const summary = ensureSheet(workbook, "SUMMARY");
  const detailed = ensureSheet(workbook, "DETAILED SOV");
  const alt = ensureSheet(workbook, "ALT EXCLUSIONS");

  // COVER
  cover.getCell("G26").value = project?.name || "";
  cover.getCell("G27").value = project?.address || "";
  cover.getCell("G28").value = project?.version || "";
  cover.getCell("G29").value = project?.date || "";

  // DETAILED SOV line items
  for (const divData of divisions || []) {
    const range = DIVISION_RANGES[divData.division];
    if (!range) continue;

    const capacity = range.endRow - range.startRow + 1;
    const items = (divData.items || []).slice(0, capacity);

    items.forEach((item, idx) => {
      const rowNum = range.startRow + idx;

      detailed.getCell(`A${rowNum}`).value = divData.division;
      detailed.getCell(`B${rowNum}`).value = item?.title || "";
      detailed.getCell(`C${rowNum}`).value = ""; // comments

      if (typeof item?.qty === "number") detailed.getCell(`D${rowNum}`).value = item.qty;
      if (typeof item?.rate === "number") detailed.getCell(`E${rowNum}`).value = item.rate;

      const f = detailed.getCell(`F${rowNum}`);
      f.numFmt = moneyFmt();
      f.value = { formula: `IF(OR(D${rowNum}="",E${rowNum}=""),"",D${rowNum}*E${rowNum})` };

      borderRow(detailed, rowNum, 1, 7);
    });

    // division total in column G at end row
    const divTotalCell = detailed.getCell(`G${range.endRow}`);
    divTotalCell.numFmt = moneyFmt();
    divTotalCell.font = { name: "Calibri", size: 10, bold: true };
    divTotalCell.value = sumRangeFormula("DETAILED SOV", "F", range.startRow, range.endRow);
  }

  // SUMMARY totals pull from DETAILED SOV
  Object.entries(DIVISION_RANGES).forEach(([div, r]) => {
    const summaryRow = SUMMARY_ROWS[div];
    if (!summaryRow) return;

    const cell = summary.getCell(`F${summaryRow}`);
    cell.numFmt = moneyFmt();
    cell.value = sumRangeFormula("DETAILED SOV", "F", r.startRow, r.endRow);
    borderRow(summary, summaryRow, 1, 6);
  });

  // TOTAL BUDGET
  const totalRow = 33;
  const divRows = Object.values(SUMMARY_ROWS).sort((a, b) => a - b);
  const first = divRows[0];
  const last = divRows[divRows.length - 1];
  summary.getCell(`F${totalRow}`).numFmt = moneyFmt();
  summary.getCell(`F${totalRow}`).value = { formula: `SUM(F${first}:F${last})` };

  // ALT EXCLUSIONS block
  alt.getCell("A4").value =
    "Change Orders resulting in a reduction in scope will be credited back at the cost credited by suppliers/vendors.\n" +
    "All applicable sales taxes are included.\n" +
    "Pricing valid for 30 days.\n" +
    "Note: GC Fee to be added to accepted alternates as necessary.\n" +
    "Asbestos/hazardous materials survey/abatement is not included.\n";
  alt.getCell("A4").alignment = { wrapText: true, vertical: "top" };
  alt.getRow(4).height = 160;

  return await workbook.xlsx.writeBuffer();
}
