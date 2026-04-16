#!/usr/bin/env node

const fs = require("fs");
const os = require("os");
const path = require("path");

const ExcelJS = require("exceljs");
const XLSX = require("xlsx");

const ROWS = parseInt(process.env.RBXL_BENCH_ROWS || "5000", 10);
const COLS = parseInt(process.env.RBXL_BENCH_COLS || "10", 10);
const WARMUP = parseInt(process.env.RBXL_BENCH_WARMUP || "1", 10);
const ITERATIONS = parseInt(process.env.RBXL_BENCH_ITERATIONS || "5", 10);
const READ_PATH = process.env.RBXL_BENCH_READ_PATH;

function rssKb() {
  return Math.round(process.memoryUsage().rss / 1024);
}

function buildDataset(rows, cols) {
  const header = Array.from({ length: cols }, (_, i) => `col_${i + 1}`);
  const body = Array.from({ length: rows }, (_, row) =>
    Array.from({ length: cols }, (_, col) => {
      switch (col % 4) {
        case 0:
          return row;
        case 1:
          return `row-${row}-col-${col}`;
        case 2:
          return (row + col) % 2 === 1;
        default:
          return ((row * 100) + col) / 10.0;
      }
    })
  );
  return { header, body };
}

async function measure(label, fn) {
  for (let i = 0; i < WARMUP; i += 1) {
    if (global.gc) global.gc();
    await fn();
  }

  const samples = [];
  const rssDeltas = [];
  let result = null;

  for (let i = 0; i < ITERATIONS; i += 1) {
    if (global.gc) global.gc();
    const before = rssKb();
    const started = process.hrtime.bigint();
    result = await fn();
    const elapsed = Number(process.hrtime.bigint() - started) / 1e9;
    samples.push(elapsed);
    rssDeltas.push(Math.max(0, rssKb() - before));
  }

  const mean = samples.reduce((sum, value) => sum + value, 0) / samples.length;
  const variance = samples.reduce((sum, value) => sum + ((value - mean) ** 2), 0) / samples.length;
  return {
    label,
    real: mean,
    real_min: Math.min(...samples),
    real_stddev: Math.sqrt(variance),
    rss_delta_kb: Math.max(...rssDeltas),
    iterations: ITERATIONS,
    result,
  };
}

async function writeWithExcelJs(filePath, header, body) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    filename: filePath,
    useStyles: false,
    useSharedStrings: false
  });
  const sheet = workbook.addWorksheet("Bench");
  sheet.addRow(header).commit();
  for (const row of body) {
    sheet.addRow(row).commit();
  }
  await workbook.commit();
}

async function readWithExcelJs(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("Bench");
  let count = 0;
  sheet.eachRow((row) => {
    count += row.cellCount;
  });
  return count;
}

function writeWithSheetJs(filePath, header, body) {
  const worksheet = XLSX.utils.aoa_to_sheet([header, ...body]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Bench");
  XLSX.writeFile(workbook, filePath, { compression: true });
}

function readWithSheetJs(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets.Bench;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  let count = 0;
  for (const row of rows) {
    count += row.length;
  }
  return count;
}

async function main() {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "rbxl-js-bench-"));
  const { header, body } = buildDataset(ROWS, COLS);
  const results = [];

  const exceljsPath = path.join(tmpDir, "exceljs.xlsx");
  const exceljsWrite = await measure("exceljs write", () => writeWithExcelJs(exceljsPath, header, body));
  exceljsWrite.size = fs.statSync(exceljsPath).size;
  results.push(exceljsWrite);
  results.push(await measure("exceljs read", () => readWithExcelJs(READ_PATH || exceljsPath)));

  const sheetjsPath = path.join(tmpDir, "sheetjs.xlsx");
  const sheetjsWrite = await measure("sheetjs write", async () => writeWithSheetJs(sheetjsPath, header, body));
  sheetjsWrite.size = fs.statSync(sheetjsPath).size;
  results.push(sheetjsWrite);
  results.push(await measure("sheetjs read", async () => readWithSheetJs(READ_PATH || sheetjsPath)));

  process.stdout.write(`${JSON.stringify(results)}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`);
  process.exit(1);
});
