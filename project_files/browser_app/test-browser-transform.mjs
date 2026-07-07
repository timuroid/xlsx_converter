import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { convertXlsxBuffer } from "./dist-test/transformer-test-entry.mjs";

const projectRoot = path.resolve("../..");
const files = [
  path.join(projectRoot, "Источники/inbox/100019128.XLSX"),
  path.join(projectRoot, "Источники/inbox/тестовая табличка_0.xlsx"),
];

for (const file of files) {
  const source = await fs.readFile(file);
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.load(source);
  const sourceWorksheet = sourceWorkbook.worksheets[0];

  const result = await convertXlsxBuffer(source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength), path.basename(file), "2026-07-01");
  const bytes = new Uint8Array(await result.blob.arrayBuffer());
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  const worksheet = workbook.worksheets[0];

  assertEqual(worksheet.getCell("T7").formula, "MAX(0,P7-O7)", `${file} T7`);
  assertEqual(worksheet.getCell("V7").formula, "U7*T7*R7", `${file} V7`);
  assertEqual(worksheet.getCell("W7").formula, "V7+R7", `${file} W7`);
  assertEqual(worksheet.getCell("X1").value, "К доплате после пересчета платежа:", `${file} X1`);
  assertEqual(worksheet.getCell("Y1").value, "к доплате в том числе по основному долгу", `${file} Y1`);
  assertEqual(worksheet.getCell("Z1").value, "к доплате в том числе по КК", `${file} Z1`);
  assertEqual(worksheet.getCell("AA1").value, "Количество дней просрочки долга после пересчета", `${file} AA1`);
  assertEqual(worksheet.getCell("X7").formula, undefined, `${file} X7`);
  assertEqual(worksheet.getCell("Y7").formula, undefined, `${file} Y7`);
  assertEqual(worksheet.getCell("Z7").formula, undefined, `${file} Z7`);
  assertEqual(worksheet.getCell("AA7").formula, undefined, `${file} AA7`);
  assertEqual(worksheet.getCell("T2").formula, undefined, `${file} T2`);
  assertNoVisibleFill(worksheet.getCell("T1"), `${file} T1 fill`);
  assertNoVisibleFill(worksheet.getCell("X1"), `${file} X1 fill`);
  assertNoVisibleFill(worksheet.getCell("T7"), `${file} T7 fill`);
  assertNoVisibleFill(worksheet.getCell("U7"), `${file} U7 fill`);
  assertNoVisibleFill(worksheet.getCell("X7"), `${file} X7 fill`);
  assertNoVisibleFill(worksheet.getCell("AB1"), `${file} AB1 fill`);
  assertSourceStylesPreserved(sourceWorksheet, worksheet, file);
  assertEqual(result.stats.drRows, 1013, `${file} DR count`);

  console.log(`${path.basename(file)} ok: ${result.outputName}`);
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}

function assertNoVisibleFill(cell, label) {
  const fill = cell.fill;
  if (fill?.type && fill.pattern !== "none") {
    throw new Error(`${label}: expected no visible fill, got ${JSON.stringify(fill)}`);
  }
}

function assertSourceStylesPreserved(sourceWorksheet, outputWorksheet, label) {
  const addresses = ["A1", "O1", "Q1", "S1", "A7", "O7", "Q7", "S7"];

  for (const address of addresses) {
    const sourceStyle = stableStringify(sourceWorksheet.getCell(address).style ?? {});
    const outputStyle = stableStringify(outputWorksheet.getCell(address).style ?? {});
    assertEqual(outputStyle, sourceStyle, `${label} ${address} source style`);
  }
}

function stableStringify(value) {
  if (Array.isArray(value)) {
    return `[${value.map(stableStringify).join(",")}]`;
  }

  if (value && typeof value === "object") {
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
      .join(",")}}`;
  }

  return JSON.stringify(value);
}
