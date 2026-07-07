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
  const result = await convertXlsxBuffer(source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength), path.basename(file), "2026-07-01");
  const bytes = new Uint8Array(await result.blob.arrayBuffer());
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  const worksheet = workbook.worksheets[0];

  assertEqual(worksheet.getCell("T7").formula, "MAX(0,P7-O7)", `${file} T7`);
  assertEqual(worksheet.getCell("V7").formula, "U7*T7*R7", `${file} V7`);
  assertEqual(worksheet.getCell("W7").formula, "V7+R7", `${file} W7`);
  assertEqual(worksheet.getCell("X7").formula, undefined, `${file} X7`);
  assertEqual(worksheet.getCell("T2").formula, undefined, `${file} T2`);
  assertEqual(result.stats.drRows, 1013, `${file} DR count`);

  console.log(`${path.basename(file)} ok: ${result.outputName}`);
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(`${label}: expected ${expected}, got ${actual}`);
  }
}
