import * as ExcelJS from "exceljs";
import JSZip from "jszip";

type Worksheet = ExcelJS.Worksheet;

export type ConversionStats = {
  dataRows: number;
  drRows: number;
  baseFormulaRows: number;
};

export type ConvertedWorkbook = {
  blob: Blob;
  outputName: string;
  stats: ConversionStats;
};

const HEADER_ROW = 1;
const FIRST_DATA_ROW = 2;
const SAP_LAST_COLUMN_INDEX = 19;
const PROCESS_DOCUMENT_TYPE = "DR";
const PENALTY_RATE = 0.001;

const HEADERS: Record<number, string> = {
  20: "Дней просрочки по КК (календарные)",
  21: "% КК",
  22: "Сумма КК",
  23: "Сумма закрытия",
  24: "К доплате после пересчета платежа:",
  25: "к доплате в том числе по основному долгу",
  26: "к доплате в том числе по КК",
  27: "Количество дней просрочки долга после пересчета",
};

const FORMULA_COLUMNS = [20, 21, 22, 23];
const NO_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "none",
};

export async function convertXlsxFile(file: File, calcDate: string): Promise<ConvertedWorkbook> {
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    throw new Error("Поддерживаются только файлы .xlsx");
  }

  const buffer = await file.arrayBuffer();
  return convertXlsxBuffer(buffer, file.name, calcDate);
}

export async function convertXlsxBuffer(
  buffer: ArrayBuffer,
  fileName: string,
  calcDate: string,
): Promise<ConvertedWorkbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("В книге нет листов");
  }

  validateInputHeaders(worksheet);
  const stats = transformWorksheet(worksheet, calcDate);
  workbook.calcProperties.fullCalcOnLoad = true;

  const output = await workbook.xlsx.writeBuffer();
  const patchedOutput = await preserveOriginalColorPalette(buffer, output);

  return {
    blob: new Blob([patchedOutput], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    outputName: buildOutputName(fileName),
    stats,
  };
}

async function preserveOriginalColorPalette(sourceBuffer: ArrayBuffer, outputBuffer: ExcelJS.Buffer): Promise<ArrayBuffer> {
  const sourceZip = await JSZip.loadAsync(sourceBuffer);
  const outputZip = await JSZip.loadAsync(outputBuffer);

  const sourceStyles = await sourceZip.file("xl/styles.xml")?.async("string");
  const outputStyles = await outputZip.file("xl/styles.xml")?.async("string");

  if (sourceStyles && outputStyles) {
    const sourceColors = sourceStyles.match(/<colors>[\s\S]*?<\/colors>/)?.[0];
    if (sourceColors) {
      const stylesWithoutColors = outputStyles.replace(/<colors>[\s\S]*?<\/colors>/, "");
      outputZip.file("xl/styles.xml", stylesWithoutColors.replace("</styleSheet>", `${sourceColors}</styleSheet>`));
    }
  }

  const sourceTheme = await sourceZip.file("xl/theme/theme1.xml")?.async("uint8array");
  if (sourceTheme) {
    outputZip.file("xl/theme/theme1.xml", sourceTheme);
  }

  return outputZip.generateAsync({ type: "arraybuffer" });
}

function validateInputHeaders(worksheet: Worksheet): void {
  const requiredHeaders: Record<number, string> = {
    15: "Дата платежа",
    16: "Дата выравнивания",
    17: "Вид документа",
    18: "Сумма в ВВ",
    19: "Валюта документа",
  };

  for (const [columnAsString, expected] of Object.entries(requiredHeaders)) {
    const column = Number(columnAsString);
    const actual = normalizedCellValue(worksheet.getCell(HEADER_ROW, column));
    if (actual !== expected) {
      throw new Error(`Не найден ожидаемый заголовок "${expected}" в колонке ${column}`);
    }
  }
}

function transformWorksheet(worksheet: Worksheet, calcDate: string): ConversionStats {
  const lastDataRow = findLastDataRow(worksheet);
  removeExistingCalculationCells(worksheet, lastDataRow);
  writeHeaders(worksheet, calcDate);

  let drRows = 0;
  let baseFormulaRows = 0;

  for (let row = FIRST_DATA_ROW; row <= lastDataRow; row += 1) {
    const documentType = normalizedCellValue(worksheet.getCell(row, 17));

    if (documentType === PROCESS_DOCUMENT_TYPE) {
      drRows += 1;
      writeBaseFormulas(worksheet, row);
      baseFormulaRows += 1;
    } else {
      clearColumns(worksheet, row, FORMULA_COLUMNS);
    }
  }

  applyLayout(worksheet);

  return {
    dataRows: Math.max(0, lastDataRow - FIRST_DATA_ROW + 1),
    drRows,
    baseFormulaRows,
  };
}

function writeHeaders(worksheet: Worksheet, calcDate: string): void {
  for (const [columnAsString, header] of Object.entries(HEADERS)) {
    const column = Number(columnAsString);
    const cell = worksheet.getCell(HEADER_ROW, column);
    cell.value = header;
    cell.style = {
      alignment: { vertical: "top", wrapText: true },
      fill: NO_FILL,
    };
  }

  const calcDateCell = worksheet.getCell("AB1");
  calcDateCell.value = parseLocalDate(calcDate);
  calcDateCell.style = {
    fill: NO_FILL,
    numFmt: "yyyy-mm-dd",
  };
}

function writeBaseFormulas(worksheet: Worksheet, row: number): void {
  const clearingDate = normalizedCellValue(worksheet.getCell(row, 16));
  const overdueFormula =
    clearingDate == null || clearingDate === ""
      ? `MAX(0,$AB$1-O${row})`
      : `MAX(0,P${row}-O${row})`;

  writeFormula(worksheet.getCell(row, 20), overdueFormula, "0");

  const rateCell = worksheet.getCell(row, 21);
  rateCell.value = PENALTY_RATE;
  rateCell.style = {
    fill: NO_FILL,
    numFmt: "0.00%",
  };

  writeFormula(worksheet.getCell(row, 22), `U${row}*T${row}*R${row}`, "#,##0.00");
  writeFormula(worksheet.getCell(row, 23), `V${row}+R${row}`, "#,##0.00");
}

function writeFormula(cell: ExcelJS.Cell, formula: string, numberFormat: string): void {
  cell.value = { formula };
  cell.style = {
    fill: NO_FILL,
    numFmt: numberFormat,
  };
}

function removeExistingCalculationCells(worksheet: Worksheet, lastDataRow: number): void {
  const maxColumn = Math.max(worksheet.columnCount, 28);
  const maxRow = Math.max(worksheet.rowCount, lastDataRow);

  for (let column = 20; column <= maxColumn; column += 1) {
    worksheet.getColumn(column).style = {};
  }

  for (let row = HEADER_ROW; row <= maxRow; row += 1) {
    const columns = Array.from({ length: maxColumn - 19 }, (_, index) => index + 20);
    clearColumns(worksheet, row, columns);
  }
}

function clearColumns(worksheet: Worksheet, row: number, columns: number[]): void {
  for (const column of columns) {
    const cell = worksheet.getCell(row, column);
    cell.value = null;
    cell.style = { fill: NO_FILL };
  }
}

function applyLayout(worksheet: Worksheet): void {
  const widths: Record<number, number> = {
    20: 18,
    21: 10,
    22: 14,
    23: 14,
    24: 18,
    25: 18,
    26: 14,
    27: 18,
    28: 12,
  };

  for (const [columnAsString, width] of Object.entries(widths)) {
    worksheet.getColumn(Number(columnAsString)).width = width;
  }

  worksheet.autoFilter = undefined;
}

function findLastDataRow(worksheet: Worksheet): number {
  for (let row = worksheet.rowCount; row >= FIRST_DATA_ROW; row -= 1) {
    for (let column = 1; column <= SAP_LAST_COLUMN_INDEX; column += 1) {
      const value = normalizedCellValue(worksheet.getCell(row, column));
      if (value != null && value !== "") {
        return row;
      }
    }
  }
  return HEADER_ROW;
}

function normalizedCellValue(cell: ExcelJS.Cell): unknown {
  const value = cell.value;

  if (value == null) {
    return null;
  }

  if (typeof value === "object" && "result" in value) {
    return value.result ?? null;
  }

  if (typeof value === "object" && "text" in value) {
    return value.text ?? null;
  }

  if (typeof value === "object" && "richText" in value && Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join("");
  }

  return value;
}

function parseLocalDate(value: string): Date {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function buildOutputName(sourceName: string): string {
  const cleanName = sourceName.replace(/\.xlsx$/i, "");
  return `${cleanName}_converted.xlsx`;
}
