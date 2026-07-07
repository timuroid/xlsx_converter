import * as ExcelJS from "exceljs";

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
};

const CALCULATION_COLUMNS = Object.keys(HEADERS).map(Number);

const HEADER_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFBDD7EE" },
};

const DATA_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFDDEBF7" },
};

const SERVICE_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFE2F0D9" },
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

  return {
    blob: new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    outputName: buildOutputName(fileName),
    stats,
  };
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
      clearColumns(worksheet, row, CALCULATION_COLUMNS);
    }
  }

  applyLayout(worksheet, lastDataRow);

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
    cell.fill = HEADER_FILL;
    cell.alignment = { vertical: "top", wrapText: true };
    cell.border = thinBorder();
  }

  const calcDateCell = worksheet.getCell("AB1");
  calcDateCell.value = parseLocalDate(calcDate);
  calcDateCell.numFmt = "yyyy-mm-dd";
  calcDateCell.fill = SERVICE_FILL;
  calcDateCell.border = thinBorder();
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
  rateCell.numFmt = "0.00%";
  rateCell.fill = DATA_FILL;

  writeFormula(worksheet.getCell(row, 22), `U${row}*T${row}*R${row}`, "#,##0.00");
  writeFormula(worksheet.getCell(row, 23), `V${row}+R${row}`, "#,##0.00");
}

function writeFormula(cell: ExcelJS.Cell, formula: string, numberFormat: string): void {
  cell.value = { formula };
  cell.numFmt = numberFormat;
  cell.fill = DATA_FILL;
}

function removeExistingCalculationCells(worksheet: Worksheet, lastDataRow: number): void {
  const maxColumn = Math.max(worksheet.columnCount, 28);
  const maxRow = Math.max(worksheet.rowCount, lastDataRow);

  for (let row = HEADER_ROW; row <= maxRow; row += 1) {
    const columns = Array.from({ length: maxColumn - 19 }, (_, index) => index + 20);
    clearColumns(worksheet, row, columns);
  }
}

function clearColumns(worksheet: Worksheet, row: number, columns: number[]): void {
  for (const column of columns) {
    const cell = worksheet.getCell(row, column);
    cell.value = null;
    cell.style = {};
  }
}

function applyLayout(worksheet: Worksheet, lastDataRow: number): void {
  const widths: Record<number, number> = {
    20: 18,
    21: 10,
    22: 14,
    23: 14,
    28: 12,
  };

  for (const [columnAsString, width] of Object.entries(widths)) {
    worksheet.getColumn(Number(columnAsString)).width = width;
  }

  for (let row = FIRST_DATA_ROW; row <= lastDataRow; row += 1) {
    for (const column of CALCULATION_COLUMNS) {
      const cell = worksheet.getCell(row, column);
      if (normalizedCellValue(cell) != null || cell.formula) {
        cell.fill = DATA_FILL;
      }
    }
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

function thinBorder(): Partial<ExcelJS.Borders> {
  return {
    top: { style: "thin", color: { argb: "FF808080" } },
    bottom: { style: "thin", color: { argb: "FF808080" } },
    left: { style: "thin", color: { argb: "FF808080" } },
    right: { style: "thin", color: { argb: "FF808080" } },
  };
}

function buildOutputName(sourceName: string): string {
  const cleanName = sourceName.replace(/\.xlsx$/i, "");
  return `${cleanName}_converted.xlsx`;
}
