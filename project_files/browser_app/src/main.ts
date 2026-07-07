import "./styles.css";
import JSZip from "jszip";
import { convertXlsxFile, type ConvertedWorkbook } from "./transformer";

type FileStatus = "pending" | "processing" | "done" | "error";

type FileRecord = {
  id: string;
  file: File;
  status: FileStatus;
  message: string;
  result?: ConvertedWorkbook;
  downloadUrl?: string;
};

const calcDateInput = queryRequired<HTMLInputElement>("#calcDate");
const dateTrigger = queryRequired<HTMLButtonElement>("#dateTrigger");
const dateTriggerText = queryRequired<HTMLSpanElement>("#dateTriggerText");
const calendarPopover = queryRequired<HTMLElement>("#calendarPopover");
const calendarMonth = queryRequired<HTMLElement>("#calendarMonth");
const calendarGrid = queryRequired<HTMLElement>("#calendarGrid");
const prevMonthButton = queryRequired<HTMLButtonElement>("#prevMonth");
const nextMonthButton = queryRequired<HTMLButtonElement>("#nextMonth");
const dropZone = queryRequired<HTMLElement>("#dropZone");
const fileInput = queryRequired<HTMLInputElement>("#fileInput");
const fileList = queryRequired<HTMLUListElement>("#fileList");
const clearButton = queryRequired<HTMLButtonElement>("#clearButton");
const bulkActions = queryRequired<HTMLElement>("#bulkActions");
const downloadAllButton = queryRequired<HTMLButtonElement>("#downloadAllButton");

const records: FileRecord[] = [];
let isProcessing = false;
let isZipping = false;
let selectedDate = startOfDay(new Date());
let visibleMonth = startOfMonth(selectedDate);

setSelectedDate(selectedDate);
renderCalendar();

dateTrigger.addEventListener("click", () => {
  if (calendarPopover.classList.contains("is-hidden")) {
    openCalendar();
  } else {
    closeCalendar();
  }
});

prevMonthButton.addEventListener("click", () => {
  visibleMonth = addMonths(visibleMonth, -1);
  renderCalendar();
});

nextMonthButton.addEventListener("click", () => {
  visibleMonth = addMonths(visibleMonth, 1);
  renderCalendar();
});

document.addEventListener("click", (event) => {
  const target = event.target;
  if (target instanceof Node && !dateTrigger.contains(target) && !calendarPopover.contains(target)) {
    closeCalendar();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeCalendar();
  }
});

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("is-dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("is-dragover");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("is-dragover");
  const files = Array.from(event.dataTransfer?.files ?? []);
  addFiles(files);
});

fileInput.addEventListener("change", () => {
  addFiles(Array.from(fileInput.files ?? []));
  fileInput.value = "";
});

clearButton.addEventListener("click", () => {
  for (const record of records) {
    if (record.downloadUrl) {
      URL.revokeObjectURL(record.downloadUrl);
    }
  }
  records.splice(0, records.length);
  renderRecords();
});

downloadAllButton.addEventListener("click", () => {
  void downloadAllConverted();
});

function addFiles(files: File[]): void {
  for (const file of files) {
    records.push({
      id: crypto.randomUUID(),
      file,
      status: "pending",
      message: "Ожидает обработки",
    });
  }

  renderRecords();
  void processQueue();
}

async function processQueue(): Promise<void> {
  if (isProcessing) {
    return;
  }

  isProcessing = true;
  try {
    for (const record of records) {
      if (record.status !== "pending") {
        continue;
      }

      record.status = "processing";
      record.message = "Конвертация...";
      renderRecords();

      try {
        const result = await convertXlsxFile(record.file, calcDateInput.value);
        record.result = result;
        record.downloadUrl = URL.createObjectURL(result.blob);
        record.status = "done";
        record.message = `Готово: DR ${result.stats.drRows}, строки с расчетом ${result.stats.baseFormulaRows}`;
      } catch (error) {
        record.status = "error";
        record.message = error instanceof Error ? error.message : "Неизвестная ошибка";
      }

      renderRecords();
    }
  } finally {
    isProcessing = false;
  }
}

function renderRecords(): void {
  fileList.replaceChildren(...records.map(renderRecord));
  renderBulkActions();
}

function renderBulkActions(): void {
  const convertedRecords = getConvertedRecords();
  const shouldShow = convertedRecords.length > 1;
  bulkActions.classList.toggle("is-hidden", !shouldShow);
  downloadAllButton.disabled = !shouldShow || isZipping;
  downloadAllButton.textContent = isZipping ? "Готовим архив..." : `Скачать все (${convertedRecords.length})`;
}

async function downloadAllConverted(): Promise<void> {
  const convertedRecords = getConvertedRecords();
  if (convertedRecords.length <= 1 || isZipping) {
    return;
  }

  isZipping = true;
  renderBulkActions();

  try {
    const zip = new JSZip();
    const usedNames = new Set<string>();

    for (const record of convertedRecords) {
      const result = record.result;
      if (!result) {
        continue;
      }
      zip.file(uniqueZipName(result.outputName, usedNames), result.blob);
    }

    const archiveBlob = await zip.generateAsync({ type: "blob" });
    const archiveUrl = URL.createObjectURL(archiveBlob);
    const link = document.createElement("a");
    link.href = archiveUrl;
    link.download = `sap_kk_converted_${calcDateInput.value}.zip`;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(archiveUrl);
  } finally {
    isZipping = false;
    renderBulkActions();
  }
}

function getConvertedRecords(): FileRecord[] {
  return records.filter((record) => record.status === "done" && record.result);
}

function openCalendar(): void {
  calendarPopover.classList.remove("is-hidden");
  dateTrigger.setAttribute("aria-expanded", "true");
}

function closeCalendar(): void {
  calendarPopover.classList.add("is-hidden");
  dateTrigger.setAttribute("aria-expanded", "false");
}

function setSelectedDate(date: Date): void {
  selectedDate = startOfDay(date);
  visibleMonth = startOfMonth(selectedDate);
  calcDateInput.value = toIsoDate(selectedDate);
  dateTriggerText.textContent = formatDisplayDate(selectedDate);
}

function renderCalendar(): void {
  calendarMonth.textContent = formatMonthTitle(visibleMonth);
  const todayIso = toIsoDate(new Date());
  const selectedIso = toIsoDate(selectedDate);
  const weekDays = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];
  const cells: HTMLElement[] = weekDays.map((day) => {
    const cell = document.createElement("div");
    cell.className = "calendar-weekday";
    cell.textContent = day;
    return cell;
  });

  const firstDay = startOfMonth(visibleMonth);
  const daysInVisibleMonth = new Date(visibleMonth.getFullYear(), visibleMonth.getMonth() + 1, 0).getDate();
  const leadingEmptyCells = (firstDay.getDay() + 6) % 7;

  for (let index = 0; index < leadingEmptyCells; index += 1) {
    const cell = document.createElement("div");
    cell.className = "calendar-empty";
    cells.push(cell);
  }

  for (let day = 1; day <= daysInVisibleMonth; day += 1) {
    const date = new Date(visibleMonth.getFullYear(), visibleMonth.getMonth(), day);
    const dateIso = toIsoDate(date);
    const button = document.createElement("button");
    button.className = "calendar-day";
    button.type = "button";
    button.textContent = String(day);
    button.setAttribute("aria-label", formatDisplayDate(date));

    if (dateIso === selectedIso) {
      button.classList.add("is-selected");
      button.setAttribute("aria-current", "date");
    }

    if (dateIso === todayIso) {
      button.classList.add("is-today");
    }

    button.addEventListener("click", () => {
      setSelectedDate(date);
      renderCalendar();
      closeCalendar();
    });
    cells.push(button);
  }

  calendarGrid.replaceChildren(...cells);
}

function renderRecord(record: FileRecord): HTMLLIElement {
  const item = document.createElement("li");
  item.className = "file-item";

  const body = document.createElement("div");
  const name = document.createElement("div");
  name.className = "file-name";
  name.textContent = record.file.name;

  const meta = document.createElement("div");
  meta.className = "file-meta";
  meta.textContent = `${formatBytes(record.file.size)} · ${record.message}`;

  const status = document.createElement("div");
  status.className = `status status--${record.status}`;
  status.textContent = statusText(record.status);

  body.append(name, meta);
  item.append(body);

  if (record.status === "done" && record.downloadUrl && record.result) {
    const link = document.createElement("a");
    link.className = "download-link";
    link.href = record.downloadUrl;
    link.download = record.result.outputName;
    link.textContent = "Скачать";
    item.append(link);
  } else {
    item.append(status);
  }

  return item;
}

function statusText(status: FileStatus): string {
  const labels: Record<FileStatus, string> = {
    pending: "В очереди",
    processing: "В работе",
    done: "Готово",
    error: "Ошибка",
  };
  return labels[status];
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} Б`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} КБ`;
  }
  return `${(bytes / 1024 / 1024).toFixed(1)} МБ`;
}

function uniqueZipName(name: string, usedNames: Set<string>): string {
  if (!usedNames.has(name)) {
    usedNames.add(name);
    return name;
  }

  const dotIndex = name.lastIndexOf(".");
  const baseName = dotIndex > 0 ? name.slice(0, dotIndex) : name;
  const extension = dotIndex > 0 ? name.slice(dotIndex) : "";
  let index = 2;
  let candidate = `${baseName}_${index}${extension}`;

  while (usedNames.has(candidate)) {
    index += 1;
    candidate = `${baseName}_${index}${extension}`;
  }

  usedNames.add(candidate);
  return candidate;
}

function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function startOfMonth(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonths(date: Date, months: number): Date {
  return new Date(date.getFullYear(), date.getMonth() + months, 1);
}

function toIsoDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDisplayDate(date: Date): string {
  return new Intl.DateTimeFormat("ru-RU", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(date);
}

function formatMonthTitle(date: Date): string {
  return new Intl.DateTimeFormat("ru-RU", {
    month: "long",
    year: "numeric",
  }).format(date);
}

function queryRequired<T extends Element>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) {
    throw new Error(`Не удалось найти элемент ${selector}`);
  }
  return element;
}
