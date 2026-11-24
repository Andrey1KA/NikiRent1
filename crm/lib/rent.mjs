const MISSING = "нет данных";
const NO_METER = "нет счётчика";

const HEADER_RULES = [
  ["apartment", "квартира"],
  ["month", "месяц"],
  ["rent", "аренда"],
  ["extra", "доп. расходы"],
  ["utilities", "коммуналка"],
  ["carry", "перенос долга"],
  ["payout", "итого к выплате"],
  ["fin", "фин. учет"],
  ["transferred", "перечислено"],
  ["debt", "долг"],
  ["complete", "полнота"],
  ["gasReading", "газ показание"],
  ["gasUse", "газ расход"],
  ["gasCharge", "газ начислено"],
  ["lightReading", "свет показание"],
  ["lightUse", "свет расход"],
  ["lightCharge", "свет начислено"],
  ["hot1", "горячая 1"],
  ["hot2", "горячая 2"],
  ["hotUse", "горячая расход"],
  ["hotCharge", "горячая начислено"],
  ["cold1", "холодная 1"],
  ["cold2", "холодная 2"],
  ["coldUse", "холодная расход"],
  ["coldCharge", "холодная начислено"],
  ["supplyCharge", "горводснабж начислено"],
];

export function parseWorkbook(XLSX, buffer) {
  let matrix;
  try {
    matrix = readMatrix(XLSX, buffer);
  } catch {
    throw new Error("Файл не похож на таблицу Excel или CSV");
  }
  return parseMatrix(matrix);
}

export function parseMatrix(matrix) {
  if (!Array.isArray(matrix) || matrix.length === 0) throw new Error("В файле нет строк");
  const headerIndex = matrix.findIndex((row) => (row || []).some((cell) => normalize(cell) === "квартира"));
  if (headerIndex < 0) throw new Error("В файле нет колонки «Квартира»");
  const columns = mapColumns(matrix[headerIndex]);
  if (columns.apartment == null || columns.month == null || columns.payout == null) {
    throw new Error("В файле не хватает колонок месяца и итога к выплате");
  }

  const rows = [];
  for (let index = headerIndex + 1; index < matrix.length; index += 1) {
    const source = matrix[index] || [];
    const apartment = textOrNull(source[columns.apartment]);
    const month = textOrNull(source[columns.month]);
    if (!apartment || !month) continue;
    const row = { apartment, month };
    for (const [field] of HEADER_RULES) {
      if (field === "apartment" || field === "month") continue;
      row[field] = parseValue(source[columns[field]]);
    }
    rows.push(row);
  }
  if (rows.length === 0) throw new Error("В файле нет строк с квартирой и месяцем");
  return rows;
}

export function monthsOf(rows) {
  const order = ["апрель", "май"];
  const present = [...new Set(rows.map((row) => row.month))];
  return present.sort((left, right) => {
    const a = order.indexOf(left);
    const b = order.indexOf(right);
    if (a === -1 && b === -1) return left.localeCompare(right, "ru");
    if (a === -1) return 1;
    if (b === -1) return -1;
    return a - b;
  });
}

export function apartmentsOf(rows) {
  return [...new Set(rows.map((row) => row.apartment))];
}

export function rowsForMonth(rows, month) {
  return rows.filter((row) => row.month === month);
}

export function previousOf(rows, row) {
  if (row.month !== "май") return null;
  return rows.find((item) => item.apartment === row.apartment && item.month === "апрель") || null;
}

export function sumField(rows, field) {
  if (rows.length === 0 || rows.some((row) => typeof row[field] !== "number")) return MISSING;
  return rows.reduce((sum, row) => sum + row[field], 0);
}

export function warningsFor(row, previous) {
  const items = [];
  const resources = [
    ["газ", row.gasReading, previous?.gasReading, null, row.gasUse],
    ["свет", row.lightReading, previous?.lightReading, null, row.lightUse],
    ["горячую воду", row.hot1, previous?.hot1, row.hot2, row.hotUse, previous?.hot2],
    ["холодную воду", row.cold1, previous?.cold1, row.cold2, row.coldUse, previous?.cold2],
  ];

  for (const [label, current, prior, currentSecond, consumption, priorSecond] of resources) {
    if (current === MISSING) {
      items.push(`Нет показания: ${label}. Расход не подставлен нулём.`);
    }
    if (typeof current === "number" && typeof prior === "number" && current < prior) {
      items.push(`Новое показание меньше предыдущего: ${label}.`);
    }
    if (typeof currentSecond === "number" && typeof priorSecond === "number" && currentSecond < priorSecond) {
      items.push(`Новое показание второго счётчика меньше предыдущего: ${label}.`);
    }
    if (typeof consumption === "number" && consumption < 0) {
      items.push(`Расход отрицательный: ${label}.`);
    }
  }

  if (typeof row.lightUse === "number" && row.lightUse > 300) {
    items.push(`Расход света ${formatQty(row.lightUse)} кВт·ч, это больше 300. Повод проверить показание.`);
  }
  if (typeof row.hotUse === "number" && typeof row.coldUse === "number" && row.hotUse + row.coldUse > 20) {
    items.push(`Горячая и холодная вода вместе ${formatQty(row.hotUse + row.coldUse)} м³, это больше 20. Повод проверить показание.`);
  }
  if (typeof row.payout === "number" && typeof row.transferred === "number" && row.payout !== row.transferred) {
    const delta = row.payout - row.transferred;
    const kind = delta > 0 ? "Недоплата" : "Переплата";
    items.push(`Начислено ${formatMoney(row.payout)}, перечислено ${formatMoney(row.transferred)}. ${kind} ${formatMoney(Math.abs(delta))}.`);
  }
  if (row.transferred == null || row.transferred === MISSING) {
    items.push("Выплата за месяц не внесена. Долг — «нет данных», а не вся сумма начисления.");
  }
  if (row.utilities === MISSING) {
    items.push("Коммуналка и расход неизвестны: нет показаний предыдущего месяца. В итог они не входят.");
  }
  if (row.complete === "неполная") {
    items.push("Коммуналка неполная: в сумму вошли только известные начисления.");
  }

  addUnchanged(items, "Газ", row.gasReading, previous?.gasReading);
  addUnchanged(items, "Свет", row.lightReading, previous?.lightReading);
  addUnchanged(items, "Первый счётчик горячей воды", row.hot1, previous?.hot1);
  addUnchanged(items, "Второй счётчик горячей воды", row.hot2, previous?.hot2);
  addUnchanged(items, "Первый счётчик холодной воды", row.cold1, previous?.cold1);
  addUnchanged(items, "Второй счётчик холодной воды", row.cold2, previous?.cold2);
  return items;
}

export function formatMoney(value) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    if (value == null || value === "") return MISSING;
    return String(value);
  }
  return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 0 }).format(value);
}

export function formatQty(value) {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    if (value == null || value === "") return MISSING;
    return String(value);
  }
  return new Intl.NumberFormat("ru-RU", { maximumFractionDigits: 3 }).format(value);
}

export function formatReading(primary, secondary) {
  if (primary === NO_METER) return NO_METER;
  if (primary == null || primary === MISSING) return MISSING;
  if (typeof secondary === "number") return `${formatQty(primary)}+${formatQty(secondary)}`;
  return formatQty(primary);
}

function addUnchanged(items, label, current, prior) {
  if (typeof current === "number" && typeof prior === "number" && current === prior) {
    items.push(`${label}: показание не изменилось (${formatQty(current)}). Стоит проверить, снимали ли счётчик.`);
  }
}

function parseValue(value) {
  if (value == null) return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value).replace(/\s+/g, " ").trim();
  if (!text) return null;
  const numeric = text.replace(/\s/g, "").replace(",", ".");
  if (/^-?\d+(\.\d+)?$/.test(numeric)) return Number(numeric);
  return text;
}

function textOrNull(value) {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  return text || null;
}

function mapColumns(headerRow) {
  const columns = {};
  headerRow.forEach((cell, index) => {
    const text = normalize(cell);
    if (!text) return;
    for (const [field, needle] of HEADER_RULES) {
      if (columns[field] == null && text === needle) columns[field] = index;
    }
  });
  return columns;
}

function readMatrix(XLSX, buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const isZip = bytes.length > 1 && bytes[0] === 0x50 && bytes[1] === 0x4b;
  if (!isZip && looksLikeText(bytes)) {
    let text = new TextDecoder("utf-8").decode(bytes);
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    const firstLine = text.split(/\r?\n/, 1)[0] || "";
    const commas = (firstLine.match(/,/g) || []).length;
    const semicolons = (firstLine.match(/;/g) || []).length;
    const workbook = XLSX.read(text, { type: "string", FS: semicolons > commas ? ";" : ",", raw: true });
    return firstSheet(XLSX, workbook);
  }
  const workbook = XLSX.read(bytes, { type: "array" });
  return firstSheet(XLSX, workbook);
}

function firstSheet(XLSX, workbook) {
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("В книге нет листов");
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: "", blankrows: false });
}

function looksLikeText(bytes) {
  const sample = bytes.subarray(0, Math.min(bytes.length, 240));
  for (const byte of sample) {
    if (byte === 0) return false;
  }
  return sample.length > 0;
}

function normalize(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}
