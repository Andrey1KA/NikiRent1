import { readFileSync } from "node:fs";
import XLSX from "xlsx";
import { parseWorkbook, rowsForMonth, sumField, warningsFor } from "../lib/rent.mjs";

const buffer = readFileSync(new URL("../public/data/kommunalka.xlsx", import.meta.url));
const rows = parseWorkbook(XLSX, buffer);
const may = rowsForMonth(rows, "май");
const april = rowsForMonth(rows, "апрель");

function byName(list, name) {
  const row = list.find((item) => item.apartment === name);
  if (!row) throw new Error(`нет строки ${name}`);
  return row;
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) throw new Error(`${label}: ${actual} !== ${expected}`);
}

assertEqual(may.length, 5, "май строк");
assertEqual(sumField(may, "rent"), 66200, "аренда мая");
assertEqual(sumField(may, "utilities"), 6441, "коммуналка мая");
assertEqual(sumField(may, "payout"), 70541, "выплата мая");
assertEqual(sumField(may, "fin"), 72641, "фин. учет мая");
assertEqual(sumField(april, "rent"), 67200, "аренда апреля");
assertEqual(sumField(april, "utilities"), "нет данных", "коммуналка апреля");
assertEqual(sumField(april, "payout"), 67200, "выплата апреля");

const lp10 = byName(may, "ЛП 10");
assertEqual(lp10.utilities, 1535, "ЛП 10 коммуналка");
assertEqual(lp10.payout, 13535, "ЛП 10 выплата");
assertEqual(lp10.gasUse, 5, "ЛП 10 газ");
assertEqual(lp10.hotUse, 5, "ЛП 10 горячая");
assertEqual(lp10.coldUse, 8, "ЛП 10 холодная");

const lp20 = byName(may, "ЛП 20");
assertEqual(lp20.utilities, 960, "ЛП 20 коммуналка");
assertEqual(lp20.payout, 11660, "ЛП 20 выплата");
assertEqual(lp20.fin, 13960, "ЛП 20 фин");
assertEqual(lp20.extra, 2300, "ЛП 20 доп");
assertEqual(lp20.gasReading, "нет счётчика", "ЛП 20 газ");

const trn = byName(may, "ТРН 26/6");
assertEqual(trn.utilities, 1253, "ТРН коммуналка");
assertEqual(trn.payout, 13653, "ТРН выплата");
assertEqual(trn.fin, 14453, "ТРН фин");
assertEqual(trn.gasReading, "нет данных", "ТРН газ");
assertEqual(trn.complete, "неполная", "ТРН полнота");
assertEqual(trn.carry, -800, "ТРН перенос");
assertEqual(trn.debt, "нет данных", "ТРН долг мая");

const lenina = byName(may, "Ленина 108");
assertEqual(lenina.utilities, 837, "Ленина коммуналка");
assertEqual(lenina.payout, 14837, "Ленина выплата");

const vlk = byName(may, "ВЛКСМ 56");
assertEqual(vlk.utilities, 1856, "ВЛКСМ коммуналка");
assertEqual(vlk.payout, 16856, "ВЛКСМ выплата");
assertEqual(vlk.fin, 15856, "ВЛКСМ фин");
assertEqual(vlk.carry, 1000, "ВЛКСМ перенос");
assertEqual(vlk.gasUse, 0, "ВЛКСМ газ расход");
assertEqual(vlk.hot1, 316, "ВЛКСМ горячая 1");
assertEqual(vlk.hot2, 10, "ВЛКСМ горячая 2");

const aprilVlk = byName(april, "ВЛКСМ 56");
assertEqual(aprilVlk.rent, 15000, "ВЛКСМ аренда апреля");
assertEqual(aprilVlk.debt, 1000, "ВЛКСМ долг апреля");
const aprilTrn = byName(april, "ТРН 26/6");
assertEqual(aprilTrn.debt, -800, "ТРН долг апреля");

for (const row of may) {
  for (const field of ["gasUse", "lightUse", "hotUse", "coldUse"]) {
    if (typeof row[field] === "number" && row[field] < 0) throw new Error(`${row.apartment} отрицательный расход`);
  }
}

const vlkWarnings = warningsFor(vlk, aprilVlk).join(" | ");
if (!vlkWarnings.includes("Газ: показание не изменилось")) throw new Error("нет предупреждения по газу ВЛКСМ");
if (!vlkWarnings.includes("Первый счётчик горячей воды")) throw new Error("нет предупреждения по горячей ВЛКСМ");
const trnWarnings = warningsFor(trn, aprilTrn).join(" | ");
if (!trnWarnings.includes("Нет показания: газ")) throw new Error("нет предупреждения по газу ТРН");
if (!trnWarnings.includes("неполная")) throw new Error("нет пометки неполной коммуналки");

const csv = readFileSync(new URL("../public/data/kommunalka.csv", import.meta.url));
const fromCsv = parseWorkbook(XLSX, csv);
assertEqual(sumField(rowsForMonth(fromCsv, "май"), "payout"), 70541, "csv выплата мая");

console.log("ok", rows.length, "строк");
