# -*- coding: utf-8 -*-
"""Плоский файл для CRM. Суммы берутся из расчёта, а не набираются в странице."""

import csv
from pathlib import Path

from openpyxl import Workbook

from build_raschety import APTS, RESULT

OUT_DIR = Path(r"D:\Users\user\WebstormProjects\Test1NikiRent\crm\public\data")
HEADERS = [
    "Квартира",
    "Месяц",
    "Аренда",
    "Доп. расходы",
    "Коммуналка",
    "Перенос долга",
    "Итого к выплате",
    "Фин. учет",
    "Перечислено",
    "Долг",
    "Полнота",
    "Газ показание",
    "Газ расход",
    "Газ начислено",
    "Свет показание",
    "Свет расход",
    "Свет начислено",
    "Горячая 1",
    "Горячая 2",
    "Горячая расход",
    "Горячая начислено",
    "Холодная 1",
    "Холодная 2",
    "Холодная расход",
    "Холодная начислено",
    "Горводснабж начислено",
]


def primary(mode, value):
    if mode == "нет счётчика":
        return "нет счётчика"
    if value is None:
        return "нет данных"
    return value


def april_flow(mode):
    if mode == "нет счётчика":
        return "нет счётчика"
    return "нет данных"


def row_for(apt, month):
    found = {res: (mode, a1, a2, m1, m2) for res, mode, a1, a2, m1, m2, _note in apt["readings"]}
    gas = found["Газ"]
    light = found["Свет"]
    hot = found["Горячая вода"]
    cold = found["Холодная вода"]
    calc = RESULT[apt["name"]]
    if month == "апрель":
        money = [
            apt["rent_apr"],
            apt["extra_apr"],
            "нет данных",
            "нет данных",
            calc["apr_total"],
            apt["rent_apr"],
            apt["paid_apr"],
            calc["apr_debt"],
            "нет данных",
        ]
        flows = {
            "Газ": april_flow(gas[0]),
            "Свет": april_flow(light[0]),
            "Горячая вода": april_flow(hot[0]),
            "Холодная вода": april_flow(cold[0]),
            "Горводснабж": april_flow(hot[0]),
        }
        readings = (gas[1], light[1], hot[1], hot[2], cold[1], cold[2])
    else:
        money = [
            apt["rent_may"],
            apt["extra_may"],
            calc["util"],
            calc["apr_debt"],
            calc["payout"],
            calc["fin"],
            "нет данных",
            "нет данных",
            "неполная" if calc["incomplete"] else "полная",
        ]
        readings = (gas[3], light[3], hot[3], hot[4], cold[3], cold[4])
        flows = calc["cons"]
        charges = calc["charges"]
        return [
            apt["name"],
            month,
            *money,
            primary(gas[0], readings[0]),
            flows["Газ"],
            charges["Газ"],
            primary(light[0], readings[1]),
            flows["Свет"],
            charges["Свет"],
            primary(hot[0], readings[2]),
            readings[3],
            flows["Горячая вода"],
            charges["Горячая вода"],
            primary(cold[0], readings[4]),
            readings[5],
            flows["Холодная вода"],
            charges["Холодная вода"],
            charges["Горводснабж"],
        ]

    charges = flows
    return [
        apt["name"],
        month,
        *money,
        primary(gas[0], readings[0]),
        charges["Газ"],
        charges["Газ"],
        primary(light[0], readings[1]),
        charges["Свет"],
        charges["Свет"],
        primary(hot[0], readings[2]),
        readings[3],
        charges["Горячая вода"],
        charges["Горячая вода"],
        primary(cold[0], readings[4]),
        readings[5],
        charges["Холодная вода"],
        charges["Холодная вода"],
        charges["Горводснабж"],
    ]


def main():
    rows = []
    for apt in APTS:
        rows.append(row_for(apt, "апрель"))
        rows.append(row_for(apt, "май"))
    assert len(rows) == 10
    assert all(len(row) == len(HEADERS) for row in rows)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "данные"
    sheet.append(HEADERS)
    for row in rows:
        sheet.append(["" if value is None else value for value in row])
    xlsx_path = OUT_DIR / "kommunalka.xlsx"
    book.save(xlsx_path)

    csv_path = OUT_DIR / "kommunalka.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADERS)
        for row in rows:
            writer.writerow(["" if value is None else value for value in row])
    print(xlsx_path)
    print(csv_path)


if __name__ == "__main__":
    main()
