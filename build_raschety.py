# -*- coding: utf-8 -*-
"""Сборка raschety-mai.xlsx. Исходный test1Data.xlsx только читается."""

from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.page import PageMargins

SRC = r"D:\Users\user\WebstormProjects\Test1NikiRent\test1Data.xlsx"
OUT = r"D:\Users\user\WebstormProjects\Test1NikiRent\raschety-mai.xlsx"

NAVY = "1F4E79"
FILLS = {
    "title": PatternFill("solid", fgColor="1F4E79"),
    "section": PatternFill("solid", fgColor="2E75B6"),
    "header": PatternFill("solid", fgColor="D6DCE4"),
    "input": PatternFill("solid", fgColor="FFF2CC"),
    "formula": PatternFill("solid", fgColor="E2EFDA"),
    "text": PatternFill("solid", fgColor="FFFFFF"),
    "warn": PatternFill("solid", fgColor="FCE4D6"),
    "note": PatternFill("solid", fgColor="DEEBF7"),
    "bad": PatternFill("solid", fgColor="F8CBAD"),
}
FONTS = {
    "title": Font(name="Calibri", size=18, bold=True, color="FFFFFF"),
    "section": Font(name="Calibri", size=13, bold=True, color="FFFFFF"),
    "header": Font(name="Calibri", size=11, bold=True, color="1F4E79"),
    "text": Font(name="Calibri", size=11, color="000000"),
    "bold": Font(name="Calibri", size=11, bold=True, color="000000"),
    "white": Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
    "small": Font(name="Calibri", size=10, italic=True, color="595959"),
}
THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
WRAP = Alignment(wrap_text=True, vertical="center")
LEFT = Alignment(vertical="center", wrap_text=True)
CENTER = Alignment(vertical="center", horizontal="center", wrap_text=True)

MONEY = '#,##0;-#,##0;0'
TAR = '0.00'
QTY = '0.###'


def D(x):
    return Decimal(str(x))


def excel_round(x):
    return int(D(x).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


TARIFF = {
    "Газ": D("9.5"),
    "Свет": D("3.83"),
    "Горячая вода": D("25.44"),
    "Холодная вода": D("25.44") + D("27.14"),
    "Горводснабж": D("107.4"),
}

APTS = [
    {
        "name": "ЛП 10",
        "rent_apr": 12000,
        "rent_may": 12000,
        "rent_note": "Шапка: в мае 12 000 + к/у. В апреле в ячейке тоже 12 000, другой ставки в файле нет.",
        "extra_apr": 0,
        "extra_may": 0,
        "extra_note": "В исходнике пусто. Пустой доп. расход = 0.",
        "paid_apr": 12000,
        "readings": [
            ("Газ", "одним числом", 10, None, 15, None, "Одно показание. Май уже стоял в исходнике как образец."),
            ("Свет", "одним числом", 2300, None, 2405, None, "Май уже стоял в исходнике как образец."),
            ("Горячая вода", "одним числом", 7, None, 10, 2, "В апреле одно число 7, в мае два счётчика 10 и 2. По приборам сопоставить нельзя: расход (10+2)−7, как в образце исходника."),
            ("Холодная вода", "одним числом", 10, None, 15, 3, "В апреле одно число 10, в мае 15 и 3. Расход (15+3)−10, как в образце исходника."),
        ],
    },
    {
        "name": "ЛП 20",
        "rent_apr": 13000,
        "rent_may": 13000,
        "rent_note": "Шапка: до 5-го числа, 13+.",
        "extra_apr": 0,
        "extra_may": 2300,
        "extra_note": "Май: 2 300 в исходнике. Вычитается из выплаты собственнику, в фин. учет второй раз не входит.",
        "paid_apr": 13000,
        "readings": [
            ("Газ", "нет счётчика", None, None, None, None, "Ни в апреле, ни в мае газа нет. Это не пропуск показания."),
            ("Свет", "одним числом", 3343, None, 3400, None, ""),
            ("Горячая вода", "одним числом", 196, None, 200, None, ""),
            ("Холодная вода", "одним числом", 156, None, 160, None, ""),
        ],
    },
    {
        "name": "ТРН 26/6",
        "rent_apr": 13200,
        "rent_may": 13200,
        "rent_note": "В шапке суммы нет. Оставил 13 200 из ячейки исходника.",
        "extra_apr": 0,
        "extra_may": 0,
        "extra_note": "В исходнике пусто. Пустой доп. расход = 0.",
        "paid_apr": 14000,
        "readings": [
            ("Газ", "нет показания", 35, None, None, None, "В апреле показание 35 есть. В задании на май газа нет. Ноль не подставлял."),
            ("Свет", "одним числом", 7402, None, 7501, None, ""),
            ("Горячая вода", "одним числом", 196, None, 201, None, ""),
            ("Холодная вода", "одним числом", 553, None, 557, None, ""),
        ],
    },
    {
        "name": "Ленина 108",
        "rent_apr": 14000,
        "rent_may": 14000,
        "rent_note": "Шапка: 14+, до 30-го числа.",
        "extra_apr": 0,
        "extra_may": 0,
        "extra_note": "В исходнике пусто. Пустой доп. расход = 0.",
        "paid_apr": 14000,
        "readings": [
            ("Газ", "одним числом", 500, None, 503, None, ""),
            ("Свет", "одним числом", 12552, None, 12701, None, ""),
            ("Горячая вода", "одним числом", 186, None, 187, None, ""),
            ("Холодная вода", "одним числом", 337, None, 339, None, ""),
        ],
    },
    {
        "name": "ВЛКСМ 56",
        "rent_apr": 15000,
        "rent_may": 14000,
        "rent_note": "Шапка: с ноября по апрель 15+, с мая по октябрь 14+. В исходнике в апреле стояло 14 000 — это ошибка. Перечислено 14 000 не менял.",
        "extra_apr": 0,
        "extra_may": 0,
        "extra_note": "В исходнике пусто. Пустой доп. расход = 0.",
        "paid_apr": 14000,
        "readings": [
            ("Газ", "одним числом", 438, None, 438, None, "Май совпал с апрелем. Расход 0. Это не пустая ячейка, показание есть."),
            ("Свет", "одним числом", 7553, None, 7650, None, ""),
            ("Горячая вода", "два счётчика", 316, 2, 316, 10, "Апрель в исходнике =316+2. Май из задания 316+10. Расход по приборам: (316−316)+(10−2)."),
            ("Холодная вода", "два счётчика", 397, 57, 397, 65, "Апрель в исходнике =397+57. Май из задания 397+65. Расход: (397−397)+(65−57)."),
        ],
    },
]


def consumption_of(mode, a1, a2, m1, m2):
    if mode == "нет счётчика":
        return "нет счётчика"
    if mode == "нет показания" or m1 is None:
        return "нет данных"
    if mode == "два счётчика":
        if None in (a1, a2, m1, m2):
            return "нет данных"
        return (m1 - a1) + (m2 - a2)
    if m2 is None:
        return m1 - a1
    return (m1 + m2) - a1


def compute():
    out = {}
    for apt in APTS:
        cons = {}
        charges = {}
        for res, mode, a1, a2, m1, m2, _note in apt["readings"]:
            c = consumption_of(mode, a1, a2, m1, m2)
            cons[res] = c
            if isinstance(c, str):
                charges[res] = c
            else:
                charges[res] = excel_round(TARIFF[res] * D(c))
        hot = cons["Горячая вода"]
        cons["Горводснабж"] = hot
        if isinstance(hot, str):
            charges["Горводснабж"] = hot
        else:
            charges["Горводснабж"] = excel_round(TARIFF["Горводснабж"] * D(hot))
        known = [v for v in charges.values() if isinstance(v, int)]
        incomplete = any(v == "нет данных" for v in charges.values())
        util = sum(known)
        apr_total = apt["rent_apr"] - apt["extra_apr"]
        apr_debt = apr_total - apt["paid_apr"]
        payout = apt["rent_may"] + util - apt["extra_may"] + apr_debt
        fin = apt["rent_may"] + util
        out[apt["name"]] = {
            "cons": cons,
            "charges": charges,
            "util": util,
            "incomplete": incomplete,
            "apr_total": apr_total,
            "apr_debt": apr_debt,
            "apr_fin": apt["rent_apr"],
            "payout": payout,
            "fin": fin,
        }
    return out


RESULT = compute()

# Цифры, которые я зафиксировал вручную по правилам листа. Если входные данные
# в этом файле разъедутся с ними, сборка остановится.
assert RESULT["ЛП 10"]["charges"] == {
    "Газ": 48, "Свет": 402, "Горячая вода": 127, "Холодная вода": 421, "Горводснабж": 537,
}
assert RESULT["ЛП 10"]["util"] == 1535
assert RESULT["ЛП 10"]["payout"] == 13535
assert RESULT["ЛП 10"]["fin"] == 13535
assert RESULT["ЛП 10"]["apr_debt"] == 0
assert RESULT["ЛП 20"]["charges"]["Газ"] == "нет счётчика"
assert RESULT["ЛП 20"]["charges"]["Свет"] == 218
assert RESULT["ЛП 20"]["charges"]["Горячая вода"] == 102
assert RESULT["ЛП 20"]["charges"]["Холодная вода"] == 210
assert RESULT["ЛП 20"]["charges"]["Горводснабж"] == 430
assert RESULT["ЛП 20"]["util"] == 960
assert RESULT["ЛП 20"]["payout"] == 11660
assert RESULT["ЛП 20"]["fin"] == 13960
assert RESULT["ЛП 20"]["incomplete"] is False
assert RESULT["ТРН 26/6"]["charges"]["Газ"] == "нет данных"
assert RESULT["ТРН 26/6"]["charges"]["Свет"] == 379
assert RESULT["ТРН 26/6"]["charges"]["Горячая вода"] == 127
assert RESULT["ТРН 26/6"]["charges"]["Холодная вода"] == 210
assert RESULT["ТРН 26/6"]["charges"]["Горводснабж"] == 537
assert RESULT["ТРН 26/6"]["util"] == 1253
assert RESULT["ТРН 26/6"]["incomplete"] is True
assert RESULT["ТРН 26/6"]["apr_debt"] == -800
assert RESULT["ТРН 26/6"]["payout"] == 13653
assert RESULT["ТРН 26/6"]["fin"] == 14453
assert RESULT["Ленина 108"]["charges"] == {
    "Газ": 29, "Свет": 571, "Горячая вода": 25, "Холодная вода": 105, "Горводснабж": 107,
}
assert RESULT["Ленина 108"]["util"] == 837
assert RESULT["Ленина 108"]["payout"] == 14837
assert RESULT["Ленина 108"]["fin"] == 14837
assert RESULT["ВЛКСМ 56"]["cons"]["Газ"] == 0
assert RESULT["ВЛКСМ 56"]["cons"]["Горячая вода"] == 8
assert RESULT["ВЛКСМ 56"]["cons"]["Холодная вода"] == 8
assert RESULT["ВЛКСМ 56"]["cons"]["Свет"] == 97
assert RESULT["ВЛКСМ 56"]["charges"] == {
    "Газ": 0, "Свет": 372, "Горячая вода": 204, "Холодная вода": 421, "Горводснабж": 859,
}
assert RESULT["ВЛКСМ 56"]["util"] == 1856
assert RESULT["ВЛКСМ 56"]["apr_debt"] == 1000
assert RESULT["ВЛКСМ 56"]["payout"] == 16856
assert RESULT["ВЛКСМ 56"]["fin"] == 15856
assert sum(RESULT[a["name"]]["payout"] for a in APTS) == 70541
assert sum(RESULT[a["name"]]["fin"] for a in APTS) == 72641
assert sum(a["rent_may"] for a in APTS) == 66200
assert sum(RESULT[a["name"]]["util"] for a in APTS) == 6441
assert sum(RESULT[a["name"]]["apr_debt"] for a in APTS) == 200


def apply(cell, kind, fmt=None, align=None):
    cell.font = FONTS["section"] if kind == "section" else FONTS["title"] if kind == "title" else FONTS["header"] if kind == "header" else FONTS["text"]
    cell.fill = FILLS[kind]
    cell.alignment = align or LEFT
    if kind != "title":
        cell.border = THIN
    if fmt:
        cell.number_format = fmt
    return cell


def write(ws, addr, value, kind="text", fmt=None, align=None):
    cell = ws[addr]
    cell.value = value
    return apply(cell, kind, fmt, align)


def comment(cell, text):
    cell.comment = Comment(text, "NikiRent")
    cell.comment.width = 280
    cell.comment.height = 80


def cons_formula(row):
    mode = f"'Ввод'!C{row}"
    a1, a2, m1, m2 = (f"'Ввод'!{col}{row}" for col in "DEFG")
    return (
        f'IF({mode}="нет счётчика","нет счётчика",'
        f'IF({mode}="нет показания","нет данных",'
        f'IF({mode}="два счётчика",'
        f'IF(OR({a1}="",{a2}="",{m1}="",{m2}=""),"нет данных",({m1}-{a1})+({m2}-{a2})),'
        f'IF({m1}="","нет данных",IF({m2}="",{m1}-{a1},({m1}+{m2})-{a1})))))'
    )


def warn_formula(calc_row, input_row):
    mode = f"'Ввод'!C{input_row}"
    a1, a2, m1, m2 = (f"'Ввод'!{col}{input_row}" for col in "DEFG")
    cons = f"C{calc_row}"
    res = f"B{calc_row}"
    return (
        f'IF(OR({cons}="нет данных",{cons}="нет счётчика"),"",'
        f'IF({mode}="два счётчика",'
        f'IF(OR({m1}<{a1},{m2}<{a2}),"новое показание меньше предыдущего",'
        f'IF(AND({m1}={a1},{m2}={a2}),"оба показания не изменились",'
        f'IF({m1}={a1},"счётчик 1 не изменился",'
        f'IF({m2}={a2},"счётчик 2 не изменился","")))),'
        f'IF(AND(ISNUMBER({cons}),{cons}<0),"новое показание меньше предыдущего",'
        f'IF(AND({res}="Свет",{cons}>300),"свет больше 300 кВт·ч",'
        f'IF(AND({res}<>"Горводснабж",ISNUMBER({cons}),{cons}=0),"показание не изменилось","")))))'
    )


def build():
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)

    vvod = wb.active
    vvod.title = "Ввод"
    ras = wb.create_sheet("Расчет", 0)
    itog = wb.create_sheet("Итог", 0)
    grid = wb.create_sheet("Коммуналка", 1)
    notes = wb.create_sheet("Замечания")
    snap = wb.create_sheet("Исходник")

    build_vvod(vvod)
    addr = build_raschet(ras)
    build_itog(itog, addr)
    build_kommunalka(grid)
    build_notes(notes)
    build_snapshot(snap)

    vvod.sheet_properties.tabColor = "FFC000"
    ras.sheet_properties.tabColor = "5B9BD5"
    itog.sheet_properties.tabColor = "548235"
    grid.sheet_properties.tabColor = "7030A0"
    notes.sheet_properties.tabColor = "C65911"
    snap.sheet_properties.tabColor = "808080"
    wb.active = 0
    wb.save(OUT)
    return addr


def build_vvod(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(0.5, 0.5, 0.6, 0.6)
    ws.freeze_panes = "A7"
    widths = {"A": 24, "B": 22, "C": 18, "D": 16, "E": 16, "F": 16, "G": 16, "H": 78}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    write(ws, "A1", "Вводные данные", "title", align=LEFT)
    ws.merge_cells("A1:H1")
    ws.row_dimensions[1].height = 28
    write(ws, "A2", "Жёлтые ячейки можно менять: пересчёт на листах «Расчет» и «Итог» пойдёт сам. Холодная вода считается формулой, как в исходнике: 25,44 + 27,14.", "note", align=LEFT)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 32

    write(ws, "A4", "Тарифы", "section")
    ws.merge_cells("A4:D4")
    for col, text in enumerate(["Показатель", "Значение", "Ед.", "Откуда"], 1):
        write(ws, f"{get_column_letter(col)}5", text, "header", align=CENTER)
    tariffs = [
        (7, "Газ", 9.5, "руб/м³", "Колонка C исходника, апрель и май совпадают."),
        (8, "Свет", 3.83, "руб/кВт·ч", "Колонка C исходника."),
        (9, "Горячая вода", 25.44, "руб/м³", "Колонка C исходника."),
        (10, "Слагаемое холодной воды", 27.14, "руб/м³", "Второе слагаемое формулы исходника 25,44+27,14. Отдельным тарифом в файле не расшифровано."),
    ]
    for row, name, val, unit, src in tariffs:
        write(ws, f"A{row}", name, "text")
        cell = write(ws, f"B{row}", val, "input", TAR, CENTER)
        comment(cell, src)
        write(ws, f"C{row}", unit, "text", align=CENTER)
        write(ws, f"D{row}", src, "note")
    write(ws, "A11", "Холодная вода", "text")
    write(ws, "B11", "=B9+B10", "formula", TAR, CENTER)
    write(ws, "C11", "руб/м³", "text", align=CENTER)
    write(ws, "D11", "Формула исходника. Сам тариф не исправлял: оба месяца в файле совпадают, другого источника нет.", "note")
    write(ws, "A12", "Горводснабж", "text")
    cell = write(ws, "B12", 107.4, "input", TAR, CENTER)
    comment(cell, "Своего счётчика нет. Начисляется на расход горячей воды.")
    write(ws, "C12", "руб/м³ горячей", "text", align=CENTER)
    write(ws, "D12", "Своего счётчика нет. В мае исходника считалось как тариф × расход горячей воды.", "note")

    write(ws, "A14", "Аренда, руб.", "section")
    ws.merge_cells("A14:D14")
    for col, text in enumerate(["Квартира", "Апрель", "Май", "Комментарий"], 1):
        write(ws, f"{get_column_letter(col)}15", text, "header", align=CENTER)
    for i, apt in enumerate(APTS):
        row = 16 + i
        write(ws, f"A{row}", apt["name"], "text")
        c1 = write(ws, f"B{row}", apt["rent_apr"], "input", MONEY, CENTER)
        c2 = write(ws, f"C{row}", apt["rent_may"], "input", MONEY, CENTER)
        comment(c1, apt["rent_note"])
        comment(c2, apt["rent_note"])
        write(ws, f"D{row}", apt["rent_note"], "note")

    write(ws, "A22", "Доп. расходы, руб. Пустая ячейка исходника записана как 0.", "section")
    ws.merge_cells("A22:D22")
    for col, text in enumerate(["Квартира", "Апрель", "Май", "Комментарий"], 1):
        write(ws, f"{get_column_letter(col)}23", text, "header", align=CENTER)
    for i, apt in enumerate(APTS):
        row = 24 + i
        write(ws, f"A{row}", apt["name"], "text")
        write(ws, f"B{row}", apt["extra_apr"], "input", MONEY, CENTER)
        write(ws, f"C{row}", apt["extra_may"], "input", MONEY, CENTER)
        write(ws, f"D{row}", apt["extra_note"], "note")

    write(ws, "A30", "Перечислено собственнику, руб.", "section")
    ws.merge_cells("A30:D30")
    for col, text in enumerate(["Квартира", "Апрель", "Май", "Комментарий"], 1):
        write(ws, f"{get_column_letter(col)}31", text, "header", align=CENTER)
    for i, apt in enumerate(APTS):
        row = 32 + i
        write(ws, f"A{row}", apt["name"], "text")
        write(ws, f"B{row}", apt["paid_apr"], "input", MONEY, CENTER)
        write(ws, f"C{row}", "нет данных", "input", align=CENTER)
        write(ws, f"D{row}", "Апрель — факт из исходника. Май в исходнике пустой: это не ноль, выплату не выдумывал. Если вписать число, долг мая посчитается сам.", "note")

    write(ws, "A38", "Показания счётчиков", "section")
    ws.merge_cells("A38:H38")
    headers = ["Квартира", "Ресурс", "Как считать", "Апрель сч.1", "Апрель сч.2", "Май сч.1", "Май сч.2", "Комментарий"]
    for col, text in enumerate(headers, 1):
        write(ws, f"{get_column_letter(col)}39", text, "header", align=CENTER)
    ws.auto_filter.ref = "A39:H59"
    ws.auto_filter.ref = "A39:H59"
    for i, apt in enumerate(APTS):
        for j, (res, mode, a1, a2, m1, m2, note) in enumerate(apt["readings"]):
            row = 40 + i * 4 + j
            write(ws, f"A{row}", apt["name"], "text")
            write(ws, f"B{row}", res, "text")
            write(ws, f"C{row}", mode, "input", align=CENTER)
            for col, val in (("D", a1), ("E", a2), ("F", m1), ("G", m2)):
                if val is None:
                    write(ws, f"{col}{row}", None, "input", QTY, CENTER)
                else:
                    write(ws, f"{col}{row}", val, "input", QTY, CENTER)
            write(ws, f"H{row}", note, "note")
            ws.row_dimensions[row].height = 36 if note else 18

    ws.row_dimensions[14].height = 22
    ws.row_dimensions[22].height = 22
    ws.row_dimensions[30].height = 22
    ws.oddHeader.left.text = "NikiRent — ввод"
    ws.oddFooter.right.text = "Жёлтое — ввод, зелёное — формула"


def build_raschet(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(0.5, 0.5, 0.6, 0.6)
    ws.freeze_panes = "A4"
    for col, w in {"A": 16, "B": 18, "C": 18, "D": 14, "E": 12, "F": 16, "G": 36, "H": 40, "I": 62}.items():
        ws.column_dimensions[col].width = w

    write(ws, "A1", "Расчёт расхода и начислений", "title")
    ws.merge_cells("A1:I1")
    ws.row_dimensions[1].height = 28
    write(ws, "A2", "Расход = май − апрель. Пустое майское показание не превращается в ноль. Два счётчика вычитаются каждый из своего апреля. Начисление = ROUND(тариф × расход; 0). Горводснабж = ROUND(107,4 × расход горячей воды; 0).", "note")
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 34
    for col, text in enumerate(["Квартира", "Ресурс", "Расход", "Ед.", "Тариф", "Начислено", "Как посчитано", "Предупреждение", "Примечание"], 1):
        write(ws, f"{get_column_letter(col)}3", text, "header", align=CENTER)

    tariff_cell = {
        "Газ": "$B$7",
        "Свет": "$B$8",
        "Горячая вода": "$B$9",
        "Холодная вода": "$B$11",
        "Горводснабж": "$B$12",
    }
    unit = {
        "Газ": "м³",
        "Свет": "кВт·ч",
        "Горячая вода": "м³",
        "Холодная вода": "м³",
        "Горводснабж": "м³ гор.",
    }
    addr = {}
    calc_row = 4
    for i, apt in enumerate(APTS):
        hot_row = None
        for j, (res, _mode, *_rest) in enumerate(apt["readings"]):
            input_row = 40 + i * 4 + j
            write(ws, f"A{calc_row}", apt["name"], "text")
            write(ws, f"B{calc_row}", res, "text")
            write(ws, f"C{calc_row}", "=" + cons_formula(input_row), "formula", QTY, CENTER)
            write(ws, f"D{calc_row}", unit[res], "text", align=CENTER)
            write(ws, f"E{calc_row}", f"='Ввод'!{tariff_cell[res]}", "formula", TAR, CENTER)
            write(ws, f"F{calc_row}", f'=IF(OR(C{calc_row}="нет данных",C{calc_row}="нет счётчика"),C{calc_row},ROUND(E{calc_row}*C{calc_row},0))', "formula", MONEY, CENTER)
            write(ws, f"G{calc_row}", f'=IF(ISNUMBER(C{calc_row}),C{calc_row}&" × "&E{calc_row}&" → "&F{calc_row},C{calc_row})', "formula")
            write(ws, f"H{calc_row}", "=" + warn_formula(calc_row, input_row), "formula")
            write(ws, f"I{calc_row}", f"='Ввод'!H{input_row}", "formula")
            addr[(apt["name"], res)] = calc_row
            if res == "Горячая вода":
                hot_row = calc_row
            ws.row_dimensions[calc_row].height = 32
            calc_row += 1
        write(ws, f"A{calc_row}", apt["name"], "text")
        write(ws, f"B{calc_row}", "Горводснабж", "text")
        write(ws, f"C{calc_row}", f'=IF(OR(C{hot_row}="нет данных",C{hot_row}="нет счётчика"),C{hot_row},C{hot_row})', "formula", QTY, CENTER)
        write(ws, f"D{calc_row}", unit["Горводснабж"], "text", align=CENTER)
        write(ws, f"E{calc_row}", f"='Ввод'!{tariff_cell['Горводснабж']}", "formula", TAR, CENTER)
        write(ws, f"F{calc_row}", f'=IF(OR(C{calc_row}="нет данных",C{calc_row}="нет счётчика"),C{calc_row},ROUND(E{calc_row}*C{calc_row},0))', "formula", MONEY, CENTER)
        write(ws, f"G{calc_row}", f'=IF(ISNUMBER(C{calc_row}),C{calc_row}&" × "&E{calc_row}&" → "&F{calc_row},C{calc_row})', "formula")
        write(ws, f"H{calc_row}", '=""', "formula")
        write(ws, f"I{calc_row}", "Подогрев воды: своего счётчика нет, база — расход горячей воды.", "note")
        addr[(apt["name"], "Горводснабж")] = calc_row
        ws.row_dimensions[calc_row].height = 32
        calc_row += 1
        calc_row += 1  # разделитель
    ws.auto_filter.ref = f"A3:I{calc_row - 1}"
    ws.oddHeader.left.text = "NikiRent — расчёт"
    return addr


def build_itog(ws, addr):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(0.45, 0.45, 0.55, 0.55)
    ws.freeze_panes = "B6"
    ws.column_dimensions["A"].width = 38
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 62
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 48
    ws.print_title_rows = "1:2"
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "NikiRent"
    ws.oddFooter.right.text = "Итог к выплате = аренда + коммуналка − доп. расходы + долг апреля"
    ws.oddFooter.left.text = "Фин. учет = аренда + коммуналка"

    write(ws, "A1", "NikiRent — квартиры, коммуналка и выплата собственникам", "title")
    ws.merge_cells("A1:H1")
    write(ws, "A2", "Итого к выплате = аренда + коммуналка − доп. расходы + долг апреля (недоплата прибавляется, переплата вычитается). Фин. учет = аренда + коммуналка: это расход компании, доп. расход в него второй раз не входит. Пустое показание и пустая выплата мая — «нет данных», не ноль. Апрельская коммуналка не считалась: нет показаний марта.", "note")
    ws.merge_cells("A2:H2")

    names = [a["name"] for a in APTS]
    # колонки итога B-F соответствуют квартирам 0-4
    col_of = {name: get_column_letter(2 + i) for i, name in enumerate(names)}

    def section(row, text):
        write(ws, f"A{row}", text, "section")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        ws.row_dimensions[row].height = 22

    def headers(row):
        write(ws, f"A{row}", "Показатель", "header")
        for i, name in enumerate(names):
            write(ws, f"{get_column_letter(2+i)}{row}", name, "header", align=CENTER)
        write(ws, f"G{row}", "Всего", "header", align=CENTER)
        write(ws, f"H{row}", "Пояснение", "header", align=CENTER)

    def money_row(row, label, kind_for_value="formula"):
        write(ws, f"A{row}", label, "bold" if False else "text")
        ws[f"A{row}"].font = FONTS["bold"]
        ws[f"A{row}"].fill = FILLS["text"]
        ws[f"A{row}"].border = THIN

    section(4, "Май — что начислено собственнику и что лежит в фин. учете")
    headers(5)

    # 6 аренда, 7 коммуналка, 8 полнота, 9 доп, 10 перенос, 11 итого, 12 фин, 13 перечислено, 14 долг
    labels_may = [
        (6, "Аренда"),
        (7, "Коммуналка"),
        (8, "Полнота коммуналки"),
        (9, "Доп. расходы"),
        (10, "Перенос долга апреля"),
        (11, "Итого к выплате"),
        (12, "Фин. учет"),
        (13, "Перечислено"),
        (14, "Долг / переплата"),
    ]
    for row, label in labels_may:
        write(ws, f"A{row}", label, "text")
        ws[f"A{row}"].font = FONTS["bold"] if row in (11, 12) else FONTS["text"]

    for i, apt in enumerate(APTS):
        col = get_column_letter(2 + i)
        vrow = 16 + i
        erow = 24 + i
        prow = 32 + i
        write(ws, f"{col}6", f"='Ввод'!C{vrow}", "formula", MONEY, CENTER)
        write(ws, f"{col}7", f"=SUMIF('Расчет'!$A$4:$A$40,{col}$5,'Расчет'!$F$4:$F$40)", "formula", MONEY, CENTER)
        write(ws, f"{col}8", f'=IF(COUNTIFS(\'Расчет\'!$A$4:$A$40,{col}$5,\'Расчет\'!$F$4:$F$40,"нет данных")>0,"неполная","полная")', "formula", align=CENTER)
        write(ws, f"{col}9", f"='Ввод'!C{erow}", "formula", MONEY, CENTER)
        write(ws, f"{col}10", f"={col}23", "formula", MONEY, CENTER)
        write(ws, f"{col}11", f"={col}6+{col}7-{col}9+{col}10", "formula", MONEY, CENTER)
        write(ws, f"{col}12", f"={col}6+{col}7", "formula", MONEY, CENTER)
        write(ws, f"{col}13", f"='Ввод'!C{prow}", "formula", align=CENTER)
        write(ws, f"{col}14", f'=IF(OR({col}13="",{col}13="нет данных"),"нет данных",{col}11-{col}13)', "formula", MONEY, CENTER)
        for r, h in ((6, 20), (7, 20), (11, 24), (12, 24)):
            ws.row_dimensions[r].height = h

    for row, fmt in ((6, MONEY), (7, MONEY), (9, MONEY), (10, MONEY), (11, MONEY), (12, MONEY)):
        write(ws, f"G{row}", f"=SUM(B{row}:F{row})", "formula", fmt, CENTER)
    write(ws, "G8", '=COUNTIF(B8:F8,"неполная")&" неполная"', "formula", align=CENTER)
    write(ws, "G13", '=IF(COUNT(B13:F13)=5,SUM(B13:F13),"нет данных")', "formula", align=CENTER)
    write(ws, "G14", '=IF(COUNT(B14:F14)=5,SUM(B14:F14),"нет данных")', "formula", align=CENTER)

    write(ws, "H6", "Ставки из шапки и ячеек. У ВЛКСМ май — 14 000.", "note")
    write(ws, "H7", "Сумма начислений по газу, свету, воде и подогреву. У ТРН без газа.", "note")
    write(ws, "H8", "«Неполная» — по какому-то ресурсу нет показания. «Нет счётчика» полноту не ломает.", "note")
    write(ws, "H9", "Уменьшают выплату собственнику. В фин. учете этой строкой повторно не сидят.", "note")
    write(ws, "H10", "Долг апреля: плюс — недоплатили собственнику, минус — переплата. В май эта сумма прибавляется.", "note")
    write(ws, "H11", "Аренда + коммуналка − доп. расходы + перенос. У ТРН цифра неполная: нет газа.", "note")
    write(ws, "H12", "Аренда + коммуналка. Расход компании за май. С выплатой совпадает, только если нет доп. расходов и переноса.", "note")
    write(ws, "H13", "В задании и в исходнике выплаты за май нет.", "note")
    write(ws, "H14", "Пока выплаты нет, долг не равен всему итогу. Ноль вместо пустой выплаты не ставил.", "note")
    for r in range(6, 15):
        ws.row_dimensions[r].height = 22
    ws.row_dimensions[11].height = 26
    ws.row_dimensions[12].height = 26

    section(16, "Апрель — факт расчёта с собственником. Коммуналка неизвестна: нет марта")
    headers(17)
    apr_labels = [
        (18, "Аренда"),
        (19, "Коммуналка"),
        (20, "Доп. расходы"),
        (21, "Итого начислено"),
        (22, "Перечислено"),
        (23, "Долг / переплата"),
        (24, "Фин. учет"),
    ]
    for row, label in apr_labels:
        write(ws, f"A{row}", label, "text")
        ws[f"A{row}"].font = FONTS["bold"] if row in (21, 23, 24) else FONTS["text"]
    for i in range(5):
        col = get_column_letter(2 + i)
        vrow = 16 + i
        erow = 24 + i
        prow = 32 + i
        write(ws, f"{col}18", f"='Ввод'!B{vrow}", "formula", MONEY, CENTER)
        write(ws, f"{col}19", "нет данных", "bad", align=CENTER)
        write(ws, f"{col}20", f"='Ввод'!B{erow}", "formula", MONEY, CENTER)
        write(ws, f"{col}21", f"={col}18-{col}20", "formula", MONEY, CENTER)
        write(ws, f"{col}22", f"='Ввод'!B{prow}", "formula", MONEY, CENTER)
        write(ws, f"{col}23", f"={col}21-{col}22", "formula", MONEY, CENTER)
        write(ws, f"{col}24", f"={col}18", "formula", MONEY, CENTER)
    for row in (18, 20, 21, 22, 23, 24):
        write(ws, f"G{row}", f"=SUM(B{row}:F{row})", "formula", MONEY, CENTER)
    write(ws, "G19", "нет данных", "bad", align=CENTER)
    write(ws, "H18", "ВЛКСМ: 15 000 вместо 14 000 из ячейки. Ноябрь–апрель в шапке указаны как 15+.", "note")
    write(ws, "H19", "Расход апреля посчитать нельзя. В исходнике здесь был 0 из-за SUM по пустым ячейкам. Ноль убрал.", "note")
    write(ws, "H20", "Доп. расходов в апреле нет.", "note")
    write(ws, "H21", "Только аренда минус доп. расходы. Коммуналка в итог апреля не входит.", "note")
    write(ws, "H22", "Факт перевода из исходника. Не исправлял.", "note")
    write(ws, "H23", "Итого − перечислено. ТРН −800: переплата. ВЛКСМ +1 000: после правки аренды перевод 14 000 не закрывает 15 000.", "note")
    write(ws, "H24", "Известная часть расхода компании — аренда. Коммуналка апреля сюда не подставлена.", "note")
    for r in range(18, 25):
        ws.row_dimensions[r].height = 22

    section(26, "Расход за май")
    headers(27)
    resources = ["Газ", "Свет", "Горячая вода", "Холодная вода"]
    for k, res in enumerate(resources):
        row = 28 + k
        write(ws, f"A{row}", res, "text")
        for i, name in enumerate(names):
            col = get_column_letter(2 + i)
            src = addr[(name, res)]
            write(ws, f"{col}{row}", f"='Расчет'!C{src}", "formula", QTY, CENTER)
        write(ws, f"G{row}", f'=IF(COUNT(B{row}:F{row})=5,SUM(B{row}:F{row}),"не всё есть")', "formula", QTY, CENTER)
    write(ws, "A32", "Горячая + холодная, м³", "text")
    ws["A32"].font = FONTS["bold"]
    for i in range(5):
        col = get_column_letter(2 + i)
        write(ws, f"{col}32", f'=IF(AND(ISNUMBER({col}30),ISNUMBER({col}31)),{col}30+{col}31,"нет данных")', "formula", QTY, CENTER)
    write(ws, "G32", '=IF(COUNT(B32:F32)=5,SUM(B32:F32),"не всё есть")', "formula", QTY, CENTER)
    write(ws, "A33", "Свет больше 300 кВт·ч", "text")
    write(ws, "A34", "Вода больше 20 м³", "text")
    for i in range(5):
        col = get_column_letter(2 + i)
        write(ws, f"{col}33", f'=IF(AND(ISNUMBER({col}29),{col}29>300),"да","нет")', "formula", align=CENTER)
        write(ws, f"{col}34", f'=IF(AND(ISNUMBER({col}32),{col}32>20),"да","нет")', "formula", align=CENTER)
    write(ws, "G33", '=IF(COUNTIF(B33:F33,"да")>0,"да","нет")', "formula", align=CENTER)
    write(ws, "G34", '=IF(COUNTIF(B34:F34,"да")>0,"да","нет")', "formula", align=CENTER)
    write(ws, "H28", "У ТРН показания нет. У ВЛКСМ расход 0: 438 и 438.", "note")
    write(ws, "H29", "Порог проверки 300 кВт·ч. Превышение само по себе не ошибка расчёта.", "note")
    write(ws, "H32", "Порог проверки 20 м³ на горячую и холодную вместе.", "note")
    write(ws, "H33", "Ни одна квартира порог не переходит.", "note")
    write(ws, "H34", "Ни одна квартира порог не переходит. У ВЛКСМ 16 м³ — ближе всех, но ниже 20.", "note")
    for r in range(28, 35):
        ws.row_dimensions[r].height = 20

    section(36, "Начислено за май, руб.")
    headers(37)
    charge_res = ["Газ", "Свет", "Горячая вода", "Холодная вода", "Горводснабж"]
    for k, res in enumerate(charge_res):
        row = 38 + k
        write(ws, f"A{row}", res, "text")
        for i, name in enumerate(names):
            col = get_column_letter(2 + i)
            src = addr[(name, res)]
            write(ws, f"{col}{row}", f"='Расчет'!F{src}", "formula", MONEY, CENTER)
        write(ws, f"G{row}", f'=IF(COUNT(B{row}:F{row})=5,SUM(B{row}:F{row}),"не всё есть")', "formula", MONEY, CENTER)
    write(ws, "H38", "ТРН: нет данных, в сумму коммуналки не входит.", "note")
    write(ws, "H42", "Подогрев на кубы горячей воды, не отдельное показание.", "note")

    section(44, "Что проверить руководителю")
    checks_text = [
        (45, f'=IF(\'Расчет\'!C{addr[("ТРН 26/6","Газ")]}="нет данных","ТРН 26/6. Нет майского показания газа. Коммуналка 1 253, итог и фин. учет без газа.","")'),
        (46, f'=IF(\'Расчет\'!H{addr[("ВЛКСМ 56","Газ")]}="","","ВЛКСМ 56, газ. "&\'Расчет\'!H{addr[("ВЛКСМ 56","Газ")]}&". Показание 438 осталось 438.")'),
        (47, f'=IF(\'Расчет\'!H{addr[("ВЛКСМ 56","Горячая вода")]}="","","ВЛКСМ 56, горячая вода. "&\'Расчет\'!H{addr[("ВЛКСМ 56","Горячая вода")]}&". Второй счётчик вырос с 2 до 10.")'),
        (48, f'=IF(\'Расчет\'!H{addr[("ВЛКСМ 56","Холодная вода")]}="","","ВЛКСМ 56, холодная вода. "&\'Расчет\'!H{addr[("ВЛКСМ 56","Холодная вода")]}&". Второй счётчик вырос с 57 до 65.")'),
        (49, '=IF(G33="да","Есть квартира, где свет за месяц больше 300 кВт·ч.","Свет. Ни у кого расход не больше 300 кВт·ч.")'),
        (50, '=IF(G34="да","Есть квартира, где горячая и холодная вода вместе больше 20 м³.","Вода. Ни у кого горячая и холодная вместе не больше 20 м³.")'),
        (51, '=IF(COUNTIF(\'Расчет\'!H4:H40,"*меньше*")>0,"Есть новое показание меньше предыдущего.","Новое показание меньше предыдущего нигде нет. Отрицательный расход из исходника убран.")'),
        (52, '="Выплата за май не внесена ни по одной квартире. Долг мая — «нет данных», а не вся сумма начисления."'),
        (53, '="ЛП 10. Майские счётчики воды уже два, апрельское показание одно. Расход посчитан по образцу исходника: сумма мая минус одно число апреля. По приборам это не разложить."'),
        (54, '="ТРН 26/6, апрель. Перечислено на 800 больше начисления. Переплата уменьшает выплату мая."'),
        (55, '="ВЛКСМ 56, апрель. Аренда по шапке 15 000, перечислено 14 000. Недоплата 1 000 прибавляется к маю."'),
    ]
    for row, formula in checks_text:
        write(ws, f"A{row}", formula, "warn")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        ws.row_dimensions[row].height = 32

    section(57, "Контроль: формулы должны сойтись с этими числами")
    write(ws, "A58", "Проверка", "header")
    write(ws, "B58", "Получилось", "header", align=CENTER)
    write(ws, "C58", "Должно быть", "header", align=CENTER)
    write(ws, "D58", "Сошлось", "header", align=CENTER)
    ws.merge_cells("D58:E58")

    control = []
    # Каждое начисление
    for apt in APTS:
        for res in charge_res:
            control.append((
                f"{apt['name']} — {res}, начислено",
                f"='Расчет'!F{addr[(apt['name'], res)]}",
                RESULT[apt["name"]]["charges"][res],
            ))
        control.append((f"{apt['name']} — коммуналка", f"='Итог'!{col_of[apt['name']]}7", RESULT[apt["name"]]["util"]))
        control.append((f"{apt['name']} — итого к выплате", f"='Итог'!{col_of[apt['name']]}11", RESULT[apt["name"]]["payout"]))
        control.append((f"{apt['name']} — фин. учет", f"='Итог'!{col_of[apt['name']]}12", RESULT[apt["name"]]["fin"]))
        control.append((f"{apt['name']} — долг апреля", f"='Итог'!{col_of[apt['name']]}23", RESULT[apt["name"]]["apr_debt"]))
    control.append(("Холодная вода, тариф", "=ROUND('Ввод'!B11,2)", 52.58))
    control.append(("Всего к выплате в мае", "='Итог'!G11", 70541))
    control.append(("Всего фин. учет мая", "='Итог'!G12", 72641))
    control.append(("Всего аренда мая", "='Итог'!G6", 66200))
    control.append(("Всего коммуналка мая", "='Итог'!G7", 6441))
    control.append(("Сумма долга апреля", "='Итог'!G23", 200))
    control.append(("ТРН, полнота", "='Итог'!D8", "неполная"))
    control.append(("ЛП 20, полнота", "='Итог'!C8", "полная"))
    control.append(("ВЛКСМ, расход газа", f"='Расчет'!C{addr[('ВЛКСМ 56','Газ')]}", 0))
    control.append(("ВЛКСМ, расход горячей", f"='Расчет'!C{addr[('ВЛКСМ 56','Горячая вода')]}", 8))
    control.append(("ВЛКСМ, расход холодной", f"='Расчет'!C{addr[('ВЛКСМ 56','Холодная вода')]}", 8))
    control.append(("Порог света, всего", "='Итог'!G33", "нет"))
    control.append(("Порог воды, всего", "='Итог'!G34", "нет"))
    control.append(("Перечислено мая, всего", "='Итог'!G13", "нет данных"))

    first = 59
    for n, (label, actual, expected) in enumerate(control):
        row = first + n
        write(ws, f"A{row}", label, "text")
        kind = "formula"
        fmt = None
        if isinstance(expected, float):
            fmt = TAR
        elif isinstance(expected, int):
            fmt = MONEY
        write(ws, f"B{row}", actual, kind, fmt, CENTER)
        exp_cell = write(ws, f"C{row}", expected, "input", fmt, CENTER)
        exp_cell.fill = FILLS["note"]
        write(ws, f"D{row}", f'=IF(AND(ISNUMBER(B{row}),ISNUMBER(C{row})),ABS(B{row}-C{row})<0.001,B{row}=C{row})', "formula", align=CENTER)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
    last = first + len(control) - 1
    total_row = last + 1
    write(ws, f"A{total_row}", "Все проверки", "text")
    ws[f"A{total_row}"].font = FONTS["bold"]
    write(ws, f"D{total_row}", f"=AND(D{first}:D{last})", "formula", align=CENTER)
    ws.merge_cells(start_row=total_row, start_column=4, end_row=total_row, end_column=5)
    write(ws, "A3", f'=IF(D{total_row},"Все контрольные проверки сошлись","Есть расхождение — смотри блок контроля внизу")', "formula")
    ws.merge_cells("A3:H3")
    ws.row_dimensions[3].height = 24

    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    green_font = Font(name="Calibri", color="006100", bold=True)
    red_font = Font(name="Calibri", color="9C0006", bold=True)
    ws.conditional_formatting.add(f"D{first}:D{total_row}", CellIsRule(operator="equal", formula=["TRUE"], fill=green, font=green_font))
    ws.conditional_formatting.add(f"D{first}:D{total_row}", CellIsRule(operator="equal", formula=["FALSE"], fill=red, font=red_font))
    ws.conditional_formatting.add("B8:F8", FormulaRule(formula=['B8="неполная"'], fill=FILLS["bad"]))
    ws.conditional_formatting.add("B13:F14", FormulaRule(formula=['B13="нет данных"'], fill=FILLS["warn"]))
    ws.conditional_formatting.add("B19:F19", FormulaRule(formula=['B19="нет данных"'], fill=FILLS["bad"]))
    ws.conditional_formatting.add("A3:H3", FormulaRule(formula=['ISNUMBER(SEARCH("сошлись",A3))'], fill=green, font=green_font))
    ws.conditional_formatting.add("A3:H3", FormulaRule(formula=['ISNUMBER(SEARCH("расхождение",A3))'], fill=red, font=red_font))
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    return total_row


def build_notes(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(0.6, 0.6, 0.6, 0.6)
    ws.column_dimensions["A"].width = 140
    paragraphs = [
        "Что сделано",
        "Это рабочая копия расчёта. Файл test1Data.xlsx не менялся, его снимок лежит на листе «Исходник». Майские показания взяты из задания. ЛП 10 в исходнике уже был заполнен как образец, я его сверил и не переписывал: газ 5 м³, свет 105 кВт·ч, горячая 5, холодная 8, коммуналка 1 535.",
        "Правила, по которым посчитаны цифры",
        "Расход = майское показание минус апрельское. Если майского показания нет, расход и начисление — «нет данных». Пустая ячейка в вычитании нулём не становится. У ВЛКСМ два счётчика воды: каждый майский прибор сравнивается со своим апрельским, потом расходы складываются. Начисление округляется как Excel ROUND до рубля. Горводснабж считается от расхода горячей воды, отдельного счётчика у него нет.",
        "Итого к выплате = аренда + коммуналка − доп. расходы + долг апреля. Долг апреля = начислено в апреле − перечислено в апреле. Плюс означает, что собственнику недоплатили, и майская выплата становится больше. Минус — переплата, майская выплата меньше. Фин. учет = аренда + коммуналка. Это расход компании. Доп. расход уже сидит в учёте расходов, поэтому второй раз в фин. учет не прибавляется и из него не вычитается. Перенос долга — это сальдо, а не новый расход месяца. Поэтому фин. учет и сумма к выплате совпадать не обязаны.",
        "Пустой доп. расход в исходнике я принял за 0: строка есть у всех квартир, сумма стоит только там, где расход был (май, ЛП 20, 2 300). Пустое показание так трактовать нельзя, и я его в ноль не превращал.",
        "Ошибки исходника и что с ними сделано",
        "1. Пустое майское показание в формуле «текущее − предыдущее» считалось нулём. Расход у ТРН, Ленины, ВЛКСМ и у света/воды ЛП 20 уходил в минус, коммуналка — в десятки тысяч минуса (ТРН −83 463, Ленина −95 251, ВЛКСМ −99 203). Теперь пустое показание даёт «нет данных».",
        "2. У ЛП 20 коммуналка мая была вбита числом 1 000, не формулой. От неё получились итог 11 700 и фин. учет 14 000. По тарифам коммуналка 960, к выплате 11 660, фин. учет 13 960. Разница 2 300 — это доп. расход.",
        "3. Формулы итога и фин. учета были разные у каждой квартиры. Где-то доп. расход вычитался из фин. учета, где-то из выплаты, где-то долг вычитался (знак переноса переворачивался), у ТРН фин. учет мая был пустой, у ЛП 10 фин. учет копировал итог к выплате. Сейчас правило одно на все квартиры.",
        "4. Знак переноса. В исходнике долг вычитали из следующего месяца. При недоплате выплата становилась меньше, при переплате — больше. Для расчёта с собственником это наоборот. Пример: по ТРН в апреле перечислили на 800 больше начисления, майская выплата уменьшена на 800.",
        "5. Аренда ВЛКСМ в апреле стояла 14 000, хотя в шапке с ноября по апрель указано 15+. Поставил 15 000. Перечисленные 14 000 не трогал: это факт выплаты. Долг апреля +1 000 перенесён в май. Если оставить апрельские 14 000, долг был бы 0, итог мая 15 856 и фин. учет 15 856. В расчёте используется вариант по шапке: итог мая 16 856, фин. учет 15 856.",
        "6. Долг мая в исходнике был равен всему итогу, потому что пустая выплата вычиталась как ноль. Оставил «нет данных».",
        "7. Коммуналка апреля была 0, потому что SUM складывал пустые начисления. Мартовских показаний нет, расход апреля неизвестен. Поставил «нет данных» и не включил её ни в итог апреля, ни в перенос. Перенос сделан только по арендному расчёту, который в исходнике реально оплачен.",
        "8. У ТРН сумма коммуналки мая не включала строку газа, у ЛП 10 в апреле «начислена коммуналка» была константой 0, а не SUM. В этой книге коммуналка — сумма всех ресурсов квартиры, а пропущенный ресурс виден отдельно как «неполная».",
        "Что не стал выдумывать",
        "Тариф холодной воды оставлен формулой 25,44 + 27,14 = 52,58. Оба месяца в исходнике совпадают, расшифровки второго слагаемого в файле нет. Аренда ТРН 13 200 взята из ячейки: в шапке суммы нет. Газ ТРН в апреле равен 35 — это сильно меньше соседних квартир, но это единственное показание, и я его не правил. У ЛП 10 апрельская вода не разложена на два счётчика, поэтому майские 10+2 и 15+3 нельзя вычесть по приборам; оставлен способ образца.",
        "Как читать отрицательный долг",
        "Минус в строке «долг / переплата» — это переплата собственнику. Плюс — недоплата. В мае долга нет, потому что перевода ещё не было, а не потому что расчёт сошёлся.",
        "Соответствие исходному листу",
        "Квартиры в исходнике шли блоками по три колонки: ЛП 10 — D/E/F, ЛП 20 — G/H/I, ТРН 26/6 — J/K/L, Ленина 108 — M/N/O, ВЛКСМ 56 — P/Q/R. Апрель — строки 3–14, май — строки 15–26. Тарифы — колонка C.",
    ]
    write(ws, "A1", "Замечания к расчёту", "title")
    ws.row_dimensions[1].height = 30
    row = 3
    for text in paragraphs:
        kind = "section" if text in {
            "Что сделано",
            "Правила, по которым посчитаны цифры",
            "Ошибки исходника и что с ними сделано",
            "Что не стал выдумывать",
            "Как читать отрицательный долг",
            "Соответствие исходному листу",
        } else "text"
        write(ws, f"A{row}", text, kind)
        ws.row_dimensions[row].height = 22 if kind == "section" else 48
        row += 1
    ws.oddHeader.left.text = "NikiRent — замечания"


def build_snapshot(ws):
    src = load_workbook(SRC)
    origin = src.active
    for col, dim in origin.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width or 12
    for row in origin.iter_rows():
        for cell in row:
            new = ws.cell(cell.row, cell.column, cell.value)
            new.number_format = cell.number_format
            if cell.font and cell.font.bold:
                new.font = Font(name="Calibri", size=11, bold=True)
    ws["V1"] = "Снимок test1Data.xlsx, лист «коммуналка». Сам файл не изменялся. Формулы здесь исходные, включая ошибочные: пустое показание они вычитают как ноль."
    ws["V1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["V"].width = 55
    ws.row_dimensions[1].height = 48
    src.close()


def _cons(curr, prev):
    return (
        f'IF(OR({curr}="нет счётчика",{prev}="нет счётчика"),"нет счётчика",'
        f'IF(OR({curr}="нет данных",{curr}="",{prev}="",NOT(ISNUMBER({curr})),NOT(ISNUMBER({prev}))),"нет данных",{curr}-{prev}))'
    )


def _charge(cons, tariff):
    return f'IF(NOT(ISNUMBER({cons})),{cons},ROUND({tariff}*{cons},0))'


def build_kommunalka(ws):
    """Копия исходной сетки: те же квартиры и строки, уже с майскими показаниями и исправленными формулами."""
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(0.4, 0.4, 0.5, 0.5)
    ws.freeze_panes = "D3"
    ws.oddHeader.left.text = "NikiRent — коммуналка, исправленная копия"
    widths = {
        "A": 14, "B": 22, "C": 12,
        "D": 28, "E": 14, "F": 16,
        "G": 28, "H": 14, "I": 16,
        "J": 28, "K": 14, "L": 16,
        "M": 28, "N": 14, "O": 16,
        "P": 28, "Q": 14, "R": 16,
        "T": 16, "U": 16, "V": 14, "W": 14, "X": 14, "Y": 14, "Z": 14, "AA": 14,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    write(ws, "A1", "Исправленная копия листа. Исходный test1Data.xlsx не менялся. Жёлтое — ввод, зелёное — формула.", "note")
    ws.merge_cells("A1:C1")
    ws.row_dimensions[1].height = 32
    titles = {
        "D": "1к ЛП 10, в мае 12 000 + к/у",
        "G": "1к ЛП 20, до 5-го числа, 13+",
        "J": "1 комн. студия ТРН 26/6. Аренда 13 200 взята из ячейки, в шапке суммы нет",
        "M": "2х Ленина 108, 14+, до 30-го числа",
        "P": "2х ВЛКСМ 56. Ноябрь–апрель 15+, май–октябрь 14+. В исходнике в апреле стояло 14 000",
    }
    for col, text in titles.items():
        write(ws, f"{col}1", text, "header")
        ws.merge_cells(start_row=1, start_column=ord(col) - 64, end_row=1, end_column=ord(col) - 62)
    write(ws, "C2", "тариф", "header", align=CENTER)
    for col in ("D", "G", "J", "M", "P"):
        base = ord(col) - 64
        write(ws, f"{col}2", "пок тек", "header", align=CENTER)
        write(ws, f"{get_column_letter(base + 1)}2", "куб, кВт", "header", align=CENTER)
        write(ws, f"{get_column_letter(base + 2)}2", "начислено", "header", align=CENTER)

    write(ws, "A3", "апрель", "section")
    resources = [(3, "Газ"), (4, "Свет"), (5, "ГОР вода"), (6, "ХОЛ вода"), (7, "горводснабж")]
    for row, name in resources:
        write(ws, f"B{row}", name, "text")
    write(ws, "C3", 9.5, "input", TAR, CENTER)
    write(ws, "C4", 3.83, "input", TAR, CENTER)
    write(ws, "C5", 25.44, "input", TAR, CENTER)
    write(ws, "C6", "=C5+27.14", "formula", TAR, CENTER)
    comment(ws["C6"], "Формула исходника: 25,44 + 27,14. Второе слагаемое в файле не расшифровано, тариф не менял.")
    write(ws, "C7", 107.4, "input", TAR, CENTER)

    # Одиночные апрельские показания. ВЛКСМ вода и май ЛП 10 — в расшифровке справа.
    april_single = {
        "D3": 10, "J3": 35, "M3": 500, "P3": 438,
        "D4": 2300, "G4": 3343, "J4": 7402, "M4": 12552, "P4": 7553,
        "D5": 7, "G5": 196, "J5": 196, "M5": 186,
        "D6": 10, "G6": 156, "J6": 553, "M6": 337,
    }
    for addr, val in april_single.items():
        write(ws, addr, val, "input", QTY, CENTER)
    for addr in ("G3",):
        write(ws, addr, "нет счётчика", "input", align=CENTER)
    for col in ("D", "G", "J", "M", "P"):
        write(ws, f"{col}7", None, "text", align=CENTER)

    write(ws, "T3", "Расшифровка двух счётчиков. В основной таблице показание — сумма, расход воды ВЛКСМ — сумма расходов приборов.", "note")
    ws.merge_cells("T3:AA3")
    for col, text in zip(["T", "U", "V", "W", "X", "Y", "Z", "AA"], ["Квартира", "Ресурс", "Апрель 1", "Апрель 2", "Май 1", "Май 2", "Расход 1", "Расход 2"]):
        write(ws, f"{col}4", text, "header", align=CENTER)
    meters = [
        (5, "ВЛКСМ 56", "ГОР вода", 316, 2, 316, 10),
        (6, "ВЛКСМ 56", "ХОЛ вода", 397, 57, 397, 65),
        (7, "ЛП 10", "ГОР вода, май", None, None, 10, 2),
        (8, "ЛП 10", "ХОЛ вода, май", None, None, 15, 3),
    ]
    for row, apt, res, a1, a2, m1, m2 in meters:
        write(ws, f"T{row}", apt, "text")
        write(ws, f"U{row}", res, "text")
        for col, val in (("V", a1), ("W", a2), ("X", m1), ("Y", m2)):
            if val is None:
                write(ws, f"{col}{row}", None, "input", QTY, CENTER)
            else:
                write(ws, f"{col}{row}", val, "input", QTY, CENTER)
        if a1 is None:
            write(ws, f"Z{row}", "апрель одним числом", "note", align=CENTER)
            write(ws, f"AA{row}", "см. основной расход", "note", align=CENTER)
        else:
            write(ws, f"Z{row}", f"={_cons(f'X{row}', f'V{row}')}", "formula", QTY, CENTER)
            write(ws, f"AA{row}", f"={_cons(f'Y{row}', f'W{row}')}", "formula", QTY, CENTER)
    write(ws, "T9", "ЛП 10: в апреле вода одним числом (D5 и D6), в мае уже два прибора. Расход как в образце исходника: сумма мая минус одно апрельское число. По приборам это не сопоставить.", "note")
    ws.merge_cells("T9:AA9")
    ws.row_dimensions[9].height = 32

    write(ws, "P5", "=V5+W5", "formula", QTY, CENTER)
    write(ws, "P6", "=V6+W6", "formula", QTY, CENTER)

    blocks = ("D", "G", "J", "M", "P")
    for col in blocks:
        nxt = get_column_letter(ord(col) - 64 + 1)
        money = get_column_letter(ord(col) - 64 + 2)
        for row, tariff in ((3, "$C$3"), (4, "$C$4"), (5, "$C$5"), (6, "$C$6")):
            write(ws, f"{nxt}{row}", "нет данных", "bad", align=CENTER)
            write(ws, f"{money}{row}", f"={_charge(f'{nxt}{row}', tariff)}", "formula", MONEY, CENTER)
        write(ws, f"{money}7", f"={_charge(f'{nxt}5', '$C$7')}", "formula", MONEY, CENTER)

    money_labels = [
        (8, "Начислена коммуналка"),
        (9, "Аренда"),
        (10, "Доп. расходы на квартиру"),
        (11, "Итого начислено"),
        (12, "перечислено"),
        (13, "долг"),
        (14, "фин. учет"),
    ]
    for row, label in money_labels:
        for col in blocks:
            write(ws, f"{col}{row}", label, "text")

    rents_apr = {"F": 12000, "I": 13000, "L": 13200, "O": 14000, "R": 15000}
    paid_apr = {"F": 12000, "I": 13000, "L": 14000, "O": 14000, "R": 14000}
    for col, val in rents_apr.items():
        write(ws, f"{col}9", val, "input", MONEY, CENTER)
        write(ws, f"{col}10", 0, "input", MONEY, CENTER)
        write(ws, f"{col}12", paid_apr[col], "input", MONEY, CENTER)
        write(ws, f"{col}8", f'=IF(COUNT({col}3:{col}7)=0,"нет данных",SUM({col}3:{col}7))', "formula", align=CENTER)
        write(ws, f"{col}11", f"={col}9-{col}10", "formula", MONEY, CENTER)
        write(ws, f"{col}13", f"={col}11-{col}12", "formula", MONEY, CENTER)
        write(ws, f"{col}14", f"={col}9", "formula", MONEY, CENTER)
    comment(ws["R9"], "В исходнике было 14 000. Шапка: с ноября по апрель 15+. Перечислено 14 000 не менял.")
    comment(ws["L12"], "Факт перевода. На 800 больше аренды, это переплата.")

    write(ws, "A15", "май", "section")
    for src, dst in ((3, 15), (4, 16), (5, 17), (6, 18), (7, 19)):
        write(ws, f"B{dst}", f"=B{src}", "formula")
        write(ws, f"C{dst}", f"=C{src}", "formula", TAR, CENTER)
    may_single = {
        "D15": 15, "J15": "нет данных", "M15": 503, "P15": 438,
        "D16": 2405, "G16": 3400, "J16": 7501, "M16": 12701, "P16": 7650,
        "G17": 200, "J17": 201, "M17": 187,
        "G18": 160, "J18": 557, "M18": 339,
    }
    for addr, val in may_single.items():
        kind = "input"
        fmt = None if isinstance(val, str) else QTY
        write(ws, addr, val, kind, fmt, CENTER)
    write(ws, "G15", "нет счётчика", "input", align=CENTER)
    for col in blocks:
        write(ws, f"{col}19", None, "text", align=CENTER)
    write(ws, "D17", "=X7+Y7", "formula", QTY, CENTER)
    write(ws, "D18", "=X8+Y8", "formula", QTY, CENTER)
    write(ws, "P17", "=X5+Y5", "formula", QTY, CENTER)
    write(ws, "P18", "=X6+Y6", "formula", QTY, CENTER)

    for col, prev_col in (("D", "D"), ("G", "G"), ("J", "J"), ("M", "M"), ("P", "P")):
        nxt = get_column_letter(ord(col) - 64 + 1)
        money = get_column_letter(ord(col) - 64 + 2)
        pairs = ((15, 3, "$C$15"), (16, 4, "$C$16"), (17, 5, "$C$17"), (18, 6, "$C$18"))
        for may_row, apr_row, tariff in pairs:
            if col == "P" and may_row in (17, 18):
                src = "Z5+AA5" if may_row == 17 else "Z6+AA6"
                write(ws, f"{nxt}{may_row}", f"=IF(OR(NOT(ISNUMBER(Z{5 if may_row == 17 else 6})),NOT(ISNUMBER(AA{5 if may_row == 17 else 6}))),\"нет данных\",{src})", "formula", QTY, CENTER)
            else:
                write(ws, f"{nxt}{may_row}", f"={_cons(f'{col}{may_row}', f'{prev_col}{apr_row}')}", "formula", QTY, CENTER)
            write(ws, f"{money}{may_row}", f"={_charge(f'{nxt}{may_row}', tariff)}", "formula", MONEY, CENTER)
        write(ws, f"{money}19", f"={_charge(f'{nxt}17', '$C$19')}", "formula", MONEY, CENTER)

    for row, label in (
        (20, "Начислена коммуналка"),
        (21, "Аренда"),
        (22, "Доп. расходы на квартиру"),
        (23, "Итого начислено"),
        (24, "перечислено"),
        (25, "долг"),
        (26, "фин. учет"),
        (27, "полнота коммуналки"),
    ):
        for col in blocks:
            write(ws, f"{col}{row}", label, "text")

    rents_may = {"F": 12000, "I": 13000, "L": 13200, "O": 14000, "R": 14000}
    for col, val in rents_may.items():
        write(ws, f"{col}21", val, "input", MONEY, CENTER)
        extra = 2300 if col == "I" else 0
        write(ws, f"{col}22", extra, "input", MONEY, CENTER)
        write(ws, f"{col}24", "нет данных", "input", align=CENTER)
        write(ws, f"{col}20", f'=IF(COUNT({col}15:{col}19)=0,"нет данных",SUM({col}15:{col}19))', "formula", MONEY, CENTER)
        write(ws, f"{col}23", f"={col}21+{col}20-{col}22+{col}13", "formula", MONEY, CENTER)
        write(ws, f"{col}25", f'=IF(OR({col}24="",{col}24="нет данных"),"нет данных",{col}23-{col}24)', "formula", MONEY, CENTER)
        write(ws, f"{col}26", f"={col}21+{col}20", "formula", MONEY, CENTER)
        write(ws, f"{col}27", f'=IF(COUNTIF({col}15:{col}19,"нет данных")>0,"неполная","полная")', "formula", align=CENTER)
    comment(ws["I22"], "2 300 вычитаются из выплаты собственнику. В фин. учет второй раз не входят.")
    comment(ws["J15"], "В апреле газ был 35. В задании на май показания нет. Ноль не ставил.")
    comment(ws["P15"], "Показание совпало с апрелем. Расход 0, это не пустая ячейка.")

    write(ws, "A29", "Как читать строки денег", "section")
    ws.merge_cells("A29:R29")
    rules = [
        "Итого начислено в мае = аренда + коммуналка − доп. расходы + долг апреля. Недоплата прибавляется, переплата вычитается.",
        "Фин. учет = аренда + коммуналка. Это расход компании. Доп. расход уже учтён отдельно, поэтому фин. учет не обязан совпадать с выплатой.",
        "В апреле коммуналка — «нет данных»: нет показаний марта. Итого апреля = аренда − доп. расходы. Ноль вместо неизвестной коммуналки не ставится.",
        "Пустая выплата мая — «нет данных». Долг мая от нуля не считается.",
        "ЛП 20: газа нет в обоих месяцах, это не пропуск показания. ТРН: газа в мае нет, коммуналка и итог неполные.",
        "ВЛКСМ, апрель: аренда 15 000 вместо 14 000 из исходной ячейки. Перечислено 14 000 оставлено, долг +1 000 перенесён в май.",
        "ТРН, апрель: перечислено на 800 больше начисления. Переплата уменьшает май.",
    ]
    for i, text in enumerate(rules):
        row = 30 + i
        write(ws, f"A{row}", text, "note")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=18)
        ws.row_dimensions[row].height = 22

    write(ws, "A38", "Сверка с листом «Итог». Должно быть ИСТИНА.", "section")
    ws.merge_cells("A38:C38")
    checks = [
        (39, "ЛП 10 коммуналка", "=F20='Итог'!B7"),
        (40, "ЛП 20 коммуналка", "=I20='Итог'!C7"),
        (41, "ТРН коммуналка", "=L20='Итог'!D7"),
        (42, "Ленина коммуналка", "=O20='Итог'!E7"),
        (43, "ВЛКСМ коммуналка", "=R20='Итог'!F7"),
        (44, "ЛП 10 к выплате", "=F23='Итог'!B11"),
        (45, "ЛП 20 к выплате", "=I23='Итог'!C11"),
        (46, "ТРН к выплате", "=L23='Итог'!D11"),
        (47, "Ленина к выплате", "=O23='Итог'!E11"),
        (48, "ВЛКСМ к выплате", "=R23='Итог'!F11"),
        (49, "ЛП 10 фин. учет", "=F26='Итог'!B12"),
        (50, "ЛП 20 фин. учет", "=I26='Итог'!C12"),
        (51, "ТРН фин. учет", "=L26='Итог'!D12"),
        (52, "Ленина фин. учет", "=O26='Итог'!E12"),
        (53, "ВЛКСМ фин. учет", "=R26='Итог'!F12"),
        (54, "ВЛКСМ долг апреля", "=R13='Итог'!F23"),
        (55, "ТРН долг апреля", "=L13='Итог'!D23"),
    ]
    for row, label, formula in checks:
        write(ws, f"A{row}", label, "text")
        write(ws, f"B{row}", formula, "formula", align=CENTER)
    write(ws, "A56", "Вся сверка", "text")
    ws["A56"].font = FONTS["bold"]
    write(ws, "B56", "=AND(B39:B55)", "formula", align=CENTER)
    green = PatternFill("solid", fgColor="C6EFCE")
    red = PatternFill("solid", fgColor="FFC7CE")
    ws.conditional_formatting.add("B39:B56", CellIsRule(operator="equal", formula=["TRUE"], fill=green, font=Font(name="Calibri", color="006100", bold=True)))
    ws.conditional_formatting.add("B39:B56", CellIsRule(operator="equal", formula=["FALSE"], fill=red, font=Font(name="Calibri", color="9C0006", bold=True)))


if __name__ == "__main__":
    build()
    print("saved", OUT)
    for name, data in RESULT.items():
        print(name, "util", data["util"], "payout", data["payout"], "fin", data["fin"], "apr_debt", data["apr_debt"], "incomplete", data["incomplete"])
        print(" ", data["cons"])
        print(" ", data["charges"])
