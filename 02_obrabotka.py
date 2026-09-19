# -*- coding: utf-8 -*-
"""
Скрипт обработки бухгалтерских операций с применением ИИ-подхода
(правило-ориентированная классификация + автоматическая проверка данных).

Что делает скрипт:
  1. Читает исходную таблицу операций (01_iskhodnye_dannye.xlsx).
  2. Автоматически классифицирует каждую операцию по упрощённому плану счетов.
  3. Находит и подсвечивает две заложенные в данные ошибки:
       - задвоенный платёж;
       - сумму, не соответствующую расчёту в назначении платежа.
  4. Формирует итоговый упрощённый баланс и отчёт о прибылях и убытках (ОПиУ).
  5. Генерирует текстовое пояснение к отчёту простым языком.
  6. Сохраняет результат в Excel (output/Itogovaya_otchetnost.xlsx)
     и текст пояснительной записки (output/poyasnitelnaya_zapiska.txt).

Демонстрационный пример для лекции «Составление финансовой отчётности
с применением ИИ-сервисов».
"""

import re
import json
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

SRC_PATH = BASE_DIR / "01_iskhodnye_dannye.xlsx"
OUT_XLSX = OUTPUT_DIR / "Itogovaya_otchetnost.xlsx"
OUT_TXT = OUTPUT_DIR / "poyasnitelnaya_zapiska.txt"
OUT_JSON = OUTPUT_DIR / "svodnye_pokazateli.json"

# Остаток денежных средств на начало месяца (условное допущение для примера)
OPENING_BALANCE = 2_500_000

COMPANY_NAME = "ТОО «Светлый Дом»"
PERIOD_TITLE = "август 2026 года"


# ---------------------------------------------------------------------------
# 1. Упрощённый план счетов и правила классификации
# ---------------------------------------------------------------------------
# Коды счетов — по Типовому плану счетов бухгалтерского учёта Республики Казахстан
# (утверждён Приказом МФ РК от 23.05.2007 №185): разделы 1 (краткосрочные активы),
# 3 (краткосрочные обязательства), 6 (доходы), 7 (расходы).
# Каждое правило: (регулярное выражение по назначению платежа, счёт, название статьи)
EXPENSE_RULES = [
    (r"аренд", "7210", "Административные расходы — аренда"),
    (r"электроэнерг|водоснабж|водоотвед", "7210", "Административные расходы — коммунальные услуги"),
    (r"зарплат", "3350", "Краткосрочная задолженность по оплате труда"),
    (r"социальн.*отчисл|страхов.*взнос", "3220", "Обязательства по социальному страхованию"),
    (r"снр|налог", "3110", "Расчёты по налогам (КПН/СНР)"),
    (r"кредит|процент", "7310", "Расходы по вознаграждениям — проценты по кредиту"),
    (r"материал", "1350", "Материалы"),
    (r"реклам", "7110", "Расходы по реализации — реклама"),
    (r"транспорт|достав", "7110", "Расходы по реализации — транспорт"),
    (r"охран", "7210", "Административные расходы — охрана"),
    (r"комисси.*банк|банк.*комисси|расчётно-кассов", "7210", "Административные расходы — банковские услуги"),
    (r"канцеляр", "7210", "Административные расходы — канцтовары"),
    (r"ремонт", "7210", "Административные расходы — ремонт"),
    (r"связ|интернет", "7210", "Административные расходы — связь"),
]
INCOME_ACCOUNT = ("6010", "Доход от реализации продукции и услуг")
DEFAULT_EXPENSE_ACCOUNT = ("7470", "Прочие расходы (не распознано)")


def classify(row):
    if row["Тип"] == "доход":
        return INCOME_ACCOUNT
    text = row["Назначение платежа"].lower()
    for pattern, code, name in EXPENSE_RULES:
        if re.search(pattern, text):
            return (code, name)
    return DEFAULT_EXPENSE_ACCOUNT


# ---------------------------------------------------------------------------
# 2. Поиск ошибок
# ---------------------------------------------------------------------------
def find_duplicates(df):
    """Возвращает (множество индексов строк-дублей, количество дублирующихся групп)."""
    key_cols = ["Дата", "Контрагент", "Назначение платежа", "Сумма, тенге", "Тип"]
    dup_mask = df.duplicated(subset=key_cols, keep=False)
    dup_idx = set(df.index[dup_mask])
    n_groups = df.loc[dup_mask, key_cols].drop_duplicates().shape[0]
    return dup_idx, n_groups


SALARY_PATTERN = re.compile(r"(\d+)\s*сотрудник\w*\s*по\s*([\d\s]+)\s*тенге")


def find_amount_mismatches(df):
    """
    Ищет операции, где в самом назначении платежа указан расчёт
    (например, «5 сотрудников по 45 000 тенге»), и сверяет его
    с фактически указанной суммой.
    Возвращает словарь {индекс_строки: (ожидаемая_сумма, фактическая_сумма)}.
    """
    mismatches = {}
    for idx, row in df.iterrows():
        m = SALARY_PATTERN.search(row["Назначение платежа"])
        if m:
            count = int(m.group(1))
            rate = int(m.group(2).replace(" ", ""))
            expected = count * rate
            actual = row["Сумма, тенге"]
            if expected != actual:
                mismatches[idx] = (expected, actual)
    return mismatches


# ---------------------------------------------------------------------------
# 3. Основная обработка
# ---------------------------------------------------------------------------
def main():
    df = pd.read_excel(SRC_PATH, sheet_name="Операции")

    df[["Счёт", "Статья"]] = df.apply(lambda r: pd.Series(classify(r)), axis=1)

    duplicate_idx, n_duplicate_groups = find_duplicates(df)
    mismatch_idx = find_amount_mismatches(df)

    # Данные по найденным ошибкам для отчёта/пояснительной записки (единый источник правды)
    duplicate_example = None
    if duplicate_idx:
        first = df.loc[min(duplicate_idx)]
        same_key = (
            (df["Дата"] == first["Дата"])
            & (df["Контрагент"] == first["Контрагент"])
            & (df["Назначение платежа"] == first["Назначение платежа"])
            & (df["Сумма, тенге"] == first["Сумма, тенге"])
        )
        duplicate_example = {
            "purpose": first["Назначение платежа"],
            "amount": int(first["Сумма, тенге"]),
            "occurrences": int(same_key.sum()),
        }
    mismatch_example = None
    if mismatch_idx:
        idx0 = next(iter(mismatch_idx))
        expected0, actual0 = mismatch_idx[idx0]
        mismatch_example = {
            "purpose": df.loc[idx0, "Назначение платежа"],
            "expected": int(expected0),
            "actual": int(actual0),
            "diff": int(expected0 - actual0),
        }

    # Пометка ошибок и комментарий по каждой строке
    comments = {}
    for idx in duplicate_idx:
        comments.setdefault(idx, []).append(
            "Задвоенный платёж: операция полностью повторяет другую строку в таблице "
            "(одинаковые дата, контрагент, назначение и сумма). Требуется исключить один из платежей."
        )
    for idx, (expected, actual) in mismatch_idx.items():
        diff = expected - actual
        comments.setdefault(idx, []).append(
            f"Сумма не соответствует расчёту в назначении платежа: "
            f"по тексту ожидается {fmt(expected)} ₸, а указано {fmt(actual)} ₸ "
            f"(расхождение {fmt(diff)} ₸). Требуется уточнить сумму у бухгалтера."
        )

    df["Ошибка"] = df.index.map(lambda i: "; ".join(comments.get(i, [])) or "")
    error_rows = set(comments.keys())

    # --- Формируем "исправленные" данные для отчётности ---
    # 1) убираем один из задвоенных платежей (оставляем первое вхождение)
    df_corrected = df.drop_duplicates(
        subset=["Дата", "Контрагент", "Назначение платежа", "Сумма, тенге", "Тип"], keep="first"
    ).copy()
    # 2) исправляем сумму там, где она не совпадает с расчётом в тексте
    for idx, (expected, actual) in mismatch_idx.items():
        if idx in df_corrected.index:
            df_corrected.loc[idx, "Сумма, тенге"] = expected

    income_mask = df_corrected["Тип"] == "доход"
    expense_mask = df_corrected["Тип"] == "расход"

    total_income = df_corrected.loc[income_mask, "Сумма, тенге"].sum()
    total_expense = df_corrected.loc[expense_mask, "Сумма, тенге"].sum()
    net_profit = total_income - total_expense
    closing_balance = OPENING_BALANCE + total_income - total_expense

    # Группировка расходов по статьям для ОПиУ
    expense_by_item = (
        df_corrected.loc[expense_mask]
        .groupby("Статья", as_index=False)["Сумма, тенге"]
        .sum()
        .sort_values("Сумма, тенге", ascending=False)
    )
    income_by_item = (
        df_corrected.loc[income_mask]
        .groupby("Статья", as_index=False)["Сумма, тенге"]
        .sum()
        .sort_values("Сумма, тенге", ascending=False)
    )

    # Сумма "наивного" (необработанного) результата — для наглядности эффекта ошибок
    naive_income = df.loc[df["Тип"] == "доход", "Сумма, тенге"].sum()
    naive_expense = df.loc[df["Тип"] == "расход", "Сумма, тенге"].sum()
    naive_profit = naive_income - naive_expense

    write_excel(df, error_rows, df_corrected, income_by_item, expense_by_item,
                total_income, total_expense, net_profit, closing_balance,
                naive_profit, n_duplicate_groups, len(mismatch_idx))

    explanation = build_explanation(
        total_income, total_expense, net_profit, naive_profit,
        n_duplicate_groups, len(mismatch_idx), closing_balance
    )
    with open(OUT_TXT, "w", encoding="utf-8") as f:
        f.write(explanation)

    summary = {
        "company": COMPANY_NAME,
        "period": PERIOD_TITLE,
        "total_income": int(total_income),
        "total_expense": int(total_expense),
        "net_profit": int(net_profit),
        "opening_balance": int(OPENING_BALANCE),
        "closing_balance": int(closing_balance),
        "naive_profit": int(naive_profit),
        "n_duplicate_groups": n_duplicate_groups,
        "n_mismatch": len(mismatch_idx),
        "duplicate_error": duplicate_example,
        "mismatch_error": mismatch_example,
        "income_by_item": [
            {"item": r["Статья"], "amount": int(r["Сумма, тенге"])}
            for _, r in income_by_item.iterrows()
        ],
        "expense_by_item": [
            {"item": r["Статья"], "amount": int(r["Сумма, тенге"])}
            for _, r in expense_by_item.iterrows()
        ],
        "explanation_paragraphs": explanation.split("\n\n")[-3:],
    }
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print("Готово.")
    print(f" - Итоговый отчёт: {OUT_XLSX}")
    print(f" - Пояснительная записка (текст): {OUT_TXT}")
    print(f" - Найдено ошибок: дубликатов — {len(duplicate_idx)}, "
          f"расхождений в сумме — {len(mismatch_idx)}")
    print(f" - Выручка: {fmt(total_income)} ₸ | Расходы: {fmt(total_expense)} ₸ | "
          f"Чистая прибыль: {fmt(net_profit)} ₸")


# ---------------------------------------------------------------------------
# 4. Текстовое пояснение простым языком
# ---------------------------------------------------------------------------
def fmt(n):
    """Форматирует сумму с пробелом в качестве разделителя разрядов: 620900 -> '620 900'."""
    return f"{n:,.0f}".replace(",", " ")


def build_explanation(total_income, total_expense, net_profit, naive_profit,
                       n_dup_groups, n_mismatch, closing_balance):
    profit_word = "прибыль" if net_profit >= 0 else "убыток"
    company_short = COMPANY_NAME.replace("ТОО", "").strip(" «»")
    p1 = (
        f"За {PERIOD_TITLE} компания «{company_short}» получила выручку "
        f"в размере {fmt(total_income)} ₸ — в основном за счёт оптовых продаж товаров и розничной "
        f"торговли, а также разовых консультационных услуг. Расходы компании за тот же период "
        f"составили {fmt(total_expense)} ₸ и включают оплату труда сотрудников, аренду офиса, "
        f"закупку материалов, коммунальные и налоговые платежи, а также услуги сторонних организаций "
        f"(реклама, охрана, банк, связь)."
    )

    p2 = (
        f"По итогам месяца финансовый результат компании — {profit_word} в размере "
        f"{fmt(abs(net_profit))} ₸. Остаток денежных средств на конец месяца составил "
        f"{fmt(closing_balance)} ₸ (при остатке {fmt(OPENING_BALANCE)} ₸ на начало месяца). "
        f"Эта сумма отражена в упрощённом балансе как денежные средства в активе и как "
        f"собственный капитал (с учётом прибыли месяца) в пассиве — баланс сходится."
    )

    p3 = (
        f"При автоматической проверке данных ИИ-сервис обнаружил один задвоенный платёж "
        f"(одна и та же операция ошибочно повторяется в таблице дважды) "
        f"и {n_mismatch} операцию, где указанная сумма не совпадала с расчётом, приведённым "
        f"в самом назначении платежа. Если бы эти ошибки не были замечены и исправлены, "
        f"итоговый финансовый результат месяца показал бы {'прибыль' if naive_profit >= 0 else 'убыток'} "
        f"всего {fmt(abs(naive_profit))} ₸ — то есть отчётность вводила бы в заблуждение "
        f"примерно на {fmt(abs(net_profit - naive_profit))} ₸. Это наглядно показывает, зачем нужна "
        f"проверка первичных данных перед составлением отчётности, и как в этом может помочь ИИ."
    )

    return "\n\n".join([
        f"ПОЯСНИТЕЛЬНАЯ ЗАПИСКА К ФИНАНСОВОЙ ОТЧЁТНОСТИ",
        f"{COMPANY_NAME}, {PERIOD_TITLE}",
        "",
        p1, p2, p3,
    ])


# ---------------------------------------------------------------------------
# 5. Запись итогового Excel-файла
# ---------------------------------------------------------------------------
def write_excel(df, error_rows, df_corrected, income_by_item, expense_by_item,
                 total_income, total_expense, net_profit, closing_balance,
                 naive_profit, n_dup, n_mismatch):
    wb = openpyxl.Workbook()

    HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    ERROR_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    ERROR_FONT = Font(name="Arial", size=9, color="9C0006")
    BODY_FONT = Font(name="Arial", size=10)
    TITLE_FONT = Font(name="Arial", bold=True, size=14)
    SUBTITLE_FONT = Font(name="Arial", italic=True, size=10, color="595959")
    THIN = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def style_header(ws, row, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER

    # ---------------- Лист 1: Операции (проверено) ----------------
    ws1 = wb.active
    ws1.title = "Операции (проверено)"
    ws1.sheet_view.showGridLines = False
    cols = ["Дата", "Контрагент", "Назначение платежа", "Сумма, тенге", "Тип", "Счёт", "Статья", "Комментарий ИИ-проверки"]
    ws1.append(cols)
    style_header(ws1, 1, len(cols))
    ws1.row_dimensions[1].height = 28
    ws1.freeze_panes = "A2"

    for i, (idx, row) in enumerate(df.iterrows(), start=2):
        values = [row["Дата"], row["Контрагент"], row["Назначение платежа"], row["Сумма, тенге"],
                  row["Тип"], row["Счёт"], row["Статья"], row["Ошибка"]]
        ws1.append(values)
        is_error = idx in error_rows
        for c in range(1, len(cols) + 1):
            cell = ws1.cell(row=i, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=(c == 8))
            if is_error:
                cell.fill = ERROR_FILL
                cell.font = ERROR_FONT if c == 8 else Font(name="Arial", size=10, color="9C0006", bold=True)
            else:
                cell.font = BODY_FONT
        ws1.cell(row=i, column=4).number_format = '#,##0" ₸"'

    widths1 = [12, 30, 48, 13, 9, 7, 30, 45]
    for i, w in enumerate(widths1, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    last_row = len(df) + 1
    legend_row = last_row + 2
    ws1.cell(row=legend_row, column=1, value="Строки, выделенные красным, — операции с ошибками, найденными ИИ-проверкой.")
    ws1.cell(row=legend_row, column=1).font = SUBTITLE_FONT

    # ---------------- Лист 2: ОПиУ ----------------
    ws2 = wb.create_sheet("Отчёт о прибылях и убытках")
    ws2.sheet_view.showGridLines = False
    ws2.cell(row=1, column=1, value=f"Отчёт о прибылях и убытках — {COMPANY_NAME}").font = TITLE_FONT
    ws2.cell(row=2, column=1, value=f"за {PERIOD_TITLE} (после проверки и исправления ошибок)").font = SUBTITLE_FONT

    r = 4
    ws2.cell(row=r, column=1, value="ДОХОДЫ").font = Font(name="Arial", bold=True, size=11)
    r += 1
    income_start = r
    for _, item in income_by_item.iterrows():
        ws2.cell(row=r, column=1, value=item["Статья"]).font = BODY_FONT
        c = ws2.cell(row=r, column=2, value=item["Сумма, тенге"])
        c.number_format = '#,##0" ₸"'
        c.font = BODY_FONT
        r += 1
    income_end = r - 1
    ws2.cell(row=r, column=1, value="Итого доходы").font = Font(name="Arial", bold=True, size=10)
    total_income_cell = f"B{r}"
    c = ws2.cell(row=r, column=2, value=f"=SUM(B{income_start}:B{income_end})")
    c.number_format = '#,##0" ₸"'
    c.font = Font(name="Arial", bold=True, size=10)
    r += 2

    ws2.cell(row=r, column=1, value="РАСХОДЫ").font = Font(name="Arial", bold=True, size=11)
    r += 1
    expense_start = r
    for _, item in expense_by_item.iterrows():
        ws2.cell(row=r, column=1, value=item["Статья"]).font = BODY_FONT
        c = ws2.cell(row=r, column=2, value=item["Сумма, тенге"])
        c.number_format = '#,##0" ₸"'
        c.font = BODY_FONT
        r += 1
    expense_end = r - 1
    ws2.cell(row=r, column=1, value="Итого расходы").font = Font(name="Arial", bold=True, size=10)
    total_expense_cell = f"B{r}"
    c = ws2.cell(row=r, column=2, value=f"=SUM(B{expense_start}:B{expense_end})")
    c.number_format = '#,##0" ₸"'
    c.font = Font(name="Arial", bold=True, size=10)
    r += 2

    ws2.cell(row=r, column=1, value="ЧИСТАЯ ПРИБЫЛЬ").font = Font(name="Arial", bold=True, size=12)
    c = ws2.cell(row=r, column=2, value=f"={total_income_cell}-{total_expense_cell}")
    c.number_format = '#,##0" ₸"'
    c.font = Font(name="Arial", bold=True, size=12, color="1E7B34" if net_profit >= 0 else "B00020")
    profit_cell = f"B{r}"

    r += 3
    ws2.cell(row=r, column=1,
             value=(f"Справочно: без учёта проверки ИИ (с задвоенным платежом и ошибкой в сумме "
                    f"зарплаты) финансовый результат составил бы {fmt(naive_profit)} ₸.")).font = SUBTITLE_FONT

    ws2.column_dimensions["A"].width = 45
    ws2.column_dimensions["B"].width = 16

    # ---------------- Лист 3: Баланс ----------------
    ws3 = wb.create_sheet("Баланс")
    ws3.sheet_view.showGridLines = False
    ws3.cell(row=1, column=1, value=f"Упрощённый баланс — {COMPANY_NAME}").font = TITLE_FONT
    ws3.cell(row=2, column=1, value=f"на конец периода ({PERIOD_TITLE})").font = SUBTITLE_FONT

    r = 4
    ws3.cell(row=r, column=1, value="АКТИВЫ").font = Font(name="Arial", bold=True, size=11)
    r += 1
    ws3.cell(row=r, column=1, value="Денежные средства (расчётный счёт)").font = BODY_FONT
    cash_row = r
    c = ws3.cell(row=r, column=2, value=OPENING_BALANCE)
    c.number_format = '#,##0" ₸"'
    c.font = BODY_FONT
    r += 1
    ws3.cell(row=r, column=1, value="  + Доходы за месяц").font = BODY_FONT
    c = ws3.cell(row=r, column=2, value=f"='Отчёт о прибылях и убытках'!{total_income_cell}")
    c.number_format = '#,##0" ₸"'
    c.font = BODY_FONT
    income_ref_row = r
    r += 1
    ws3.cell(row=r, column=1, value="  − Расходы за месяц").font = BODY_FONT
    c = ws3.cell(row=r, column=2, value=f"=-'Отчёт о прибылях и убытках'!{total_expense_cell}")
    c.number_format = '#,##0" ₸"'
    c.font = BODY_FONT
    expense_ref_row = r
    r += 1
    ws3.cell(row=r, column=1, value="Итого активы").font = Font(name="Arial", bold=True, size=10)
    c = ws3.cell(row=r, column=2, value=f"=B{cash_row}+B{income_ref_row}+B{expense_ref_row}")
    c.number_format = '#,##0" ₸"'
    c.font = Font(name="Arial", bold=True, size=10)
    assets_total_row = r
    r += 2

    ws3.cell(row=r, column=1, value="ПАССИВЫ").font = Font(name="Arial", bold=True, size=11)
    r += 1
    ws3.cell(row=r, column=1, value="Собственный капитал на начало месяца").font = BODY_FONT
    c = ws3.cell(row=r, column=2, value=OPENING_BALANCE)
    c.number_format = '#,##0" ₸"'
    c.font = BODY_FONT
    equity_row = r
    r += 1
    ws3.cell(row=r, column=1, value="  + Чистая прибыль за месяц").font = BODY_FONT
    c = ws3.cell(row=r, column=2, value=f"='Отчёт о прибылях и убытках'!{profit_cell}")
    c.number_format = '#,##0" ₸"'
    c.font = BODY_FONT
    profit_ref_row = r
    r += 1
    ws3.cell(row=r, column=1, value="Обязательства").font = BODY_FONT
    c = ws3.cell(row=r, column=2, value=0)
    c.number_format = '#,##0" ₸"'
    c.font = BODY_FONT
    liab_row = r
    r += 1
    ws3.cell(row=r, column=1, value="Итого пассивы").font = Font(name="Arial", bold=True, size=10)
    c = ws3.cell(row=r, column=2, value=f"=B{equity_row}+B{profit_ref_row}+B{liab_row}")
    c.number_format = '#,##0" ₸"'
    c.font = Font(name="Arial", bold=True, size=10)
    liabilities_total_row = r
    r += 2

    ws3.cell(row=r, column=1, value="Проверка: Активы = Пассивы?").font = Font(name="Arial", italic=True, size=10)
    c = ws3.cell(row=r, column=2, value=f'=IF(B{assets_total_row}=B{liabilities_total_row},"Сходится","Не сходится!")')
    c.font = Font(name="Arial", italic=True, size=10, bold=True)

    ws3.column_dimensions["A"].width = 42
    ws3.column_dimensions["B"].width = 16

    # ---------------- Лист 4: Пояснительная записка ----------------
    ws4 = wb.create_sheet("Пояснительная записка")
    ws4.sheet_view.showGridLines = False
    ws4.column_dimensions["A"].width = 100
    explanation_text = build_explanation(total_income, total_expense, net_profit,
                                          naive_profit, n_dup, n_mismatch, closing_balance)
    row_n = 1
    for line in explanation_text.split("\n"):
        cell = ws4.cell(row=row_n, column=1, value=line)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(name="Arial", bold=(row_n in (1, 2)), size=13 if row_n == 1 else 10)
        ws4.row_dimensions[row_n].height = 18 if line else 8
        row_n += 1

    wb.save(OUT_XLSX)


if __name__ == "__main__":
    main()
