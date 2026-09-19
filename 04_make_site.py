# -*- coding: utf-8 -*-
"""
Собирает index.html из index.template.html, output/svodnye_pokazateli.json
и итогового Excel-отчёта. Единый источник цифр — output/svodnye_pokazateli.json,
поэтому суммы на сайте никогда не расходятся с отчётом и пояснительной запиской.

Запуск: python 04_make_site.py (после 02_obrabotka.py)
"""

import html
import json
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE = BASE_DIR / "index.template.html"
OUT_HTML = BASE_DIR / "index.html"
SUMMARY_JSON = BASE_DIR / "output" / "svodnye_pokazateli.json"
FINAL_XLSX = BASE_DIR / "output" / "Itogovaya_otchetnost.xlsx"


def fmt(n):
    return f"{n:,.0f}".replace(",", " ")


def build_rows():
    wb = openpyxl.load_workbook(FINAL_XLSX)
    ws = wb["Операции (проверено)"]

    data_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        date, contragent, purpose, amount, tip, schet, statya, error = row
        if amount is None or tip not in ("доход", "расход"):
            break  # дошли до пустой строки / легенды под таблицей
        data_rows.append(row)
    n_ops = len(data_rows)

    out = []
    for date, contragent, purpose, amount, tip, schet, statya, error in data_rows:
        error = error or ""
        cls = ' class="err-row"' if error else ""
        tip_cls = "income" if tip == "доход" else "expense"
        err_html = f'<div class="err-note">⚠ {html.escape(error)}</div>' if error else ""
        out.append(
            f'<tr{cls}><td>{date}</td><td>{html.escape(contragent)}</td>'
            f'<td>{html.escape(purpose)}{err_html}</td>'
            f'<td class="num">{fmt(amount)}</td>'
            f'<td class="{tip_cls}">{tip}</td><td>{schet}</td><td>{html.escape(statya)}</td></tr>'
        )
    return "\n".join(out), n_ops


def main():
    with open(SUMMARY_JSON, encoding="utf-8") as f:
        d = json.load(f)

    rows_html, n_ops = build_rows()

    tpl = TEMPLATE.read_text(encoding="utf-8")
    replacements = {
        "{{COMPANY}}": d["company"],
        "{{PERIOD}}": d["period"],
        "{{N_OPERATIONS}}": str(n_ops),
        "{{TOTAL_INCOME}}": fmt(d["total_income"]),
        "{{TOTAL_EXPENSE}}": fmt(d["total_expense"]),
        "{{NET_PROFIT}}": fmt(d["net_profit"]),
        "{{CLOSING_BALANCE}}": fmt(d["closing_balance"]),
        "{{NAIVE_PROFIT}}": fmt(d["naive_profit"]),
        "{{PROFIT_DIFF}}": fmt(d["net_profit"] - d["naive_profit"]),
        "{{DUP_PURPOSE}}": html.escape(d["duplicate_error"]["purpose"]),
        "{{DUP_AMOUNT}}": fmt(d["duplicate_error"]["amount"]),
        "{{DUP_OCCURRENCES}}": str(d["duplicate_error"]["occurrences"]),
        "{{MISMATCH_PURPOSE}}": html.escape(d["mismatch_error"]["purpose"]),
        "{{MISMATCH_ACTUAL}}": fmt(d["mismatch_error"]["actual"]),
        "{{MISMATCH_EXPECTED}}": fmt(d["mismatch_error"]["expected"]),
        "{{MISMATCH_DIFF}}": fmt(d["mismatch_error"]["diff"]),
        "{{EXPLAIN_P1}}": html.escape(d["explanation_paragraphs"][0], quote=False),
        "{{EXPLAIN_P2}}": html.escape(d["explanation_paragraphs"][1], quote=False),
        "{{EXPLAIN_P3}}": html.escape(d["explanation_paragraphs"][2], quote=False),
    }
    for token, value in replacements.items():
        tpl = tpl.replace(token, value)
    tpl = tpl.replace("<!--ROWS-->", rows_html)

    OUT_HTML.write_text(tpl, encoding="utf-8")
    print(f"Сохранено: {OUT_HTML}")


if __name__ == "__main__":
    main()
