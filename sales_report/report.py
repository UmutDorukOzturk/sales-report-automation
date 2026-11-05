"""The quarterly Excel report: dashboard, branch and product tables, daily trend, clean data and a data-quality log."""
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from . import config
from .load import FileLog

NAVY, BLUE, LIGHT, GREY, GREEN, RED = "1F3864", "2E75B6", "EEF3FA", "6B7280", "16794A", "B42318"
MONTH_COLORS = ["BDD7EE", "6FA8DC", "1F4E79"]
EUR, EUR0 = "€#,##0.00", "€#,##0"
PCT = "+0.0%;-0.0%;0.0%"
DATE = "[$-409]d mmm yyyy"
TABLE_STYLE = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)


def period_label(data: pd.DataFrame) -> str:
    first, last = data["date"].min(), data["date"].max()
    if first.year == last.year and (first.month - 1) // 3 == (last.month - 1) // 3 and first.month % 3 == 1 \
            and last.month % 3 == 0:
        return f"Q{(first.month - 1) // 3 + 1} {first.year}"
    return f"{first:%b} – {last:%b %Y}"


def write_report(data: pd.DataFrame, logs: list[FileLog], fixes: list, outliers: pd.DataFrame, path: Path,
                 run_at: datetime, source: str) -> None:
    data = data.copy()
    data["month"] = data["date"].map(lambda d: d.replace(day=1))
    months = sorted(data["month"].unique())
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    branches_ws = wb.create_sheet("By branch")
    products_ws = wb.create_sheet("Products")
    daily_ws = wb.create_sheet("Daily")
    data_ws = wb.create_sheet("Clean data")
    quality_ws = wb.create_sheet("Data quality")
    notes_ws = wb.create_sheet("Notes")

    n_branches = _branches_sheet(branches_ws, data, months)
    n_products = _products_sheet(products_ws, data, months)
    n_days = _daily_sheet(daily_ws, data)
    _data_sheet(data_ws, data)
    _quality_sheet(quality_ws, logs, fixes, outliers)
    _summary_sheet(summary, data, logs, fixes, outliers, months, run_at, source,
                   branches_ws, n_branches, products_ws, n_products, daily_ws, n_days)
    _notes_sheet(notes_ws, run_at, source, len(logs))
    wb.save(path)


# ---------------------------------------------------------------- helpers

def _headers(ws, row: int, headers: list[str], widths: list[float], start_col: int = 1) -> None:
    for i, (text, width) in enumerate(zip(headers, widths)):
        ws.cell(row, start_col + i, text)
        ws.column_dimensions[get_column_letter(start_col + i)].width = width


def _table(ws, name: str, first_row: int, last_row: int, first_col: int, last_col: int) -> None:
    ref = f"{get_column_letter(first_col)}{first_row}:{get_column_letter(last_col)}{max(last_row, first_row + 1)}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TABLE_STYLE
    ws.add_table(table)


def _total_row(ws, row: int, values: list, formats: list[Optional[str]]) -> None:
    line = Side(style="medium", color=NAVY)
    for c, (value, fmt) in enumerate(zip(values, formats), start=1):
        cell = ws.cell(row, c, value)
        cell.font = Font(bold=True, color=NAVY)
        cell.border = Border(top=line)
        if fmt:
            cell.number_format = fmt


def _section(ws, cell: str, text: str) -> None:
    ws[cell] = text
    ws[cell].font = Font(size=13, bold=True, color=NAVY)


# ---------------------------------------------------------------- sheets

def _branches_sheet(ws, data: pd.DataFrame, months: list) -> int:
    pivot = data.pivot_table(index="branch", columns="month", values="revenue", aggfunc="sum").fillna(0)
    pivot["total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("total", ascending=False)
    items = data.groupby("branch")["qty"].sum()
    headers = ["Branch"] + [f"{m:%b %Y}" for m in months] + ["Total", f"{months[-1]:%b} vs {months[0]:%b}",
                                                             "Share", "Items sold"]
    _headers(ws, 1, headers, [22] + [13] * len(months) + [14, 13, 9, 12])
    for r, (branch, row) in enumerate(pivot.iterrows(), start=2):
        ws.cell(r, 1, branch)
        for c, m in enumerate(months, start=2):
            ws.cell(r, c, round(row[m], 2)).number_format = EUR0
        col = len(months) + 2
        ws.cell(r, col, round(row["total"], 2)).number_format = EUR0
        growth = row[months[-1]] / row[months[0]] - 1 if row[months[0]] else None
        cell = ws.cell(r, col + 1, growth)
        cell.number_format = PCT
        cell.font = Font(bold=True, color=GREEN if (growth or 0) >= 0 else RED)
        ws.cell(r, col + 2, row["total"] / pivot["total"].sum()).number_format = "0%"
        ws.cell(r, col + 3, int(items[branch])).number_format = "#,##0"
    last = len(pivot) + 1
    _table(ws, "Branches", 1, last, 1, len(headers))
    totals = [pivot[m].sum() for m in months]
    _total_row(ws, last + 1, ["All branches"] + totals + [sum(totals), totals[-1] / totals[0] - 1, 1,
                                                          int(items.sum())],
               [None] + [EUR0] * (len(months) + 1) + [PCT, "0%", "#,##0"])
    ws.freeze_panes = "B2"
    return len(pivot)


def _products_sheet(ws, data: pd.DataFrame, months: list) -> int:
    g = data.groupby("product").agg(units=("qty", "sum"), revenue=("revenue", "sum")).sort_values("revenue",
                                                                                                   ascending=False)
    by_month = data.pivot_table(index="product", columns="month", values="revenue", aggfunc="sum").fillna(0)
    headers = ["Product", "Units sold", "Revenue", "Avg price", "Share"] + [f"{m:%b}" for m in months]
    _headers(ws, 1, headers, [18, 12, 13, 11, 9] + [12] * len(months))
    for r, (product, row) in enumerate(g.iterrows(), start=2):
        values = [product, int(row["units"]), round(row["revenue"], 2), row["revenue"] / row["units"],
                  row["revenue"] / g["revenue"].sum()] + [round(by_month.loc[product, m], 2) for m in months]
        formats = [None, "#,##0", EUR0, EUR, "0.0%"] + [EUR0] * len(months)
        for c, (value, fmt) in enumerate(zip(values, formats), start=1):
            cell = ws.cell(r, c, value)
            if fmt:
                cell.number_format = fmt
    _table(ws, "Products", 1, len(g) + 1, 1, len(headers))
    return len(g)


def _daily_sheet(ws, data: pd.DataFrame) -> int:
    daily = data.pivot_table(index="date", columns="branch", values="revenue", aggfunc="sum").fillna(0)
    daily.insert(0, "All branches", daily.sum(axis=1))
    headers = ["Date", "Weekday"] + list(daily.columns)
    _headers(ws, 1, headers, [13, 11] + [15] * len(daily.columns))
    for r, (d, row) in enumerate(daily.iterrows(), start=2):
        ws.cell(r, 1, d).number_format = "[$-409]d mmm"
        ws.cell(r, 2, d.strftime("%A"))
        for c, value in enumerate(row, start=3):
            ws.cell(r, c, round(value, 2)).number_format = EUR0
    _table(ws, "Daily", 1, len(daily) + 1, 1, len(headers))
    ws.freeze_panes = "C2"
    return len(daily)


def _data_sheet(ws, data: pd.DataFrame) -> None:
    headers = ["Date", "Month", "Branch", "Product", "Qty", "Unit price", "Revenue", "Original name", "Source file"]
    _headers(ws, 1, headers, [13, 10, 20, 15, 8, 11, 11, 22, 34])
    ordered = data.sort_values(["date", "branch", "product"])
    for r, row in enumerate(ordered.itertuples(index=False), start=2):
        ws.cell(r, 1, row.date).number_format = DATE
        ws.cell(r, 2, row.month).number_format = "[$-409]mmm yyyy"
        ws.cell(r, 3, row.branch)
        ws.cell(r, 4, row.product)
        ws.cell(r, 5, int(row.qty))
        ws.cell(r, 6, row.price).number_format = EUR
        ws.cell(r, 7, row.revenue).number_format = EUR
        ws.cell(r, 8, row.product_raw if row.product_raw != row.product else "")
        ws.cell(r, 9, row.source_file)
    _table(ws, "CleanData", 1, len(ordered) + 1, 1, len(headers))
    ws.freeze_panes = "A2"


def _quality_sheet(ws, logs: list[FileLog], fixes: list, outliers: pd.DataFrame) -> None:
    ws.sheet_view.showGridLines = False
    _section(ws, "A1", "Files read")
    headers = ["File", "Branch", "Type", "Date format", "Header row", "Rows read", "Blank rows", "Total rows",
               "Duplicates", "Prices calculated", "Rows used", "Notes"]
    _headers(ws, 2, headers, [34, 19, 9, 17, 11, 10, 11, 11, 11, 16, 10, 52])
    for r, log in enumerate(logs, start=3):
        values = [log.file, log.branch, log.file_type, log.date_order, log.header_row, log.rows_read,
                  log.blank_removed, log.total_rows_removed, log.duplicates_removed, log.prices_filled,
                  log.rows_kept, "; ".join(log.notes)]
        for c, value in enumerate(values, start=1):
            cell = ws.cell(r, c, value if value != 0 or c in (6, 11) else "–")
            if 5 <= c <= 11:
                cell.alignment = Alignment(horizontal="center")
            if c in (7, 8, 9, 10) and value:
                cell.font = Font(bold=True, color="9A6200")
    end = len(logs) + 2
    _table(ws, "FilesRead", 2, end, 1, len(headers))

    row = end + 3
    _section(ws, f"A{row}", "Product names corrected")
    _headers(ws, row + 1, ["Name in the file", "Corrected to", "Branch", "Rows"], [34, 19, 9, 17])
    for i, (raw, fixed, branch, n) in enumerate(fixes, start=row + 2):
        ws.cell(i, 1, repr(raw).strip("'") if raw != raw.strip() else raw)
        ws.cell(i, 2, fixed)
        ws.cell(i, 3, branch)
        ws.cell(i, 4, n)
    _table(ws, "NameFixes", row + 1, row + 1 + len(fixes), 1, 4)
    ws.column_dimensions["A"].width = 34

    row = row + len(fixes) + 4
    _section(ws, f"A{row}", "Unusual values to confirm with the branch")
    _headers(ws, row + 1, ["Date", "Branch", "Product", "Qty", "Typical day", "Note"], [34, 19, 9, 17, 11, 10])
    for i, o in enumerate(outliers.itertuples(), start=row + 2):
        ws.cell(i, 1, o.date).number_format = DATE
        ws.cell(i, 2, o.branch)
        ws.cell(i, 3, o.product)
        ws.cell(i, 4, int(o.qty)).font = Font(bold=True, color=RED)
        ws.cell(i, 5, int(o.typical))
        ws.cell(i, 6, f"about {o.qty / max(o.typical, 1):.0f}× a normal day: possible typing error. "
                      f"Kept in the totals until confirmed")
    if outliers.empty:
        ws.cell(row + 2, 1, "None found")
    _table(ws, "Unusual", row + 1, row + 1 + max(len(outliers), 1), 1, 6)


def _kpi_card(ws, col: int, label: str, value, note: str, fmt: Optional[str] = None,
              value_color: str = NAVY, size: int = 20) -> None:
    fill = PatternFill("solid", fgColor=LIGHT)
    accent = Side(style="thick", color=BLUE)
    for row in (5, 6, 7):
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
        for c in (col, col + 1):
            ws.cell(row, c).fill = fill
        ws.cell(row, col).border = Border(left=accent)
        ws.cell(row, col).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.cell(5, col, label.upper()).font = Font(size=9, bold=True, color=GREY)
    big = ws.cell(6, col, value)
    big.font = Font(size=size, bold=True, color=value_color)
    if fmt:
        big.number_format = fmt
    ws.cell(7, col, note).font = Font(size=9, color=GREY)


def _summary_sheet(ws, data, logs, fixes, outliers, months, run_at, source, branches_ws, n_branches,
                   products_ws, n_products, daily_ws, n_days) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    for col in range(2, 14):
        ws.column_dimensions[get_column_letter(col)].width = 11.5
    ws.row_dimensions[6].height = 30
    label = period_label(data)
    ws["B2"] = f"{config.COMPANY} · Sales Report {label}"
    ws["B2"].font = Font(size=20, bold=True, color=NAVY)
    ws["B3"] = (f"{len(logs)} files from {data['branch'].nunique()} branches   ·   "
                f"{data['date'].min():%d %b} – {data['date'].max():%d %b %Y}   ·   built {run_at:%d %b %Y, %H:%M}")
    ws["B3"].font = Font(size=10, color=GREY)

    by_branch = data.groupby("branch")["revenue"].sum().sort_values(ascending=False)
    by_product = data.groupby("product")["revenue"].sum().sort_values(ascending=False)
    by_month = data.groupby("month")["revenue"].sum()
    growth = by_month.iloc[-1] / by_month.iloc[0] - 1
    fixed_rows = (sum(l.blank_removed + l.total_rows_removed + l.duplicates_removed + l.prices_filled for l in logs)
                  + sum(n for *_, n in fixes))
    _kpi_card(ws, 2, "Revenue", data["revenue"].sum(), f"{label} · all branches", EUR0)
    _kpi_card(ws, 4, "Items sold", int(data["qty"].sum()), f"{len(data):,} sales lines", "#,##0")
    _kpi_card(ws, 6, "Top branch", by_branch.index[0], f"€{by_branch.iloc[0]:,.0f} · {by_branch.iloc[0] / by_branch.sum():.0%} of revenue",
              size=12)
    _kpi_card(ws, 8, "Best seller", by_product.index[0], f"€{by_product.iloc[0]:,.0f} revenue", size=12)
    _kpi_card(ws, 10, f"{months[-1]:%b} vs {months[0]:%b}", growth, "revenue growth", PCT,
              GREEN if growth >= 0 else RED)
    _kpi_card(ws, 12, "Fixes applied", fixed_rows, "see 'Data quality'", "#,##0")

    chart = BarChart()
    chart.type = "col"
    chart.title = "Revenue per branch and month"
    chart.title.overlay = False
    chart.add_data(Reference(branches_ws, min_col=2, max_col=1 + len(months), min_row=1, max_row=1 + n_branches),
                   titles_from_data=True)
    chart.set_categories(Reference(branches_ws, min_col=1, min_row=2, max_row=1 + n_branches))
    for series, color in zip(chart.series, MONTH_COLORS):
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.solidFill = color
    chart.gapWidth = 60
    chart.legend.position = "t"
    chart.legend.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.numFmt = "€#,##0"
    chart.y_axis.majorGridlines.spPr = None
    chart.width, chart.height = 12.6, 7.6
    ws.add_chart(chart, "B9")

    chart = BarChart()
    chart.type = "bar"
    chart.title = "Revenue by product"
    chart.title.overlay = False
    chart.add_data(Reference(products_ws, min_col=3, min_row=2, max_row=1 + n_products), titles_from_data=False)
    chart.set_categories(Reference(products_ws, min_col=1, min_row=2, max_row=1 + n_products))
    chart.legend = None
    chart.series[0].graphicalProperties.solidFill = BLUE
    chart.series[0].graphicalProperties.line.solidFill = BLUE
    chart.gapWidth = 45
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    for flag in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
        setattr(chart.dataLabels, flag, False)
    chart.x_axis.delete = False
    chart.y_axis.delete = True
    chart.y_axis.majorGridlines = None
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = float(by_product.iloc[0]) * 1.28
    chart.x_axis.scaling.orientation = "maxMin"
    chart.width, chart.height = 12.6, 7.6
    ws.add_chart(chart, "H9")

    line = LineChart()
    line.title = "Daily revenue, all branches"
    line.title.overlay = False
    line.add_data(Reference(daily_ws, min_col=3, min_row=1, max_row=1 + n_days), titles_from_data=True)
    line.set_categories(Reference(daily_ws, min_col=1, min_row=2, max_row=1 + n_days))
    series = line.series[0]
    series.graphicalProperties.line.solidFill = NAVY
    series.graphicalProperties.line.width = 19050
    series.marker.symbol = "none"
    series.smooth = False
    line.legend = None
    line.x_axis.delete = False
    line.y_axis.delete = False
    line.y_axis.numFmt = "€#,##0"
    line.x_axis.number_format = "[$-409]d mmm"
    line.x_axis.tickLblSkip = 14
    line.x_axis.tickMarkSkip = 7
    line.width, line.height = 26.2, 6.6
    ws.add_chart(line, "B25")

    _section(ws, "B39", "What the automation fixed")
    lines = []
    dupes = [l for l in logs if l.duplicates_removed]
    if dupes:
        lines.append(f"{sum(l.duplicates_removed for l in dupes)} rows exported twice removed "
                     f"({', '.join(sorted({l.branch for l in dupes}))})")
    totals = sum(l.total_rows_removed for l in logs)
    if totals:
        lines.append(f"{totals} subtotal / grand-total rows removed, so nothing is counted twice")
    if fixes:
        lines.append(f"{sum(n for *_, n in fixes):,} product names standardised "
                     f"(e.g. '{fixes[0][0]}' → {fixes[0][1]}, 'Expresso' → Espresso)")
    filled = sum(l.prices_filled for l in logs)
    if filled:
        lines.append(f"{filled:,} missing unit prices calculated from total ÷ quantity")
    titles = [l for l in logs if l.header_row > 1]
    if titles:
        lines.append(f"Title rows skipped and blank rows removed in {len(titles)} Excel files; "
                     "US dates (mm/dd) and Dutch number formats (3,40) read correctly")
    for o in outliers.itertuples():
        lines.append(f"⚠ Please confirm: {int(o.qty)} × {o.product} at {o.branch} on {o.date:%d %b} "
                     f"(a normal day is about {int(o.typical)})")
    for i, text in enumerate(lines, start=40):
        cell = ws.cell(i, 2, text if text.startswith("⚠") else f"✓  {text}")
        cell.font = Font(color=RED if text.startswith("⚠") else "333333", bold=text.startswith("⚠"))


def _notes_sheet(ws, run_at: datetime, source: str, n_files: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 100
    ws["A1"] = "About this report"
    ws["A1"].font = Font(size=16, bold=True, color=NAVY)
    rows = [
        ("Source folder", f"{source} ({n_files} files)"),
        ("Built", f"{run_at:%d %B %Y, %H:%M}"),
        ("How to update", "Put next month's exports in the folder and double-click run.bat. The whole report is rebuilt."),
        ("What is cleaned", "Header rows are found automatically; title, blank, subtotal and total rows are removed; "
                            "rows exported twice are removed; product names are matched to the official list "
                            "(typos, capitals, Dutch names); missing prices or totals are calculated; dates in "
                            "European, US and Excel format are read correctly."),
        ("Unusual values", "Daily quantities far outside a branch's normal range are listed on 'Data quality'. "
                           "They are kept in the totals until someone confirms them."),
        ("", ""),
        ("Prepared by", "Umut Doruk Öztürk · Python automation & data analysis"),
    ]
    for r, (key, value) in enumerate(rows, start=3):
        ws.cell(r, 1, key).font = Font(bold=True, color=NAVY)
        cell = ws.cell(r, 2, value)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 1).alignment = Alignment(vertical="top")
