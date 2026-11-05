"""Read one branch export, whatever its format, into standard columns.

Nothing here is written for a specific branch: the header row, the column names (English or
Dutch), the CSV delimiter, the decimal separator and the date order are all detected.
"""
import csv
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd

from . import config

TOTAL_ROW = re.compile(r"(?i)\b(?:sub)?total\b|\btotaal\b|\bsum\b")


@dataclass
class FileLog:
    file: str
    branch: str = ""
    file_type: str = ""
    header_row: int = 0
    rows_read: int = 0
    blank_removed: int = 0
    total_rows_removed: int = 0
    duplicates_removed: int = 0
    prices_filled: int = 0
    date_order: str = ""
    rows_kept: int = 0
    notes: list[str] = field(default_factory=list)


def _norm(text) -> str:
    return re.sub(r"[^a-z0-9€()]+", " ", str(text).lower()).strip()


def parse_number(value) -> Optional[float]:
    """3.40 / '3,40' / '€3.40' / '1.234,56' / '2 pcs' -> float."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = re.sub(r"[^\d,.\-]", "", str(value))
    if not re.search(r"\d", s):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if len(s) - s.rfind(",") - 1 in (1, 2) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def read_table(path: Path) -> tuple[pd.DataFrame, str]:
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, header=None, dtype=object), "Excel"
    sample = path.read_text(encoding="utf-8-sig", errors="replace")[:4096]
    delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    df = pd.read_csv(path, header=None, dtype=str, sep=delimiter, encoding="utf-8-sig", keep_default_na=False)
    return df.replace("", None), f"CSV ({'tab' if delimiter == chr(9) else delimiter})"


def find_header(raw: pd.DataFrame) -> tuple[int, dict[str, int]]:
    """The first row in which at least 3 cells are known column names."""
    lookup = {_norm(name): std for std, names in config.COLUMNS.items() for name in names}
    for i in range(min(len(raw), 30)):
        mapping = {}
        for j, cell in enumerate(raw.iloc[i]):
            std = lookup.get(_norm(cell)) if cell is not None else None
            if std and std not in mapping:
                mapping[std] = j
        if len(mapping) >= 3 and "product" in mapping:
            return i, mapping
    raise ValueError("no header row found (expected columns like Date, Product, Qty, Total)")


def branch_of(path: Path) -> str:
    name = path.stem.lower()
    for branch, words in config.BRANCHES.items():
        if any(w in name for w in words):
            return branch
    return path.stem


def _date_order(values: list[str]) -> str:
    """'DMY' or 'MDY' for slash/dash dates: a first part above 12 means day-first, a second part above 12 month-first."""
    firsts, seconds = [], []
    for v in values:
        m = re.match(r"^\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", v)
        if m:
            firsts.append(int(m.group(1)))
            seconds.append(int(m.group(2)))
    if any(s > 12 for s in seconds) and not any(f > 12 for f in firsts):
        return "MDY"
    return "DMY"


def parse_dates(series: pd.Series) -> tuple[pd.Series, str]:
    texts = [str(v) for v in series if isinstance(v, str)]
    order = _date_order(texts)

    def one(v) -> Optional[date]:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        if isinstance(v, (datetime, pd.Timestamp)):
            return v.date()
        if isinstance(v, date):
            return v
        s = str(v).strip()
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", s)
        if m:
            a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return date(y, b, a) if order == "DMY" else date(y, a, b)
        return None

    kinds = {"Excel dates" if not texts else ("ISO (yyyy-mm-dd)" if all(re.match(r"^\d{4}-", t) for t in texts)
                                              else ("dd-mm-yyyy" if order == "DMY" else "mm/dd/yyyy (US)"))}
    return series.map(one), kinds.pop()


def load_file(path: Path) -> tuple[pd.DataFrame, FileLog]:
    log = FileLog(file=path.name, branch=branch_of(path))
    raw, log.file_type = read_table(path)
    header, cols = find_header(raw)
    log.header_row = header + 1
    if header > 0:
        log.notes.append(f"{header} title row(s) above the header skipped")
    data = raw.iloc[header + 1:].reset_index(drop=True)
    log.rows_read = len(data)

    blank = data.isna().all(axis=1)
    log.blank_removed = int(blank.sum())
    data = data[~blank]

    get = lambda name: data.iloc[:, cols[name]] if name in cols else pd.Series([None] * len(data), index=data.index)
    product = get("product").map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
    dates, log.date_order = parse_dates(get("date"))
    totals = product.str.contains(TOTAL_ROW) | (dates.isna() & (product.str.strip() == ""))
    log.total_rows_removed = int(totals.sum())

    out = pd.DataFrame({
        "date": dates, "product_raw": product, "qty": get("qty").map(parse_number),
        "price": get("price").map(parse_number), "revenue": get("revenue").map(parse_number),
    })[~totals]
    unknown_dates = int(out["date"].isna().sum())
    if unknown_dates:
        log.notes.append(f"{unknown_dates} rows without a readable date removed")
        out = out[out["date"].notna()]

    out[["qty", "price", "revenue"]] = out[["qty", "price", "revenue"]].astype(float)
    missing_price = out["price"].isna() & out["qty"].gt(0) & out["revenue"].notna()
    if missing_price.any():
        out.loc[missing_price, "price"] = (out.loc[missing_price, "revenue"] / out.loc[missing_price, "qty"]).round(2)
        log.prices_filled = int(missing_price.sum())
        log.notes.append("no price column: unit price calculated as total ÷ quantity")
    missing_revenue = out["revenue"].isna() & out["qty"].notna() & out["price"].notna()
    if missing_revenue.any():
        out.loc[missing_revenue, "revenue"] = (out.loc[missing_revenue, "qty"] * out.loc[missing_revenue, "price"]).round(2)
        log.notes.append(f"{int(missing_revenue.sum())} missing totals calculated as quantity × price")

    dupes = out.duplicated(subset=["date", "product_raw", "qty", "revenue"])
    log.duplicates_removed = int(dupes.sum())
    out = out[~dupes]

    out["branch"] = log.branch
    out["source_file"] = path.name
    log.rows_kept = len(out)
    return out.reset_index(drop=True), log
