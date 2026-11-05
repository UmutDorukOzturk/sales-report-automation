from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from sales_report.clean import flag_outliers, standardise_products
from sales_report.load import load_file, parse_number

SAMPLE = Path(__file__).resolve().parents[1] / "sample_data"


@pytest.mark.parametrize("text, expected", [
    ("3,40", 3.40), ("€3.40", 3.40), ("1.234,56", 1234.56), ("1,234.56", 1234.56), ("2 pcs", 2.0),
    (12, 12.0), ("", None), (None, None),
])
def test_parse_number(text, expected):
    assert parse_number(text) == expected


def test_semicolon_csv_with_dutch_headers_and_decimal_commas(tmp_path):
    f = tmp_path / "oost juli.csv"
    f.write_text("Datum;Product;Aantal;Prijs;Omzet\n01-07-2026;Thee;3;2,80;8,40\n13-07-2026;Thee;1;2,80;2,80\n",
                 encoding="utf-8")
    df, log = load_file(f)
    assert log.branch == "Amsterdam Oost" and log.date_order == "dd-mm-yyyy"
    assert list(df["date"]) == [date(2026, 7, 1), date(2026, 7, 13)]
    assert list(df["revenue"]) == [8.40, 2.80]


def test_us_dates_euro_signs_and_duplicates(tmp_path):
    f = tmp_path / "DenHaag-export.csv"
    rows = "07/13/2026,Tea,2,€2.80,€5.60\n"
    f.write_text("Transaction Date,Product Name,Qty Sold,Unit Price,Sales\n" + rows + rows, encoding="utf-8")
    df, log = load_file(f)
    assert df.loc[0, "date"] == date(2026, 7, 13)
    assert log.duplicates_removed == 1 and len(df) == 1


def test_rotterdam_subtotals_and_typos_are_cleaned():
    df, log = load_file(SAMPLE / "ROTTERDAM_07_2026.xlsx")
    assert log.total_rows_removed == 6 and log.blank_removed == 1
    assert not df["product_raw"].str.contains("total", case=False).any()
    df, fixes, unknown = standardise_products(df)
    assert set(df["product"]) <= {"Espresso", "Americano", "Cappuccino", "Latte", "Flat White", "Tea",
                                  "Hot Chocolate", "Croissant", "Banana Bread", "Sandwich"}
    assert unknown == []


def test_missing_price_column_is_calculated():
    df, log = load_file(SAMPLE / "Haarlem Jul26.xlsx")
    assert log.prices_filled == len(df)
    assert (abs(df["qty"] * df["price"] - df["revenue"]) < 0.01).all()


def test_typing_error_is_flagged_but_normal_days_are_not():
    days = pd.date_range("2026-07-01", periods=30).date
    qty = [40, 38, 42, 41, 39] * 6
    qty[10] = 410
    df = pd.DataFrame({"date": days, "branch": "Rotterdam", "product": "Latte", "qty": qty})
    flagged = flag_outliers(df)
    assert list(flagged["qty"]) == [410]


def test_whole_sample_folder():
    frames = [load_file(p)[0] for p in sorted(SAMPLE.iterdir())]
    data = pd.concat(frames, ignore_index=True)
    assert len(data) == 6 * (31 + 31 + 30) * 10          # 6 branches, 92 days, 10 products
    assert data["date"].min() == date(2026, 7, 1) and data["date"].max() == date(2026, 9, 30)
