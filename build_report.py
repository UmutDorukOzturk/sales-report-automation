"""Merge every branch's sales export into one clean, formatted Excel report.

Usage:
    python build_report.py sample_data              # all .csv / .xlsx files in the folder
    python build_report.py "C:\\Exports\\Q3" --out-dir reports
"""
import argparse
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

from sales_report.clean import flag_outliers, standardise_products
from sales_report.load import load_file
from sales_report.report import period_label, write_report

ROOT = Path(__file__).resolve().parent
EXTENSIONS = {".csv", ".xlsx", ".xlsm", ".xls"}


def step(n: int, text: str, result: str) -> None:
    print(f"[{n}/4] {(text + ' ').ljust(30, '.')} {result}", flush=True)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder", type=Path, help="folder with the branch exports")
    ap.add_argument("--out-dir", type=Path, default=ROOT / "output")
    args = ap.parse_args()

    started, run_at = time.time(), datetime.now()
    files = sorted((p for p in args.folder.iterdir() if p.suffix.lower() in EXTENSIONS and not p.name.startswith("~$")),
                   key=lambda p: p.name.lower())
    if not files:
        raise SystemExit(f"No .csv or .xlsx files found in {args.folder}")
    args.out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Sales report builder · {len(files)} files in '{args.folder.name}'\n")

    frames, logs = [], []
    for path in files:
        try:
            frame, log = load_file(path)
        except ValueError as exc:
            print(f"      skipped {path.name}: {exc}")
            continue
        frames.append(frame)
        logs.append(log)
    data = pd.concat(frames, ignore_index=True)
    step(1, "Reading files", f"{len(logs)} files · {data['branch'].nunique()} branches · "
                             f"{sum(l.rows_read for l in logs):,} rows")

    data, fixes, unknown = standardise_products(data)
    removed = sum(l.blank_removed + l.total_rows_removed for l in logs)
    dupes = sum(l.duplicates_removed for l in logs)
    step(2, "Cleaning", f"{removed} blank/total rows · {dupes} duplicates · {sum(n for *_, n in fixes):,} names fixed")
    if unknown:
        print(f"      products not in the official list (kept as they are): {', '.join(unknown)}")

    outliers = flag_outliers(data)
    step(3, "Checking for unusual values", f"{len(outliers)} to confirm with the branch")

    label = period_label(data).replace(" – ", "-").replace(" ", "_")
    out = args.out_dir / f"sales_report_{label}.xlsx"
    write_report(data, logs, fixes, outliers, out, run_at, args.folder.name)
    clean_csv = args.out_dir / f"clean_data_{label}.csv"
    (data.sort_values(["date", "branch", "product"])
         [["date", "branch", "product", "qty", "price", "revenue", "product_raw", "source_file"]]
         .to_csv(clean_csv, index=False, encoding="utf-8-sig"))
    shown = out.relative_to(ROOT) if out.is_relative_to(ROOT) else out
    step(4, "Writing Excel report", str(shown))

    print(f"\nDone in {time.time() - started:.1f}s · {len(data):,} clean rows · "
          f"€{data['revenue'].sum():,.0f} revenue · {period_label(data)}")


if __name__ == "__main__":
    main()
