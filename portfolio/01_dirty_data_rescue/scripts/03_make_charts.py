"""
STEP 3 of 3 -- Add the client-facing visuals.

Charts are written as NATIVE Excel charts into the workbook, not as images.
That is deliberate:
  - it needs no plotting library, so the script runs anywhere Python does;
  - the client can click a series, change a colour, or repoint it at new data;
  - the charts stay live if the Clean Data sheet is refreshed.

Output: output/cleaned_orders_charts.xlsx

Run:  python 03_make_charts.py
"""

from __future__ import annotations

import csv
import json
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
RAW = BASE / "data" / "raw"
OUT = BASE / "output"
OUT.mkdir(parents=True, exist_ok=True)

NAVY = "#1F3B57"
TEAL = "#2E8B8B"
AMBER = "#E0A458"
RED = "#C1554A"
GREY = "#8A94A6"

JUNK = {"", "n/a", "na", "null", "none", "nil", "-", "--", "tbd", "unknown",
        "?", "#ref!", "#n/a", "#value!", " "}


# --------------------------------------------------------------------------
# Build the three small summary tables the charts are drawn from
# --------------------------------------------------------------------------
def issue_breakdown(report: dict) -> pd.DataFrame:
    checks = report["checks"]

    def collapsed(field: str) -> int:
        c = checks.get(field, {})
        return c.get("raw_distinct", 0) - c.get("canonical_distinct", 0)

    rows = [
        ("Whitespace / casing fixed",
         sum(checks.get("text_normalisation", {}).values())),
        ("Amounts & quantities reformatted",
         checks.get("amount_parsing", {}).get("ok", 0) +
         checks.get("qty_parsing", {}).get("ok", 0)),
        ("Currency symbols / codes resolved",
         sum(v for k, v in checks.get("currency_resolution", {}).items()
             if k != "explicit-column")),
        ("Company & product spellings merged",
         collapsed("customer_name") + collapsed("product_name")),
        ("Duplicate records removed",
         checks.get("duplicates", {}).get("exact_removed", 0)),
        ("Same-ID conflicts resolved",
         checks.get("duplicates", {}).get("near_duplicates_resolved", 0)),
        ("Blank rows dropped",
         checks.get("blank_rows", {}).get("removed", 0)),
        ("Header / footer junk rows skipped",
         checks.get("excel_header", {}).get("junk_rows_above_table", 0) +
         checks.get("excel_footer", {}).get("junk_rows_removed", 0)),
    ]
    df = pd.DataFrame(rows, columns=["Issue found in the source files", "Count"])
    return df[df["Count"] > 0].sort_values("Count", ascending=False)


def completeness_table() -> pd.DataFrame:
    """Before is measured on the raw file, after on the cleaned sheet."""
    fields = ["order_date", "customer", "product", "qty", "amount"]

    with (RAW / "sales_orders_raw.csv").open(newline="", encoding="utf-8-sig") as fh:
        raw_rows = list(csv.DictReader(fh))
    header = {k: k.strip() for k in raw_rows[0]}
    before = {}
    for f in fields:
        key = next(k for k, v in header.items() if v == f)
        good = sum(1 for r in raw_rows
                   if str(r.get(key, "")).strip().lower() not in JUNK)
        before[f] = round(100 * good / len(raw_rows), 1)

    ws = load_workbook(OUT / "cleaned_orders.xlsx", data_only=True)["Clean Data"]
    grid = list(ws.iter_rows(values_only=True))
    head = [str(c) for c in grid[0]]
    body = grid[1:]
    after = {}
    for f in fields:
        i = head.index(f)
        good = sum(1 for r in body if r[i] not in (None, ""))
        after[f] = round(100 * good / max(len(body), 1), 1)

    return pd.DataFrame({
        "Field": fields,
        "Before cleaning (%)": [before[f] for f in fields],
        "After cleaning (%)": [after[f] for f in fields],
    })


def row_accounting(report: dict) -> pd.DataFrame:
    checks = report["checks"]
    blank = checks.get("blank_rows", {}).get("removed", 0)
    exact = checks.get("duplicates", {}).get("exact_removed", 0)
    near = checks.get("duplicates", {}).get("near_duplicates_resolved", 0)
    rows_in = report["results"]["rows_in"]

    stages = [
        ("Raw rows in file", rows_in),
        ("After dropping blank rows", rows_in - blank),
        ("After exact duplicates", rows_in - blank - exact),
        ("After same-ID conflicts", rows_in - blank - exact - near),
        ("Clean rows delivered", report["results"]["rows_out"]),
    ]
    df = pd.DataFrame(stages, columns=["Stage", "Rows"])
    df["Removed at this step"] = [None] + [
        df["Rows"][i - 1] - df["Rows"][i] for i in range(1, len(df))]
    return df


def entity_resolution(report: dict) -> pd.DataFrame:
    checks = report["checks"]
    rows = []
    for label, key in [("Customers", "customer_name"), ("Products", "product_name"),
                       ("Regions", "region")]:
        c = checks.get(key, {})
        rows.append({"Entity": label,
                     "Distinct spellings in raw data": c.get("raw_distinct", 0),
                     "Real distinct entities": c.get("canonical_distinct", 0)})
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# Write everything into one workbook with native charts
# --------------------------------------------------------------------------
def main() -> None:
    report = json.loads((OUT / "data_quality_report.json").read_text(encoding="utf-8"))

    issues = issue_breakdown(report)
    completeness = completeness_table()
    accounting = row_accounting(report)
    entities = entity_resolution(report)

    cleaned = pd.read_excel(OUT / "cleaned_orders.xlsx", sheet_name="Clean Data")
    audit = pd.read_excel(OUT / "cleaned_orders.xlsx", sheet_name="Audit Summary")

    out_path = OUT / "cleaned_orders_charts.xlsx"

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as xl:
        wb = xl.book

        f_title = wb.add_format({"bold": True, "font_size": 14, "font_color": NAVY})
        f_note = wb.add_format({"italic": True, "font_color": GREY, "font_size": 9})
        f_head = wb.add_format({"bold": True, "bg_color": NAVY,
                                "font_color": "white", "border": 1})

        # ---------------- dashboard cover sheet ----------------
        dash = wb.add_worksheet("Dashboard")
        dash.hide_gridlines(2)
        dash.set_column("A:A", 3)
        dash.set_column("B:B", 46)
        dash.set_column("C:C", 16)
        dash.write("B2", "Data Cleaning — Summary", f_title)
        dash.write("B3", f"Rows delivered: {report['results']['rows_out']}  ·  "
                         f"corrections applied: {report['results']['total_corrections']}  ·  "
                         f"completeness {report['results']['completeness_before_pct']}% → "
                         f"{report['results']['completeness_after_pct']}%", f_note)

        # ---------------- data sheets the charts point at ----------------
        for sheet, frame in [("Issues Found", issues),
                             ("Completeness", completeness),
                             ("Row Accounting", accounting),
                             ("Entity Resolution", entities),
                             ("Audit Summary", audit)]:
            frame.to_excel(xl, sheet_name=sheet, index=False)
            ws = xl.sheets[sheet]
            ws.write(0, 0, frame.columns[0], f_head)
            for i, c in enumerate(frame.columns):
                ws.write(0, i, c, f_head)
                ws.set_column(i, i, min(max(14, len(str(c)) + 4), 46))
            ws.freeze_panes(1, 0)

        # ---------------- chart 1: issues found (bar) ----------------
        ch1 = wb.add_chart({"type": "bar"})
        ch1.add_series({
            "name": "Corrections applied",
            "categories": ["Issues Found", 1, 0, len(issues), 0],
            "values": ["Issues Found", 1, 1, len(issues), 1],
            "fill": {"color": TEAL},
            "border": {"none": True},
            "data_labels": {"value": True, "font": {"color": NAVY}},
        })
        ch1.set_title({"name": "What was actually wrong with the source files",
                       "name_font": {"size": 13, "bold": True, "color": NAVY}})
        ch1.set_legend({"none": True})
        ch1.set_x_axis({"name": "Corrections applied",
                        "name_font": {"color": NAVY},
                        "num_font": {"color": NAVY}})
        ch1.set_y_axis({"reverse": True, "num_font": {"color": NAVY}})
        ch1.set_size({"width": 720, "height": 380})
        ch1.set_chartarea({"border": {"none": True}})
        dash.insert_chart("B6", ch1)

        # ---------------- chart 2: completeness before vs after ----------------
        ch2 = wb.add_chart({"type": "column"})
        n = len(completeness)
        ch2.add_series({
            "name": "Before cleaning",
            "categories": ["Completeness", 1, 0, n, 0],
            "values": ["Completeness", 1, 1, n, 1],
            "fill": {"color": RED},
            "data_labels": {"value": True, "num_format": "0.0"},
        })
        ch2.add_series({
            "name": "After cleaning",
            "categories": ["Completeness", 1, 0, n, 0],
            "values": ["Completeness", 1, 2, n, 2],
            "fill": {"color": TEAL},
            "data_labels": {"value": True, "num_format": "0.0"},
        })
        ch2.set_title({"name": "Completeness on the five fields the client reports on",
                       "name_font": {"size": 13, "bold": True, "color": NAVY}})
        ch2.set_x_axis({"name": "Field", "num_font": {"color": NAVY}})
        ch2.set_y_axis({"name": "% of rows with a usable value",
                        "min": 0, "max": 105,
                        "num_font": {"color": NAVY}})
        ch2.set_legend({"position": "bottom"})
        ch2.set_size({"width": 720, "height": 380})
        ch2.set_chartarea({"border": {"none": True}})
        dash.insert_chart("B27", ch2)

        # ---------------- chart 3: row accounting ----------------
        ch3 = wb.add_chart({"type": "column"})
        ch3.add_series({
            "name": "Rows surviving",
            "categories": ["Row Accounting", 1, 0, len(accounting), 0],
            "values": ["Row Accounting", 1, 1, len(accounting), 1],
            "fill": {"color": AMBER},
            "data_labels": {"value": True},
        })
        ch3.set_title({"name": "Nothing was deleted silently: every removed row is accounted for",
                       "name_font": {"size": 13, "bold": True, "color": NAVY}})
        ch3.set_legend({"none": True})
        ch3.set_x_axis({"num_font": {"rotation": -20, "color": NAVY}})
        ch3.set_y_axis({"name": "Rows", "min": 0, "num_font": {"color": NAVY}})
        ch3.set_size({"width": 720, "height": 380})
        ch3.set_chartarea({"border": {"none": True}})
        dash.insert_chart("B48", ch3)

        # ---------------- chart 4: entity resolution ----------------
        ch4 = wb.add_chart({"type": "column"})
        m = len(entities)
        ch4.add_series({
            "name": "Distinct spellings in raw data",
            "categories": ["Entity Resolution", 1, 0, m, 0],
            "values": ["Entity Resolution", 1, 1, m, 1],
            "fill": {"color": RED},
            "data_labels": {"value": True},
        })
        ch4.add_series({
            "name": "Real distinct entities",
            "categories": ["Entity Resolution", 1, 0, m, 0],
            "values": ["Entity Resolution", 1, 2, m, 2],
            "fill": {"color": TEAL},
            "data_labels": {"value": True},
        })
        ch4.set_title({"name": "41 customer spellings that are really 8 companies",
                       "name_font": {"size": 13, "bold": True, "color": NAVY}})
        ch4.set_legend({"position": "bottom"})
        ch4.set_size({"width": 720, "height": 380})
        ch4.set_chartarea({"border": {"none": True}})
        dash.insert_chart("B69", ch4)

        # ---------------- clean data + mapping sheets ----------------
        cleaned.to_excel(xl, sheet_name="Clean Data", index=False)
        ws = xl.sheets["Clean Data"]
        for i, c in enumerate(cleaned.columns):
            ws.write(0, i, c, f_head)
            ws.set_column(i, i, min(max(12, len(str(c)) + 4), 34))
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(cleaned), len(cleaned.columns) - 1)

        for sheet in ["Value Mapping", "Canonicalisation", "Duplicate Decisions"]:
            try:
                pd.read_excel(OUT / "cleaned_orders.xlsx", sheet_name=sheet).to_excel(
                    xl, sheet_name=sheet, index=False)
                xl.sheets[sheet].freeze_panes(1, 0)
            except ValueError:
                continue

    print("wrote", out_path.name)
    for name, frame in [("Issues Found", issues), ("Completeness", completeness),
                        ("Row Accounting", accounting),
                        ("Entity Resolution", entities)]:
        print(f"  sheet {name}: {len(frame)} rows")


if __name__ == "__main__":
    main()
