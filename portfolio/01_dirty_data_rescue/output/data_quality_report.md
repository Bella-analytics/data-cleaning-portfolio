# Data Quality Report

Generated: 2026-10-10T22:17:32

## 1. Sources

| File | Rows in |
|---|---|
| `sales_orders_raw.csv` | 73 |
| `customer_master_raw.csv` | 19 |
| `inventory_jp_raw.xlsx` | 12 |

## 2. Outcome

- Rows in: **73**
- Rows out: **62** (removed 11)
- Total corrections applied: **499**
- Completeness on key fields: **97.26% → 100.0%**
- Distinct customers: 41 spellings → **8 real companies**
- Distinct products: 26 spellings → **5 real SKUs**
- Dates that could not be parsed: 0 (kept, flagged, never dropped)
- Orders with no matching customer record: 0 (kept, flagged)

## 3. Every check that ran

| Check | Outcome | Count |
|---|---|---|
| `amount_parsing` | ok | 73 |
| `blank_rows` | removed | 4 |
| `csv_header` | columns_normalised | 15 |
| `currency_resolution` | explicit-column | 73 |
| `customer_merge` | master_rows_after_dedup | 8 |
| `customer_merge` | orders_without_master_record | 0 |
| `customer_name` | raw_distinct | 41 |
| `customer_name` | canonical_distinct | 8 |
| `date_parsing` | mon-d-y | 5 |
| `date_parsing` | iso-slash | 13 |
| `date_parsing` | iso | 7 |
| `date_parsing` | day-first | 14 |
| `date_parsing` | numeric-assumed-day-first | 19 |
| `date_parsing` | d-mon-y | 10 |
| `date_parsing` | month-first | 5 |
| `duplicates` | exact_removed | 7 |
| `duplicates` | near_duplicates_resolved | 4 |
| `excel_footer` | junk_rows_removed | 4 |
| `excel_header` | junk_rows_above_table | 5 |
| `product_name` | raw_distinct | 26 |
| `product_name` | canonical_distinct | 5 |
| `qty_parsing` | ok | 63 |
| `qty_parsing` | missing | 10 |
| `region` | raw_distinct | 14 |
| `region` | canonical_distinct | 5 |
| `status_normalisation` | title_cased | 61 |

## 4. Decisions a human may want to overrule

These are judgement calls, not facts. They are listed here so the
client can audit them instead of discovering them later.

- **Ambiguous numeric dates** (`03/04/2024`): the file mixes `15/03/2024` and `03/15/2024`, so a value where both readings are valid is unresolvable from the data alone. One convention was assumed for the whole file and every row is labelled with the method used in the audit sheet.
- **Near-duplicate order IDs** (same ID, different date/amount): kept the most complete record; all candidates are in the `Duplicate Decisions` sheet.
- **Company names merged by similarity** (edit distance ≤ 3 after normalisation). Check the `Canonicalisation` sheet before using customer counts in any report.
- **Missing values were never filled in.** A blank stays blank and is flagged, so downstream totals are not silently distorted.

## 5. Column notes

- `order_id`: str, 0 missing
- `order_date`: object, 0 missing
- `customer`: str, 0 missing
- `product`: str, 0 missing
- `qty`: Int64, 8 missing
- `unit_price`: float64, 8 missing
- `amount`: float64, 0 missing
- `currency`: str, 0 missing
- `region`: str, 0 missing
- `sales_rep`: str, 0 missing
- `status`: str, 0 missing
- `credit_limit`: str, 0 missing
- `contact_email`: str, 0 missing
- `flag_date_unparsed`: bool, 0 missing
- `flag_amount_missing`: bool, 0 missing
- `flag_qty_zero_or_missing`: bool, 0 missing
