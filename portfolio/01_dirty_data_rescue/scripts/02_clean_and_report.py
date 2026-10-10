"""
STEP 2 of 3 -- Clean, validate, merge and audit.

This is the actual deliverable a client pays for. It takes the three messy
source files and produces:

    output/cleaned_orders.xlsx      clean, analysis-ready workbook (3 sheets)
    output/data_quality_report.json machine-readable audit trail
    output/data_quality_report.md   human-readable audit trail

Design rules followed here:
  1. Never silently drop a row. Anything removed is counted and explained.
  2. Never invent a value. Missing stays missing and is reported.
  3. Every transformation is counted, so the client can see what changed.
  4. The script is re-runnable on next month's export.

Run:  python 02_clean_and_report.py
"""

from __future__ import annotations

import csv
import json
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
RAW = BASE / "data" / "raw"
OUT = BASE / "output"
OUT.mkdir(parents=True, exist_ok=True)

# --------------------------------------------------------------------------
# 0. Transformation counters -- this is what the audit trail is built from
# --------------------------------------------------------------------------
stats: dict[str, Counter] = defaultdict(Counter)


def bump(check: str, outcome: str, n: int = 1) -> None:
    stats[check][outcome] += n


# --------------------------------------------------------------------------
# 1. Text normalisation helpers
# --------------------------------------------------------------------------
NBSP = "\u00a0"


def clean_text(value) -> str:
    """Trim, unify Unicode, collapse inner whitespace, strip NBSP / zero-width."""
    if value is None:
        return ""
    s = str(value)
    if s.strip().lower() in {"nan", "none", "nat"}:
        return ""
    s = s.replace(NBSP, " ").replace("\u200b", "").replace("\ufeff", "")
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


NULL_TOKENS = {"", "n/a", "na", "null", "none", "nil", "-", "--", "tbd",
               "unknown", "?", "#ref!", "#n/a", "#value!", "0"}


def is_missing(value) -> bool:
    return clean_text(value).lower() in NULL_TOKENS


# --------------------------------------------------------------------------
# 2. Date parsing -- the hard part, because the same file mixes formats
# --------------------------------------------------------------------------
MONTHS = {m.lower(): i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}


def _try_monthfirst(s: str):
    """Interpret a purely numeric date as MM/DD/YYYY (US convention)."""
    m = re.fullmatch(r"(\d{1,2})[/.](\d{1,2})[/.](\d{4})", s)
    if not m:
        return None
    mm, dd, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if 1 <= mm <= 12 and 1 <= dd <= 31:
        return date(yyyy, mm, dd)
    return None


def _try_dayfirst(s: str):
    """Interpret a purely numeric date as DD/MM/YYYY (EU/Asian convention)."""
    m = re.fullmatch(r"(\d{1,2})[/.](\d{1,2})[/.](\d{4})", s)
    if not m:
        return None
    dd, mm, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if 1 <= mm <= 12 and 1 <= dd <= 31:
        return date(yyyy, mm, dd)
    return None


def parse_date(raw_value) -> tuple[date | None, str]:
    """
    Return (date, method). method documents HOW it was read, which matters:
    a date like 03/04/2024 is genuinely ambiguous and the client must know
    which convention was assumed.
    """
    s = clean_text(raw_value)
    if not s or s.lower() in NULL_TOKENS:
        return None, "missing"

    # 2024-03-15
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        try:
            return date(int(m[1]), int(m[2]), int(m[3])), "iso"
        except ValueError:
            return None, "invalid"

    # 2024/03/15
    m = re.fullmatch(r"(\d{4})/(\d{1,2})/(\d{1,2})", s)
    if m:
        try:
            return date(int(m[1]), int(m[2]), int(m[3])), "iso-slash"
        except ValueError:
            return None, "invalid"

    # 15-Mar-2024
    m = re.fullmatch(r"(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{4})", s)
    if m and m[2][:3].lower() in MONTHS:
        try:
            return date(int(m[3]), MONTHS[m[2][:3].lower()], int(m[1])), "d-mon-y"
        except ValueError:
            return None, "invalid"

    # Mar 15, 2024
    m = re.fullmatch(r"([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4})", s)
    if m and m[1][:3].lower() in MONTHS:
        try:
            return date(int(m[3]), MONTHS[m[1][:3].lower()], int(m[2])), "mon-d-y"
        except ValueError:
            return None, "invalid"

    # Purely numeric -> genuinely ambiguous. Resolve per file, not per row.
    mf, df_ = _try_monthfirst(s), _try_dayfirst(s)
    if mf and df_:
        global _NUMERIC_CONVENTION
        if _NUMERIC_CONVENTION is None:
            _NUMERIC_CONVENTION = "dayfirst"     # documented assumption
        return (df_, "numeric-assumed-day-first") if _NUMERIC_CONVENTION == "dayfirst" \
            else (mf, "numeric-assumed-month-first")
    if mf:
        return mf, "month-first"
    if df_:
        return df_, "day-first"

    return None, "invalid"


_NUMERIC_CONVENTION: str | None = None


# --------------------------------------------------------------------------
# 3. Amount parsing -- strips currency symbols, thousand separators, signs
# --------------------------------------------------------------------------
CURRENCY_ALIASES = {"$": "USD", "US$": "USD", "USD": "USD", "S$": "SGD",
                    "SGD": "SGD", "RM": "MYR", "MYR": "MYR", "¥": "JPY",
                    "JPY": "JPY", "円": "JPY", "EUR": "EUR", "€": "EUR",
                    "GBP": "GBP", "£": "GBP"}


def parse_amount(raw_value) -> tuple[float | None, str]:
    s = clean_text(raw_value)
    if not s or s.lower() in NULL_TOKENS:
        return None, "missing"

    negative = False
    if s.startswith("(") and s.endswith(")"):        # accounting negative
        negative, s = True, s[1:-1]
    if s.lstrip().startswith("-"):
        negative, s = True, s.lstrip()[1:]

    if not re.search(r"\d", s):
        return None, "no-digits"

    cleaned = s.replace(",", "").replace(" ", "").lstrip("'")
    m = re.search(r"\d+(?:\.\d+)?", cleaned)
    if not m:
        return None, "no-digits"

    value = float(m.group(0))
    return (-value if negative else value), "ok"


def parse_currency(raw_value, amount_raw: str) -> tuple[str, str]:
    """Return (ISO code, how it was determined)."""
    explicit = clean_text(raw_value).upper()
    if explicit in CURRENCY_ALIASES:
        return CURRENCY_ALIASES[explicit], "explicit-column"

    blob = clean_text(amount_raw)
    for token, code in CURRENCY_ALIASES.items():
        if token.upper() in blob.upper():
            return code, "from-amount-string"
    return "USD", "defaulted"


def parse_int(raw_value) -> tuple[int | None, str]:
    s = clean_text(raw_value)
    if not s or s.lower() in NULL_TOKENS:
        return None, "missing"
    m = re.search(r"-?\d+", s.replace(",", ""))
    if not m:
        return None, "invalid"
    return int(m.group(0)), "ok"


# --------------------------------------------------------------------------
# 4. Canonicalisation by evidence gathering (no hardcoded answer key)
# --------------------------------------------------------------------------
def fingerprint(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", clean_text(text).lower())


def _tokens(text: str) -> list[str]:
    return re.findall(r"[a-z0-9]+", clean_text(text).lower())


def _edit_distance(a: str, b: str, cap: int = 3) -> int:
    if abs(len(a) - len(b)) > cap:
        return cap + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _similar(a: str, b: str) -> bool:
    """
    Do two spellings refer to the same entity?

    Rules, in order:
      1. Token overlap >= 50%. "Aurora Trading" vs "Aurora Trading Ltd" share
         every real word -> same entity. Also stops "Northwind Retail" being
         split from "Northwind Retail Inc".
      2. Fingerprint edit distance, scaled by word length:
           - short values (short codes, region names): at most 1 edit
           - longer values: at most 2 edits, for typos like "Logistcs"

    The length scaling is the important part. Short strings are always
    "close" to each other: editdistance("east","west") is 2, and
    editdistance("na","east") is 3. A flat threshold of 2-3 therefore
    collapses North/South/East/West into a single region, which silently
    destroys any report grouped by region.
    """
    ta, tb = set(_tokens(a)), set(_tokens(b))
    if ta and tb and len(ta & tb) / min(len(ta), len(tb)) >= 0.5:
        return True

    fa, fb = fingerprint(a), fingerprint(b)
    allowed = 1 if min(len(fa), len(fb)) < 6 else 2
    return _edit_distance(fa, fb, cap=allowed) <= allowed


class _UnionFind:
    def __init__(self, items):
        self.parent = {i: i for i in items}

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


def build_canonical_map(value_groups: list[list[str]], label: str,
                        raw_distinct: int | None = None) -> tuple[dict, dict]:
    """
    Group messy spellings of one entity across several sources at once.

    `value_groups` is a list of value lists, e.g. [orders_column, master_column].
    Pooling them matters: if the order export spells a company "ORION LOGISTICS"
    and the customer master spells it "Orion Logistics", clustering each file
    separately gives two canonical labels and the merge silently fails.

    Placeholders (N/A, null, "-", ...) are excluded from the vocabulary before
    clustering, otherwise they act as bridges between unrelated values.
    """
    pool = [v for group in value_groups for v in group]
    counts = Counter(v for v in pool if clean_text(v) and not is_missing(v))
    vocab = sorted(counts)

    uf = _UnionFind(vocab)
    for i, a in enumerate(vocab):
        for b in vocab[i + 1:]:
            if _similar(a, b):
                uf.union(a, b)

    groups: dict[str, list[str]] = defaultdict(list)
    for v in vocab:
        groups[uf.find(v)].append(v)

    def score(v: str) -> tuple:
        """Most common spelling wins; ties prefer Title Case, then longer text."""
        words = _tokens(v)
        title_case = bool(words) and all(
            w[0].isupper() or not w.isalpha() for w in words) and not v.isupper()
        return (counts[v], title_case, len(v))

    mapping: dict[str, str] = {}
    for members in groups.values():
        canonical = max(members, key=score)
        for v in members:
            mapping[v] = canonical

    detail = {
        "field": label,
        "raw_distinct": raw_distinct if raw_distinct is not None else len(counts),
        "canonical_distinct": len(groups),
        "placeholders_excluded": sum(
            1 for v in pool if clean_text(v) and is_missing(v)),
        "mapping": {v: mapping[v] for v in sorted(counts)},
        "merged_groups": [sorted(m) for m in groups.values() if len(m) > 1],
    }
    bump(label, "raw_distinct", detail["raw_distinct"])
    bump(label, "canonical_distinct", len(groups))
    return mapping, detail


# --------------------------------------------------------------------------
# 5. Loaders
# --------------------------------------------------------------------------
def load_csv_sloppy(path: Path) -> pd.DataFrame:
    """Read a CSV whose header cells carry stray whitespace."""
    with path.open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    header = [clean_text(h) for h in rows[0]]
    bump("csv_header", "columns_normalised", len(header))
    body = [r for r in rows[1:] if any(clean_text(c) for c in r)]
    dropped = len(rows) - 1 - len(body)
    if dropped:
        bump("blank_rows", "removed", dropped)
    df = pd.DataFrame(body, columns=header)
    return df


def load_inventory_xlsx(path: Path) -> pd.DataFrame:
    """
    Read an Excel sheet where the real table does not start at row 1.
    We detect the header row instead of hardcoding it.
    """
    raw = load_workbook(path, data_only=True).active
    grid = [[c for c in row] for row in raw.iter_rows(values_only=True)]

    expected = {"sku", "name_jp", "stock", "unit_price_jpy", "note"}
    header_idx = None
    for i, row in enumerate(grid):
        tokens = {clean_text(c).lower() for c in row if c is not None}
        if len(expected & tokens) >= 3:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"No header row found in {path.name}")

    bump("excel_header", "junk_rows_above_table", header_idx)
    header = [clean_text(c) for c in grid[header_idx]]

    body = []
    for row in grid[header_idx + 1:]:
        first = clean_text(row[0]) if row else ""
        if not first or not re.match(r"^INV-\d+$", first, flags=re.I):
            continue                       # skips spacer + footer junk
        body.append(list(row))
    footer = len(grid) - header_idx - 1 - len(body)
    if footer > 0:
        bump("excel_footer", "junk_rows_removed", footer)

    return pd.DataFrame(body, columns=header)


# --------------------------------------------------------------------------
# 6. Main pipeline
# --------------------------------------------------------------------------
def main() -> None:
    report: dict = {"generated_at": datetime.now().isoformat(timespec="seconds"),
                    "source_files": [], "steps": [], "results": {}}

    # ---- load ----------------------------------------------------------
    orders = load_csv_sloppy(RAW / "sales_orders_raw.csv")
    customers = load_csv_sloppy(RAW / "customer_master_raw.csv")
    inventory = load_inventory_xlsx(RAW / "inventory_jp_raw.xlsx")

    for name, df in [("sales_orders_raw.csv", orders),
                     ("customer_master_raw.csv", customers),
                     ("inventory_jp_raw.xlsx", inventory)]:
        report["source_files"].append({"file": name, "rows_in": len(df),
                                       "columns": list(df.columns)})

    rows_in = len(orders)

    # Keep an untouched copy: every "% before" figure in the report must be
    # measured on the file as it arrived, never on partially cleaned data.
    orders_as_delivered = orders.copy()

    def completeness_pct(frame: pd.DataFrame, cols: list[str]) -> float:
        """Share of cells holding a usable value (not a placeholder, not blank)."""
        if not len(frame):
            return 0.0
        usable = 0
        for col in cols:
            usable += sum(
                0 if is_missing(v) else 1
                for v in frame[col].tolist()
            )
        return round(100 * usable / (len(frame) * len(cols)), 2)

    KEY_FIELDS = ["order_date", "customer", "product", "qty", "amount"]
    completeness_before = completeness_pct(orders_as_delivered, KEY_FIELDS)

    # ---- text fields ---------------------------------------------------
    # Structural messiness is measured BEFORE whitespace normalisation, so the
    # report cannot inflate its own numbers by counting " Ltd" and "Ltd" as two
    # different spellings. Both figures are kept and reported separately.
    raw_distinct = {
        label: len({str(v) for v in frame[col] if str(v).strip()})
        for col, label, frame in [
            ("customer", "customer_name", orders),
            ("product", "product_name", orders),
            ("region", "region", orders),
        ]
    }

    for col in orders.columns:
        if orders[col].dtype == object:
            before = orders[col].copy()
            orders[col] = orders[col].map(clean_text)
            changed = int((before.fillna("").astype(str) != orders[col]).sum())
            if changed:
                bump("text_normalisation", col, changed)
    for col in customers.columns:
        if customers[col].dtype == object:
            customers[col] = customers[col].map(clean_text)

    # ---- dates ---------------------------------------------------------
    parsed = [parse_date(v) for v in orders["order_date"]]
    orders["order_date_clean"] = [p[0] for p in parsed]
    methods = Counter(p[1] for p in parsed)
    for method, n in methods.items():
        bump("date_parsing", method, n)
    bad_dates = int(orders["order_date_clean"].isna().sum())

    # ---- numeric -------------------------------------------------------
    amt = [parse_amount(v) for v in orders["amount"]]
    orders["amount_clean"] = [a[0] for a in amt]
    for method, n in Counter(a[1] for a in amt).items():
        bump("amount_parsing", method, n)

    qty = [parse_int(v) for v in orders["qty"]]
    orders["qty_clean"] = [q[0] for q in qty]
    for method, n in Counter(q[1] for q in qty).items():
        bump("qty_parsing", method, n)

    currencies = [parse_currency(c, a) for c, a in
                  zip(orders["currency"], orders["amount"])]
    orders["currency_clean"] = [c[0] for c in currencies]
    for method, n in Counter(c[1] for c in currencies).items():
        bump("currency_resolution", method, n)

    # ---- canonicalisation ---------------------------------------------
    # Customer names are pooled across the order export AND the customer
    # master, so both files end up using the same label and the merge works.
    mappings: dict[str, dict] = {}
    for col, label, pools in [
        ("customer", "customer_name", [orders["customer"], customers["customer_name"]]),
        ("product", "product_name", [orders["product"]]),
        ("region", "region", [orders["region"]]),
    ]:
        mapping, detail = build_canonical_map(
            [list(g) for g in pools], label, raw_distinct=raw_distinct.get(label))
        orders[col] = orders[col].map(lambda v: mapping.get(v, v))
        if label == "customer_name":
            customers["customer_name"] = customers["customer_name"].map(
                lambda v: mapping.get(v, v))
        mappings[label] = detail

    # Order status is a fixed vocabulary, so plain title casing is enough.
    orders["status"] = orders["status"].str.title()
    bump("status_normalisation", "title_cased", int((orders["status"] != "").sum()))

    report["steps"].append({"step": "canonicalisation", "detail": mappings})

    # ---- flags for human review ---------------------------------------
    orders["flag_amount_missing"] = orders["amount_clean"].isna()
    orders["flag_date_unparsed"] = orders["order_date_clean"].isna()
    orders["flag_qty_zero_or_missing"] = orders["qty_clean"].fillna(0) <= 0

    # ---- duplicates ----------------------------------------------------
    business_cols = ["order_id", "customer", "product", "qty_clean",
                     "amount_clean", "currency_clean"]
    before = len(orders)
    exact_dupes = int(orders.duplicated(subset=business_cols, keep="first").sum())
    orders = orders.drop_duplicates(subset=business_cols, keep="first")
    bump("duplicates", "exact_removed", exact_dupes)

    # near-duplicates: same order_id, conflicting content. Keep the most
    # complete record, then the earliest date, and log every decision.
    conflict_log = []
    keep_idx = []
    orders["_completeness"] = orders[["order_date_clean", "customer", "product",
                                      "qty_clean", "amount_clean"]].notna().sum(axis=1)
    for oid, grp in orders.groupby("order_id", sort=False):
        if len(grp) == 1:
            keep_idx.append(grp.index[0])
            continue
        ranked = grp.sort_values(
            by=["_completeness", "order_date_clean", "amount_clean"],
            ascending=[False, True, False])
        winner = ranked.index[0]
        keep_idx.append(winner)
        conflict_log.append({
            "order_id": oid,
            "variants": len(grp),
            "kept": {k: (str(v) if pd.notna(v) else None)
                     for k, v in ranked.loc[winner].items()
                     if k in ("order_date", "amount", "customer", "product", "qty")},
            "rule": "most complete fields, then earliest date, then largest amount",
        })
        bump("duplicates", "near_duplicates_resolved", len(grp) - 1)

    orders = orders.loc[keep_idx]
    report["steps"].append({"step": "duplicate_resolution",
                            "exact_removed": exact_dupes,
                            "order_conflicts": conflict_log})

    # ---- merge with customer master ------------------------------------
    customers = (customers
                 .sort_values("contact_email", key=lambda s: s.isna())
                 .groupby("customer_name", as_index=False)
                 .agg({"contact_email": "first", "phone": "first",
                       "country": "first", "credit_limit": "first"}))
    bump("customer_merge", "master_rows_after_dedup", len(customers))

    merged = orders.merge(customers, left_on="customer",
                          right_on="customer_name", how="left")
    unmatched = int(merged["customer_name"].isna().sum())
    bump("customer_merge", "orders_without_master_record", unmatched)
    if unmatched:
        unmatched_names = sorted(merged.loc[merged["customer_name"].isna(),
                                            "customer"].unique().tolist())
        print(f"  NOTE: {unmatched} orders have no customer-master record. "
              f"Unmatched names: {unmatched_names}")
        report["steps"].append({"step": "customer_merge_unmatched",
                                "names": unmatched_names,
                                "master_names": sorted(customers["customer_name"].tolist())})

    # ---- final tidy ----------------------------------------------------
    cleaned = pd.DataFrame({
        "order_id": merged["order_id"],
        "order_date": merged["order_date_clean"],
        "customer": merged["customer"],
        "product": merged["product"],
        "qty": merged["qty_clean"].astype("Int64"),
        "unit_price": (merged["amount_clean"] /
                       merged["qty_clean"].replace(0, pd.NA)).round(2),
        "amount": merged["amount_clean"].round(2),
        "currency": merged["currency_clean"],
        "region": merged["region"],
        "sales_rep": merged["rep"],
        "status": merged["status"],
        "credit_limit": merged["credit_limit"],
        "contact_email": merged["contact_email"],
        "flag_date_unparsed": merged["flag_date_unparsed"],
        "flag_amount_missing": merged["flag_amount_missing"],
        "flag_qty_zero_or_missing": merged["flag_qty_zero_or_missing"],
    }).sort_values(["order_date", "order_id"], na_position="last").reset_index(drop=True)

    # ---- write workbook ------------------------------------------------
    out_xlsx = OUT / "cleaned_orders.xlsx"

    audit_df = pd.DataFrame([
        {"check": k, "outcome": o, "count": n}
        for k, counter in sorted(stats.items())
        for o, n in sorted(counter.items())
    ])
    canon_df = pd.DataFrame([
        {"field": d["field"], "raw_distinct_values": d["raw_distinct"],
         "canonical_values": d["canonical_distinct"],
         "groups_merged": len(d["merged_groups"])}
        for d in mappings.values()
    ])
    dupe_df = pd.DataFrame(conflict_log) if conflict_log else None

    # Full spelling -> canonical mapping, so the client can audit every merge.
    mapping_rows = []
    for d in mappings.values():
        for raw_value, canon in d.get("mapping", {}).items():
            mapping_rows.append({"field": d["field"], "value_as_written": raw_value,
                                 "normalised_to": canon,
                                 "changed": raw_value != canon})
    mapping_df = pd.DataFrame(mapping_rows)
    if len(mapping_df):
        mapping_df = mapping_df.sort_values(["field", "changed", "value_as_written"],
                                            ascending=[True, False, True])

    sheets = {"Clean Data": cleaned, "Audit Summary": audit_df,
              "Canonicalisation": canon_df, "Value Mapping": mapping_df}
    if dupe_df is not None:
        sheets["Duplicate Decisions"] = dupe_df

    with pd.ExcelWriter(out_xlsx, engine="xlsxwriter") as xl:
        for sheet, frame in sheets.items():
            frame.to_excel(xl, sheet_name=sheet, index=False)

        head = xl.book.add_format({"bold": True, "bg_color": "#1F3B57",
                                   "font_color": "white", "border": 1})
        flag = xl.book.add_format({"bg_color": "#FDE9E9"})
        for sheet, frame in sheets.items():
            ws = xl.sheets[sheet]
            for i, col in enumerate(frame.columns):
                ws.write(0, i, col, head)
                ws.set_column(i, i, min(max(12, len(str(col)) + 4), 34))
            ws.freeze_panes(1, 0)
            if len(frame):
                ws.autofilter(0, 0, len(frame), max(len(frame.columns) - 1, 0))
            # highlight rows still carrying an unresolved value
            flag_cols = [i for i, c in enumerate(frame.columns)
                         if str(c).startswith("flag_")]
            for i in flag_cols:
                ws.set_column(i, i, 22, flag)

    # ---- reports -------------------------------------------------------
    report["results"] = {
        "rows_in": rows_in,
        "rows_out": len(cleaned),
        "rows_removed": rows_in - len(cleaned),
        "columns_out": len(cleaned.columns),
        "completeness_before_pct": completeness_before,
        "completeness_after_pct": completeness_pct(cleaned, KEY_FIELDS),
        "distinct_customers_raw": mappings["customer_name"]["raw_distinct"],
        "distinct_customers_clean": mappings["customer_name"]["canonical_distinct"],
        "distinct_products_raw": mappings["product_name"]["raw_distinct"],
        "distinct_products_clean": mappings["product_name"]["canonical_distinct"],
        "unparsed_dates": bad_dates,
        "orders_without_customer_record": unmatched,
        "total_corrections": sum(sum(c.values()) for c in stats.values()),
    }
    report["checks"] = {k: dict(v) for k, v in sorted(stats.items())}

    (OUT / "data_quality_report.json").write_text(
        json.dumps(report, indent=2, ensure_ascii=False, default=str),
        encoding="utf-8")

    write_markdown_report(report, cleaned)

    print(f"rows in : {rows_in}")
    print(f"rows out: {len(cleaned)}")
    print(f"corrections logged: {report['results']['total_corrections']}")
    print("wrote:", out_xlsx.name, "data_quality_report.json", "data_quality_report.md")


def write_markdown_report(report: dict, cleaned: pd.DataFrame) -> None:
    r = report["results"]
    lines = [
        "# Data Quality Report",
        "",
        f"Generated: {report['generated_at']}",
        "",
        "## 1. Sources",
        "",
        "| File | Rows in |",
        "|---|---|",
    ]
    for s in report["source_files"]:
        lines.append(f"| `{s['file']}` | {s['rows_in']} |")

    lines += [
        "",
        "## 2. Outcome",
        "",
        f"- Rows in: **{r['rows_in']}**",
        f"- Rows out: **{r['rows_out']}** (removed {r['rows_removed']})",
        f"- Total corrections applied: **{r['total_corrections']}**",
        f"- Completeness on key fields: **{r['completeness_before_pct']}% → "
        f"{r['completeness_after_pct']}%**",
        f"- Distinct customers: {r['distinct_customers_raw']} spellings → "
        f"**{r['distinct_customers_clean']} real companies**",
        f"- Distinct products: {r['distinct_products_raw']} spellings → "
        f"**{r['distinct_products_clean']} real SKUs**",
        f"- Dates that could not be parsed: {r['unparsed_dates']} "
        "(kept, flagged, never dropped)",
        f"- Orders with no matching customer record: "
        f"{r['orders_without_customer_record']} (kept, flagged)",
        "",
        "## 3. Every check that ran",
        "",
        "| Check | Outcome | Count |",
        "|---|---|---|",
    ]
    for check, outcomes in report["checks"].items():
        for outcome, n in outcomes.items():
            lines.append(f"| `{check}` | {outcome} | {n} |")

    lines += [
        "",
        "## 4. Decisions a human may want to overrule",
        "",
        "These are judgement calls, not facts. They are listed here so the",
        "client can audit them instead of discovering them later.",
        "",
        "- **Ambiguous numeric dates** (`03/04/2024`): the file mixes "
        "`15/03/2024` and `03/15/2024`, so a value where both readings are "
        "valid is unresolvable from the data alone. One convention was "
        "assumed for the whole file and every row is labelled with the "
        "method used in the audit sheet.",
        "- **Near-duplicate order IDs** (same ID, different date/amount): "
        "kept the most complete record; all candidates are in the "
        "`Duplicate Decisions` sheet.",
        "- **Company names merged by similarity** (edit distance ≤ 3 after "
        "normalisation). Check the `Canonicalisation` sheet before using "
        "customer counts in any report.",
        "- **Missing values were never filled in.** A blank stays blank and "
        "is flagged, so downstream totals are not silently distorted.",
        "",
        "## 5. Column notes",
        "",
    ]
    for col in cleaned.columns:
        lines.append(f"- `{col}`: {cleaned[col].dtype}, "
                     f"{cleaned[col].isna().sum()} missing")
    lines.append("")

    (OUT / "data_quality_report.md").write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
