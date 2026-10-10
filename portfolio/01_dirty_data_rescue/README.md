# Messy Sales Data → Clean, Auditable Dataset

A working example of the job I do most often: **you send me files that came out
of three different systems, and you get back one clean dataset you can actually
pivot, plus a written record of every change I made.**

This repository is self-contained. Anyone can clone it, run three commands, and
get the identical output — nothing here depends on a private file or a library
that is not standard.

---

## 1. The scenario

A distributor exports data from three places and none of them agree with each other:

| File | What it is | What is wrong with it |
|---|---|---|
| `data/raw/sales_orders_raw.csv` | Order export from the sales platform | mixed date formats, currency typed into the amount column, the same company spelled 5 ways, duplicate rows, blank lines |
| `data/raw/customer_master_raw.csv` | Contact list maintained by hand | missing emails and phones, placeholders like `N/A`, the same "company" again |
| `data/raw/inventory_jp_raw.xlsx` | Stock sheet from a supplier | 5 junk rows above the table, 4 junk rows below it, `在庫切れ` and `N/A` mixed into a numeric column |

The sales file alone carries **73 rows**, of which only 62 are real orders.

---

## 2. What was actually wrong

Six classes of problem, all of them found by the script rather than hardcoded:

**1. Invisible characters (fixed 260+ cells).** Leading/trailing spaces, double
spaces, non-breaking spaces (`\u00a0`) and zero-width spaces pasted in from
web pages. These break `VLOOKUP`, `groupby` and de-duplication silently — the
values *look* identical on screen.

**2. Six different date formats in one column.** `2024-03-15`, `15/03/2024`,
`03/15/2024`, `15-Mar-2024`, `2024/03/15`, `Mar 15, 2024`.

`03/04/2024` is genuinely ambiguous — it is either 3 April or 4 March, and the
data cannot tell you which. The script detects the ambiguity, picks **one
convention for the whole file**, and labels every row with the method used so a
human can check it later. This is the kind of thing that quietly shifts a
monthly report by 30 days if nobody catches it.

**3. Currency typed into the amount column, seven ways.** `1,234.56`,
`$1,234.56`, `USD 1,234.56`, `1234.56 USD`, `'1,234.56`, `1234.560`,
`(1,234.56)`. The apostrophe form is an Excel "store as text" artifact;
parentheses mean negative in accounting format. All of it is parsed to a float
and a resolved ISO currency code.

**4. The same company, spelled many ways.** This is the expensive one for
reporting:

| Raw spellings found | Real entities |
|---|---|
| 41 customer name spellings | **8 companies** |
| 26 product descriptions | **5 SKUs** |
| 8 region labels | **5 regions** |

`Aurora Trading`, `AURORA TRADING`, `Aurora Trading  Ltd.`, `Aurora Tradng Ltd`
are one customer. If you report on this file as-is, your customer count is wrong
by 5x and your per-customer totals are split across four rows.

**5. Duplicates, two kinds.** 7 byte-identical rows, plus 4 rows sharing an
`order_id` but disagreeing on date and amount. The second kind is the dangerous
one: a naive `drop_duplicates()` keeps whichever row happened to sort first.

**6. Junk rows in the Excel file.** The real table starts at row 6, not row 1.
Pandas would have read `VENDOR STOCK REPORT` as the header and produced a
dataframe of garbage.

---

## 3. Results

| Measure | Before | After |
|---|---|---|
| Rows delivered | 73 raw | **62 real orders** |
| Completeness, 5 key fields | 97.3% | **100%** |
| Corrections logged | — | **499** |
| Customer entities | 41 spellings | **8** |
| Product entities | 26 spellings | **5** |
| Orders unmatched to a customer record | 38 | **0** |
| Rows deleted without explanation | — | **0** |

Nothing is ever dropped silently. Every removed row is attributed to a named
stage (`blank rows`, `exact duplicates`, `same-ID conflicts`) and the count is
in the audit sheet.

---

## 4. What you get

| Output | Purpose |
|---|---|
| `output/cleaned_orders_charts.xlsx` | The deliverable. A **Dashboard** sheet with four live Excel charts, the clean data, the audit trail, and every mapping I applied. |
| `output/data_quality_report.md` | Readable summary, including the decisions a human may want to overrule. |
| `output/data_quality_report.json` | The same audit trail, machine-readable, for a pipeline to consume. |
| `scripts/01_generate_dirty_data.py` | Rebuilds the messy inputs. Exists so this is reproducible, not a screenshot. |
| `scripts/02_clean_and_report.py` | The actual cleaning pipeline. Re-runnable on next month's export. |
| `scripts/03_make_charts.py` | Builds the summary workbook. |
| `docs/engineering_notes.md` | The one bug worth reading about: how a fuzzy-match threshold silently merged all five regions into one, and how it was caught. |

The workbook contains one audit sheet per decision type:

- **Audit Summary** — every check that ran and how many rows it touched
- **Value Mapping** — `value_as_written` → `normalised_to`, for every change
- **Canonicalisation** — which spellings were merged into which entity
- **Duplicate Decisions** — each conflicting `order_id`, the variants, and which won
- **Clean Data** — the result, with `flag_*` columns marking anything unresolved

---

## 5. How it works

No ML, no paid API, no magic. Four deliberate decisions:

**Dates are parsed by explicit pattern matching, and ambiguity is reported, not
hidden.** Every value gets a `method` label (`iso`, `day-first`,
`numeric-assumed-day-first`, …) that lands in the audit trail.

**Entity resolution uses two rules with a length-scaled threshold.**

1. *Token overlap ≥ 50%* — `Aurora Trading` matches `Aurora Trading Ltd`,
   because they share the words that matter.
2. *Edit distance on the normalised string* — but **only 1 edit for short
   values, 2 for longer ones.**

That scaling is the whole trick. Short strings are always "close" to each
other: `editdistance("east","west")` is 2, and `editdistance("na","east")` is 3.
I found this the hard way while building this — an earlier version with a flat
threshold of 2 merged **North, South, East and West into a single region**,
because `n/a` acted as a bridge between unrelated values. On a real client file
that silently destroys any report grouped by region. The threshold is now
length-scaled, placeholders are removed from the vocabulary before clustering,
and `05_limitations` below is honest about what is still guesswork.

**Missing values are never filled in.** A blank stays blank and is flagged
(`flag_amount_missing`, `flag_qty_zero_or_missing`). Imputing a value would
make downstream totals look complete while being wrong.

**The pipeline is idempotent and re-runnable.** Point it at next month's export
and the same rules apply; new spellings surface in the `Value Mapping` sheet
instead of being silently accepted.

---

## 6. Limitations — read this before reusing the pattern

I would rather state these than have you find them yourself:

- **British/American spelling merges are a judgement call.** `Blue Harbour Foods`
  and `Blue Harbor Foods` were merged by token overlap. It is almost certainly
  right here, but at 10,000 customers you need an explicit alias table with a
  human approval step, not an edit-distance rule.
- **Ambiguous numeric dates are resolved by convention, not by fact.** If your
  file contains `03/04/2024`, one reading is assumed for the entire file. Confirm
  the convention before trusting date-based totals.
- **Same-ID conflicts are resolved by rule** (most complete record → earliest
  date → largest amount). All candidates are exported so a human can overrule
  any of them.
- **Region placeholders (`N/A`, `null`, `-`) are excluded from region counting.**
  They are 7 of 73 rows. Where those orders belong is a business question, not a
  parsing question.
- **The supplier stock sheet is parsed for structure but not merged into the
  result.** It demonstrates multi-language and junk-header handling; joining it
  to orders would need a SKU crosswalk the scenario does not provide.

---

## 7. Running it

```bash
pip install pandas openpyxl xlsxwriter
python scripts/01_generate_dirty_data.py   # rebuild the messy inputs
python scripts/02_clean_and_report.py      # clean + audit
python scripts/03_make_charts.py           # build the summary workbook
```

Tested with Python 3.11, pandas 3.0, openpyxl 3.x, xlsxwriter 3.x.

---

## 8. Using this on your own data

The cleaning script is structured so the reusable parts are separate from the
scenario-specific parts. To adapt it:

1. Replace the loaders in section 5 with your file layout. The Excel reader
   already **detects** the header row instead of assuming row 1.
2. Leave sections 1–4 alone: text normalisation, date parsing, amount parsing
   and entity resolution are generic.
3. Add your own checks to `stats` with `bump(...)` and they appear in the audit
   sheet automatically.
4. Set `KEY_FIELDS` to the columns your reports depend on; that drives the
   completeness measure.

If you want this run against your files, the first step is the same as it was
here: send me one sample and I will tell you what is wrong with it before you
pay for anything.
