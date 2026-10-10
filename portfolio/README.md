# Data Cleaning & Automation Portfolio

Third-year **Big Data Management & Application** student. I turn messy
CSV/Excel exports and web data into clean, analysis-ready datasets — and I
leave a written record of every change so you can check my work.

**Working languages:** Python (pandas, openpyxl, xlsxwriter) · SQL · Excel ·
data visualisation. English and Chinese; comfortable with Japanese data files.

---

## Why these samples exist

Most new freelancers on Upwork have a profile and no evidence. I would rather
show you the actual work, with the actual numbers, including the parts that are
still guesswork.

Every project below is **fully reproducible**: run the scripts and you get the
same output. Nothing here is a screenshot, and nothing depends on a private file
or a paid service.

---

## Projects

| # | Project | What it demonstrates | Scale |
|---|---|---|---|
| 01 | [Messy Sales Data → Clean, Auditable Dataset](01_dirty_data_rescue/) | Data cleaning, entity resolution, duplicate handling, audit reporting | 73 raw rows → 62 clean orders, 499 corrections logged, 41 customer spellings → 8 real companies |

*More projects are being prepared: a web-scraping pipeline and an event-study
analysis of stock price reactions to natural disasters.*

---

## The short version of what I do

**Data cleaning that survives scrutiny.** Anyone can run `dropna()` and
`drop_duplicates()`. The hard part is knowing which rows *should* be dropped,
which "different" values are the same thing, and being able to explain every
decision afterwards. That is what project 01 is about.

Three rules I hold to:

1. **Never delete a row silently.** Every removed row is attributed to a named
   stage and counted.
2. **Never invent a value.** A blank stays blank and is flagged. Filling it in
   makes a report look complete while being wrong.
3. **Report the ambiguities.** If `03/04/2024` could be two different dates, I
   say so and label which reading I assumed — rather than picking one and
   hoping nobody notices.

---

## Technical skills

| Area | Tools |
|---|---|
| Data cleaning & transformation | pandas, NumPy, openpyxl, xlsxwriter |
| Web data collection | Python requests, JSON/API parsing, pagination, rate limiting |
| Databases | SQL (joins, aggregation, window functions), MySQL, SQLite |
| Analysis & modelling | Statistical testing, regression, event-study methodology (AR/CAR), market model |
| Reporting | Excel (pivot-ready output, native charts), Markdown/Word reports, JSON audit trails |
| Languages | Python, SQL, English (professional), Chinese (native), some Japanese |

---

## How a job with me usually runs

1. **You send one sample file (or 10 rows).** I tell you what is wrong with it
   and what cleaning it will involve — **before you pay anything.**
2. **You get a fixed price and a delivery date**, plus a list of exactly what is
   and is not included.
3. **You get the cleaned file, the script that produced it, and a report** of
   every change. If you need it again next month, the script re-runs on the new
   export.
4. **Anything ambiguous is flagged, not guessed at.** Where a decision needs
   your business knowledge, I ask instead of assuming.

---

## Repository layout

```
portfolio/
├── README.md                     <- you are here
├── .gitignore
└── 01_dirty_data_rescue/
    ├── README.md                 <- full case study: method, results, limitations
    ├── data/raw/                 <- the messy source files (deliberately bad)
    ├── scripts/                  <- 01 generate, 02 clean, 03 chart
    ├── output/                   <- clean data, audit reports, summary workbook
    └── docs/
        └── engineering_notes.md  <- the bug worth reading about
```

## Running any project

```bash
pip install pandas openpyxl xlsxwriter
cd 01_dirty_data_rescue
python scripts/01_generate_dirty_data.py
python scripts/02_clean_and_report.py
python scripts/03_make_charts.py
```

Tested with Python 3.11, pandas 3.0, openpyxl 3.x, xlsxwriter 3.x.

---

## Contact

Available for fixed-price data cleaning, scraping and reporting jobs.
Send me a sample and I will tell you what is wrong with it.
