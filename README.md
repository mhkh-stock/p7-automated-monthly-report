# P7 — Automated Monthly Report

**Client pain:** every month, somebody rebuilds the same report by hand —
copying sales into a sheet, re-totaling, re-charting, re-writing the
summary. It takes a day, and the month-on-month number is wrong whenever
anyone pastes a row in the wrong place.
**This delivery:** paste the month's lines in, and the report builds
itself. **Status: SAMPLE delivery**
**Tools:** Excel (live formulas + native charts) · Python (openpyxl
builder) · Google Apps Script (Sheets version of the same automation)
**Deliverable:** `outputs/monthly_report_SAMPLE.xlsx` (Data + Summary +
Monthly Trend, 560 lines, Jul–Sep 2026)
**Automation:** `apps_script/monthly_report_automation.gs` — the same
report inside Google Sheets (Summary only), with `setupMonthlyTrigger()`
to switch on the monthly 1st-of-month run.

> ⚠️ **What this is and is not.** "CityMart" is a fictional shop and
> all 560 lines are generated SAMPLE data (`src/generate_sample_data.py`,
> seed 42 — re-running reproduces the identical file). What is real is
> the automation: every number on Summary and Monthly Trend is a live
> formula over the Data sheet, not a typed value, so a new month's lines
> change the report without anyone touching a total or a chart.

## Business question it answers

> "Give me last month's numbers, versus the month before, without
> anyone spending a day building the report — and make next month's
> report build itself the same way."

## What is inside

| File | What it does |
|---|---|
| `outputs/monthly_report_SAMPLE.xlsx` | The deliverable: Data (560 lines, Excel Table) + Summary (KPIs + category chart, all formulas) + Monthly Trend (trend table + line chart) |
| `data/raw/sales_SAMPLE.csv` | The 560 SAMPLE lines the workbook is built from |
| `src/generate_sample_data.py` | Generates the sample CSV (deterministic, seed 42) |
| `src/build_report.py` | Builds the workbook: Data table, live-formula Summary/Trend, native charts |
| `src/checks.py` | Data-quality checks as code — 16 checks, all passing |
| `apps_script/monthly_report_automation.gs` | Google Sheets version: "Monthly Report" menu + `buildMonthlyReport()` (rebuilds Summary) + `setupMonthlyTrigger()` (run once for the monthly run) |

## 3 key findings (from this sample)

1. **Formulas, not values, are the automation.** Summary's revenue,
   growth, orders and average order value are `SUMPRODUCT` / ratio
   formulas over Data — the report cannot go stale when new lines
   arrive, because there is nothing typed to go stale. *So what:* a
   client who pastes October's lines gets October's report for free.
2. **The month-on-month story is growth, not a single month.** Sep
   2026 revenue (PKR 2,862,500) is +15.8% on Aug (PKR 2,471,500),
   which itself was +24.4% on Jul (PKR 1,987,250) — three rising
   months, not one good month. *So what:* the Monthly Trend sheet
   exists so a single strong month is never mistaken for a trend.
3. **One category carries the month.** Furniture is PKR 1,291,500 of
   Sep's PKR 2,862,500 (45%) — Office Chair alone is PKR 737,100.
   *So what:* the category chart is on the Summary sheet, not buried,
   because the "what drove the month" answer is the first thing a
   shop owner asks.

## Verified numbers

| KPI | Value |
|---|---|
| Lines in the report | 560 (Jul 170, Aug 185, Sep 205) |
| Grand total, Jul–Sep 2026 | PKR 7,321,250 |
| Report month (Sep 2026) revenue | PKR 2,862,500 |
| Previous month (Aug 2026) revenue | PKR 2,471,500 |
| Month-on-month growth | +15.82% |
| Sep orders / average order value | 205 / PKR 13,963 |
| Sep top category / top product | Furniture PKR 1,291,500 / Office Chair PKR 737,100 |
| Checks | 16 passing (`python3 src/checks.py`) |

Verification: `python3 src/checks.py` (16 checks pass) plus an
independent reopen of the workbook — 3 sheets, 561 Data rows
including the header, live formulas in Summary B5/B7, one chart on
Summary and one on Monthly Trend.

## How to run it

1. Open `outputs/monthly_report_SAMPLE.xlsx` — that is the deliverable.
2. To regenerate the sample data: `python3 src/generate_sample_data.py`.
3. To rebuild the workbook: `python3 src/build_report.py`.
4. To check the data: `python3 src/checks.py`.
5. To reuse it for a real client: replace `data/raw/sales_SAMPLE.csv`
   with the client's lines (same 7 columns) and re-run step 3 — or
   paste the Apps Script into their Google Sheet and run
   `buildMonthlyReport()` from the "Monthly Report" menu.
6. Next month: paste the new lines into Data (extend the SalesData
   table) and re-run step 3 — the builder re-derives the report month
   and rebuilds Summary + Monthly Trend from Data. Pasting lines alone
   does not move Summary: its formulas are built for the months in the
   file at build time.

## Limits

- SAMPLE data only — 3 months (Jul–Sep 2026), one fictional shop;
  replace the CSV for a real client.
- Excel: pasting a new month's lines into Data does not move Summary on
  its own — re-run `src/build_report.py` (see step 6 above).
- Google Sheets version rebuilds Summary only (KPIs + category chart);
  it does not build the Monthly Trend sheet. Run
  `setupMonthlyTrigger()` once to switch on the 1st-of-month run.
- No POS/shop integration and no data cleaning — a messy source file
  needs the P2 cleaning service first (see `PACKAGE-NOTE.md`).
- Formulas assume the 7-column layout (`order_id, date, product,
  category, qty, unit_price_pkr, amount_pkr`) and ISO dates
  (`YYYY-MM-DD`).

## Tools

- Python 3 · openpyxl 3.1.2 (pinned in `requirements.txt`) · Google Apps Script (Sheets)
- Excel / Google Sheets (the client's tool, not a new tool to learn)

## Data source note

**SAMPLE data, generated 3 Nov 2026.** No real business's sales are
involved. The automation, the formulas and the workbook structure
are the product; the numbers exist to prove them.
