# Excel VBA Monthly Reporting Automation

A portfolio project by **Mahesh M**: a VBA macro that turns a raw monthly transactions CSV into a formatted management summary in one run - import, clean, aggregate, format.

> **Demo project.** All sample data is synthetic and randomly generated (fixed seed). It does not come from any real company or employer, and the numbers are not real-world metrics.

## What the macro does

`vba/MonthlyReporting.bas`, entry point `BuildMonthlyReport`:

1. **Import** - asks for the CSV and loads it values-only into a fresh `Raw` sheet.
2. **Clean** - dedupes on `txn_id` (the sample file has 5 injected duplicates), trims padded labels, normalises mixed date formats (`yyyy-mm-dd` and `m/d/yyyy`) to ISO.
3. **Summarise** - builds the `Summary` sheet: clean row count, grand total, and totals by category and by region.
4. **Format** - styled title and header band, currency formats, autofit, freeze panes, and a highlight on the top-earning category.

Late binding throughout (`CreateObject("Scripting.Dictionary")`) - no extra references to enable.

## Run it

1. Open a blank workbook, save as `.xlsm` (macro-enabled).
2. `Alt+F11` -> File -> Import File -> select `vba/MonthlyReporting.bas`.
3. Run `BuildMonthlyReport` and pick `data/transactions_2026-01.csv`.

Requires desktop Excel (Windows or Mac). No add-ins.

## Sample data and verification

`data/transactions_2026-01.csv` is 216 synthetic January transactions (6 columns: txn_id, date, category, region, rep, amount) with defects injected on purpose: 5 duplicate rows, 4 padded category labels, 3 US-format dates.

`data/expected_summary.csv` is the ground truth, computed independently in Python (`scripts/generate_sample_data.py`) - after running the macro, the Summary sheet should match it exactly:

| Dimension | Member | Total |
|---|---|---|
| Category | Training | $141,715.65 |
| Category | Support | $140,613.45 |
| Category | Services | $132,399.31 |
| Category | Hardware | $125,623.70 |
| Category | Software | $104,778.05 |
| Region | East | $176,521.61 |
| Region | North | $161,791.00 |
| Region | West | $155,496.02 |
| Region | South | $151,321.53 |
| **Grand total** | | **$645,130.16** |

Clean rows: **211** (216 raw minus 5 duplicates).

## Layout

```
vba/MonthlyReporting.bas           the macro (import, clean, summarise, format)
scripts/generate_sample_data.py    regenerates the sample CSV + expected ground truth
data/transactions_2026-01.csv      sample raw input
data/expected_summary.csv          independently computed expected output
```

## Possible extensions

Multi-month consolidation across files, a pivot-table output variant, automatic chart sheet, and email-as-attachment distribution.

---
Author: Mahesh M. Released under the MIT License.
