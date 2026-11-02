# MRT-3 ridership data, 2023 to 2025

MRT-3 station entries and exits, by hour and by day, converted from three Excel workbooks to CSV.

**The data was acquired through the Department of Transportation (DOTr) Freedom of Information (FOI) process.**

## Choose a file

| Year | Original workbook | Hourly CSV | Daily CSV | Days | Hourly rows |
| --- | --- | --- | --- | ---: | ---: |
| 2023 | [Excel](Daily%20Hourly%202023.xlsx) | [Hourly](csv/mrt3_hourly_2023.csv) | [Daily](csv/mrt3_daily_2023.csv) | 365 | 8,760 |
| 2024 | [Excel](Daily%20Hourly%202024.xlsx) | [Hourly](csv/mrt3_hourly_2024.csv) | [Daily](csv/mrt3_daily_2024.csv) | 366 | 8,784 |
| 2025 | [Excel](Daily%20Hourly%202025.xlsx) | [Hourly](csv/mrt3_hourly_2025.csv) | [Daily](csv/mrt3_daily_2025.csv) | 365 | 8,760 |
| Total | | | | **1,096** | **26,304** |

The table shows counts after filtering `date_in_sheet_month` to `true`. Apply this filter before counting days, joining on dates, or calculating averages: every monthly worksheet has 31 day blocks, so shorter months spill into the next month. The 20 extra blocks contain only zeros or blanks and are retained for traceability.

Use **hourly files** for time-of-day patterns. Use **daily files** for the reported daily totals, copied from each source `TOTAL` row. Daily totals are not recalculated from hourly values, and some disagree with them.

Before filtering, each yearly file has 8,928 hourly rows or 372 daily rows. The reporting period is January 1, 2023 through December 31, 2025. Every date has a daily row and 24 hourly slots, although some counts are blank or zero.

Other files:

- [Raw CSVs](csv/raw): all 36 worksheets, saved as `csv/raw/YYYY/MM.csv`. These have no added header row: CSV row 1 is Excel row 1, and the first field is column A. Excel date serials and supplementary tables are preserved.
- [Quality summary](csv/quality_summary.csv): coverage, blank and fractional counts, and issue counts for each worksheet.
- [Issue details](csv/data_quality_issues.csv): individual flags, with source cell locations and comparison values.
- [Source manifest](csv/source_manifest.json): workbook hashes, sheet names, and conversion conventions.

## What the numbers mean

Each row contains **Entry** and **Exit** values for 13 stations, plus the workbook's system totals, `total_entry` and `total_exit`. Keep entries and exits separate; adding them does not give a count of unique passengers. These are aggregate reports, with no passenger records, trip pairings, or train direction.

The system totals already cover the stations. Do not add them to station counts again. Blanks stay blank, recorded zeros stay zero, and decimals are kept exactly as stored. The source does not explain the fractional values or provide a detailed counting method.

## Fields

The hourly files have 42 columns; the daily files have 36. Both include these fields:

| Field | Meaning |
| --- | --- |
| `report_date` | Date on the daily report heading, in `YYYY-MM-DD` form. |
| `source_year`, `source_month` | Workbook year and worksheet month; month is 1 to 12. |
| `date_in_sheet_month` | `true` when the report date belongs to that worksheet's month; otherwise `false`. |
| `source_workbook`, `source_sheet` | Original filename and sheet name. Keep sheet names as text. |
| `source_row`, `source_time_column` | Excel row number and column letter for the interval label or `TOTAL`. |
| `<station>_entry`, `<station>_exit` | The station's reported values; a number or an empty field. |
| `total_entry`, `total_exit` | Reported system totals for the hour or day. |

Hourly files also include:

| Field | Meaning |
| --- | --- |
| `hour_sequence` | Slot 1 to 24 in the source report's order. |
| `time_label` | Original label, such as `03:00 - 03:59`. |
| `hour_start`, `hour_end` | Times parsed from that label, such as `03:00` and `03:59`. |
| `expected_hour_start` | Time inferred from the slot's position in the usual daily template. |
| `hour_label_matches_sequence` | `true` if the label agrees with its position; otherwise `false`. |

Boolean fields use lowercase `true` and `false`. The 28 count columns allow decimals and blanks. For example, `north_ave_entry` and `north_ave_exit` are North Ave's two columns. Prefixes follow the workbook labels:

| Station label | Prefix | Station label | Prefix |
| --- | --- | --- | --- |
| North Ave | `north_ave` | Quezon Ave | `quezon_ave` |
| GMA Kamuning | `gma_kamuning` | Cubao | `cubao` |
| Santolan | `santolan` | Ortigas | `ortigas` |
| Shaw Blvd | `shaw_blvd` | Boni Ave | `boni_ave` |
| Guadalupe | `guadalupe` | Buendia | `buendia` |
| Ayala Ave | `ayala_ave` | Magallanes | `magallanes` |
| Taft | `taft` | System total | `total` |

Every analysis row can be traced using `(source_workbook, source_sheet, source_row)`. After the month filter, use `(report_date, hour_sequence)` as the hourly key and `report_date` as the daily key. Hour labels alone are not unique.

## Dates, blanks, and known issues

The daily template starts at 03:00 and ends at 02:59. All 24 slots keep the date on the report heading, including the after-midnight hours. No calendar-day shift or timezone conversion is applied. If you build timestamps, document how you handle those final three slots.

**A blank is not a recorded zero.** Some rows have blank station counts even when the system total is zero. All-zero days are retained too; the files do not explain the reason for them. State which days you include when calculating averages.

| Source finding | 2023 | 2024 | 2025 |
| --- | ---: | ---: | ---: |
| Extra month-end day blocks | 7 | 6 | 7 |
| Mislabeled hourly slots | 20 | 24 | 24 |
| Daily count cells differing from summed hours | 0 | 0 | 33 |
| System total cells differing from summed stations | 2 | 2 | 12 |
| Excel error cells, outside the main count tables | 168 | 0 | 0 |
| Supplementary tables | 6 | 0 | 0 |

All 68 hour-label mismatches say `03:00 - 03:59` in the final slot, where `02:00 - 02:59` would be expected. The original labels are preserved. `expected_hour_start` describes the template position; it is not an official correction.

The six supplementary tables and 168 `#DIV/0!` cells are in [July 2023's raw export](csv/raw/2023/07.csv), outside the primary report. They are preserved there and excluded from the yearly analysis tables.

The quality summary counts blank and fractional cells across both hourly and daily tables, including system totals and extra blocks. Sum checks omit blank cells and flag differences greater than `0.000001`. This is a consistency check, not a decision to treat missing observations as zero. No source discrepancies have been silently corrected; the [issue details](csv/data_quality_issues.csv) show their locations.

## Load the data

In Excel, import a yearly CSV as UTF-8 with a comma delimiter and use its first row as headers. Keep `source_sheet` as text and filter `date_in_sheet_month` to `true`. Raw exports should be imported without headers.

This example needs pandas and runs from the repository root:

```python
from pathlib import Path
import pandas as pd


def load_reports(kind):
    frames = [
        pd.read_csv(path, dtype={
            "source_sheet": "string",
            "date_in_sheet_month": "string",
            "hour_label_matches_sequence": "string",
        }, parse_dates=["report_date"])
        for path in sorted(Path("csv").glob(f"mrt3_{kind}_*.csv"))
    ]
    data = pd.concat(frames, ignore_index=True)
    return data.loc[data["date_in_sheet_month"].eq("true")].copy()


daily = load_reports("daily")
hourly = load_reports("hourly")
monthly = (
    daily.set_index("report_date")[["total_entry", "total_exit"]]
    .resample("MS").sum(min_count=1)
)
flagged_hours = hourly.loc[hourly["hour_label_matches_sequence"].eq("false")]
print(monthly)
```

For exact numeric text, read with `dtype=str, keep_default_na=False` or use Python's `csv` module. Spreadsheet and pandas number handling can change displayed precision. Use `decimal.Decimal` when you need exact decimal arithmetic.

## Rebuild or check the exports

These commands need Python 3.10 or newer and no extra packages:

```console
python scripts/convert_to_csv.py          # Create or overwrite the 45 export files
python scripts/convert_to_csv.py --check  # Compare all exports with a fresh conversion
```

The exports use UTF-8 without a BOM, comma delimiters, and LF line endings. CSV keeps cell values, including stored numeric precision, but cannot preserve workbook formatting or merged ranges. The original Excel files remain unchanged.

## Attribution

> Source: Department of Transportation (DOTr), MRT-3 hourly ridership reports, 2023 to 2025. Acquired through DOTr's Freedom of Information (FOI) process and converted to CSV in this repository.

When sharing analysis, mention whether you used reported daily totals or hourly values and explain any changes to blanks, dates, or flagged values. No reuse license or release conditions were included with the supplied files.
