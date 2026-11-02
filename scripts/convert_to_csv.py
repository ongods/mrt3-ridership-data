"""Export the supplied MRT-3 workbooks using only the Python standard library.

Run from any directory: python scripts/convert_to_csv.py [--check]
This is a converter for these report layouts, not a general Excel renderer.
"""

from __future__ import annotations

import argparse
import calendar
import csv
from collections import Counter
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from hashlib import sha256
from io import StringIO
import json
from pathlib import Path
import posixpath
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
STATIONS = [
    ("North Ave", "north_ave"), ("Quezon Ave", "quezon_ave"),
    ("GMA Kamuning", "gma_kamuning"), ("Cubao", "cubao"),
    ("Santolan", "santolan"), ("Ortigas", "ortigas"),
    ("Shaw Blvd", "shaw_blvd"), ("Boni Ave", "boni_ave"),
    ("Guadalupe", "guadalupe"), ("Buendia", "buendia"),
    ("Ayala Ave", "ayala_ave"), ("Magallanes", "magallanes"),
    ("Taft", "taft"), ("Total", "total"),
]
COUNTS = [f"{slug}_{flow}" for _, slug in STATIONS for flow in ("entry", "exit")]
META = ["report_date", "source_year", "source_month", "date_in_sheet_month",
        "source_workbook", "source_sheet", "source_row", "source_time_column"]
HOURLY = META + ["hour_sequence", "time_label", "hour_start", "hour_end",
                 "expected_hour_start", "hour_label_matches_sequence"] + COUNTS
DAILY = META + COUNTS
ISSUES = ["issue", "source_workbook", "source_sheet", "report_date",
          "source_row", "source_cell", "field", "source_value", "comparison_value", "detail"]
TOLERANCE = Decimal("0.000001")
TIME = re.compile(r"^(\d{2}):00\s*-\s*(\d{2}):59$")


def column_number(label: str) -> int:
    value = 0
    for char in label:
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def column_label(number: int) -> str:
    label = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        label = chr(65 + remainder) + label
    return label


def number(value: str) -> Decimal | None:
    try:
        result = Decimal(value)
        return result if result.is_finite() else None
    except InvalidOperation:
        return None


def read_workbook(path: Path):
    """Read stored values, retaining Excel serials and number text verbatim."""
    with ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ["".join(t.text or "" for t in item.findall(".//s:t", NS))
                       for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        properties = workbook.find("s:workbookPr", NS)
        epoch = date(1904, 1, 1) if properties is not None and properties.get("date1904") in ("1", "true") else date(1899, 12, 30)
        relationships = {r.get("Id"): r.get("Target") for r in ET.fromstring(
            archive.read("xl/_rels/workbook.xml.rels"))}
        for sheet in workbook.find("s:sheets", NS):
            target = relationships[sheet.get(RID)]
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            tree = ET.fromstring(archive.read(target))
            cells, types, formulas = {}, {}, {}
            max_row = max_col = 0
            for row in tree.findall("s:sheetData/s:row", NS):
                max_row = max(max_row, int(row.get("r")))
                for cell in row.findall("s:c", NS):
                    address = cell.get("r")
                    match = re.fullmatch(r"([A-Z]+)(\d+)", address)
                    col, row_number = column_number(match[1]), int(match[2])
                    max_col = max(max_col, col)
                    max_row = max(max_row, row_number)
                    kind = cell.get("t", "n")
                    value_element = cell.find("s:v", NS)
                    value = value_element.text or "" if value_element is not None else ""
                    if kind == "s":
                        value = strings[int(value)]
                    elif kind == "inlineStr":
                        value = "".join(t.text or "" for t in cell.findall("s:is//s:t", NS))
                    if cell.find("s:f", NS) is not None:
                        formulas[address] = cell.find("s:f", NS).text or ""
                        if value_element is None:
                            raise ValueError(f"Formula without stored result: {path.name}/{sheet.get('name')}/{address}")
                    if value != "":
                        cells[(row_number, col)] = value
                        types[(row_number, col)] = kind
            yield sheet.get("name"), cells, types, formulas, max_row, max_col, epoch


def csv_text(headers, rows):
    stream = StringIO(newline="")
    writer = csv.writer(stream, lineterminator="\n")
    if headers is not None:
        writer.writerow(headers)
    writer.writerows(rows)
    return stream.getvalue()


def convert(check_only=False):
    outputs, issues, summaries, manifest = {}, [], [], []
    for year in (2023, 2024, 2025):
        path = ROOT / f"Daily Hourly {year}.xlsx"
        hourly, daily = [], []
        sources = list(read_workbook(path))
        if len(sources) != 12:
            raise ValueError(f"Expected 12 monthly sheets in {path.name}")
        for month, (sheet, cells, types, formulas, height, width, epoch) in enumerate(sources, 1):
            raw_path = f"csv/raw/{year}/{month:02d}.csv"
            outputs[raw_path] = csv_text(None, (
                [cells.get((row, col), "") for col in range(1, width + 1)]
                for row in range(1, height + 1)))
            stats = Counter()
            sheet_hourly, sheet_daily = [], []
            consumed = set()

            def issue(kind, report_date="", row="", col=None, field="", value="", comparison="", detail=""):
                stats[kind] += 1
                issues.append([kind, path.name, sheet, report_date, row,
                               f"{column_label(col)}{row}" if col is not None else "",
                               field, value, comparison, detail])

            headers = sorted((row, col) for (row, col), value in cells.items() if value == "TIME" and col <= 2)
            for (row, col), value in sorted(cells.items()):
                if value == "TIME" and col > 2:
                    issue("supplementary_table", row=row, col=col, value=value,
                          detail="Additional table to the right of the primary report; preserved in raw CSV only.")
            if len(headers) != 31:
                raise ValueError(f"Expected 31 report blocks in {path.name}/{sheet}")
            for header_row, time_col in headers:
                serial = cells.get((header_row - 2, time_col), "")
                serial_number = number(serial)
                if serial_number is None or serial_number != int(serial_number):
                    raise ValueError(f"Missing date at {path.name}/{sheet}/{header_row - 2}")
                report_date = epoch + timedelta(days=int(serial_number))
                in_month = report_date.year == year and report_date.month == month
                if not in_month:
                    issue("outside_sheet_month", report_date.isoformat(), header_row - 2, time_col,
                          "report_date", report_date.isoformat(), f"{year}-{month:02d}",
                          "Extra report block retained; exclude date_in_sheet_month=false for calendar analysis.")
                columns = []
                for index, (station, slug) in enumerate(STATIONS):
                    col = time_col + 1 + index * 2
                    if cells.get((header_row, col)) != station:
                        raise ValueError(f"Unexpected station header: {path.name}/{sheet}/{header_row}/{col}")
                    for offset, flow in enumerate(("entry", "exit")):
                        if cells.get((header_row + 1, col + offset), "").lower() != flow:
                            raise ValueError(f"Unexpected entry/exit header: {path.name}/{sheet}/{header_row + 1}")
                        columns.append(col + offset)
                block_counts = []
                for sequence in range(1, 26):
                    row = header_row + 1 + sequence
                    label = cells.get((row, time_col), "")
                    meta = [report_date.isoformat(), year, month, str(in_month).lower(), path.name,
                            sheet, row, column_label(time_col)]
                    values = [cells.get((row, col), "") for col in columns]
                    for field, col, value in zip(COUNTS, columns, values):
                        consumed.add((row, col))
                        if not value:
                            stats["blank_count_cells"] += 1
                        else:
                            numeric = number(value)
                            if numeric is None:
                                issue("nonnumeric_count", report_date.isoformat(), row, col, field, value)
                            else:
                                stats["numeric_count_cells"] += 1
                                if numeric != numeric.to_integral_value():
                                    stats["fractional_count_cells"] += 1
                                if numeric < 0:
                                    issue("negative_count", report_date.isoformat(), row, col, field, value)
                    if sequence <= 24:
                        match = TIME.fullmatch(label)
                        if match is None or match[1] != match[2] or int(match[1]) > 23:
                            raise ValueError(f"Unexpected hour label: {path.name}/{sheet}/{row}: {label!r}")
                        expected = f"{(sequence + 2) % 24:02d}:00"
                        start, end = match[1] + ":00", match[2] + ":59"
                        matches = start == expected
                        if not matches:
                            issue("hour_label_sequence_mismatch", report_date.isoformat(), row, time_col,
                                  "time_label", label, expected,
                                  "Label retained; expected hour is inferred only from position in the 03:00-to-02:59 template.")
                        # Later report layouts repeat the report date in column A.
                        if time_col > 1:
                            row_serial = cells.get((row, time_col - 1), "")
                            if row_serial and row_serial != serial:
                                issue("row_date_mismatch", report_date.isoformat(), row, time_col - 1,
                                      "row_date_serial", row_serial, serial, "Normalized report_date uses the block heading.")
                        record = meta + [sequence, label, start, end, expected, str(matches).lower()] + values
                        hourly.append(record)
                        sheet_hourly.append(record)
                        block_counts.append(values)
                    else:
                        if label != "TOTAL":
                            raise ValueError(f"Expected TOTAL: {path.name}/{sheet}/{row}")
                        daily.append(meta + values)
                        sheet_daily.append(meta + values)
                        # Compare stored daily totals to numeric hourly cells.
                        # Blanks contribute no numeric amount; this does not assert that blank = zero.
                        for index, (field, col, value) in enumerate(zip(COUNTS, columns, values)):
                            total = number(value)
                            parts = [number(hour[index]) for hour in block_counts]
                            if total is None or any(v is None and block_counts[i][index] != "" for i, v in enumerate(parts)):
                                continue
                            computed = sum((v for v in parts if v is not None), Decimal(0))
                            if abs(total - computed) > TOLERANCE:
                                issue("daily_hourly_sum_mismatch", report_date.isoformat(), row, col,
                                      field, value, str(computed),
                                      "Stored daily total differs from sum of numeric hourly values; blanks omitted.")
                    # Compare each stored system total to the 13 station values.
                    for offset, flow in enumerate(("entry", "exit")):
                        total = number(values[26 + offset])
                        station_values = [values[i] for i in range(offset, 26, 2)]
                        numeric_values = [number(v) for v in station_values]
                        if total is None or any(v is None and raw != "" for v, raw in zip(numeric_values, station_values)):
                            continue
                        computed = sum((v for v in numeric_values if v is not None), Decimal(0))
                        if abs(total - computed) > TOLERANCE:
                            issue("system_station_sum_mismatch", report_date.isoformat(), row, columns[26 + offset],
                                  f"total_{flow}", values[26 + offset], str(computed),
                                  "Stored system total differs from sum of numeric station values; blanks omitted.")
            error_cells = [(position, value) for position, value in cells.items() if types[position] == "e"]
            for (row, col), value in error_cells:
                issue("excel_error", row=row, col=col, value=value,
                      detail="In count table." if (row, col) in consumed else "Outside normalized count table; retained in raw CSV.")
            dates = {record[0] for record in sheet_daily if record[3] == "true"}
            expected_dates = {date(year, month, day).isoformat() for day in range(1, calendar.monthrange(year, month)[1] + 1)}
            if dates != expected_dates or len(sheet_hourly) != 744 or len(sheet_daily) != 31:
                raise ValueError(f"Coverage/count validation failed: {path.name}/{sheet}")
            summaries.append({
                "source_workbook": path.name, "source_sheet": sheet, "source_year": year,
                "source_month": month, "raw_csv": raw_path, "raw_rows": height, "raw_columns": width,
                "raw_nonempty_cells": len(cells), "formula_cells": len(formulas),
                "report_blocks": len(headers), "in_month_days": len(dates),
                "hourly_rows": len(sheet_hourly), "in_month_hourly_rows": len(dates) * 24,
                "daily_rows": len(sheet_daily), "outside_sheet_month_blocks": stats["outside_sheet_month"],
                "blank_count_cells": stats["blank_count_cells"], "numeric_count_cells": stats["numeric_count_cells"],
                "fractional_count_cells": stats["fractional_count_cells"],
                "hour_label_sequence_mismatches": stats["hour_label_sequence_mismatch"],
                "row_date_mismatches": stats["row_date_mismatch"],
                "daily_hourly_sum_mismatches": stats["daily_hourly_sum_mismatch"],
                "system_station_sum_mismatches": stats["system_station_sum_mismatch"],
                "negative_count_cells": stats["negative_count"], "nonnumeric_count_cells": stats["nonnumeric_count"],
                "excel_error_cells": len(error_cells),
                "supplementary_tables": stats["supplementary_table"],
            })
        outputs[f"csv/mrt3_hourly_{year}.csv"] = csv_text(HOURLY, hourly)
        outputs[f"csv/mrt3_daily_{year}.csv"] = csv_text(DAILY, daily)
        manifest.append({"workbook": path.name, "sha256": sha256(path.read_bytes()).hexdigest(),
                         "size_bytes": path.stat().st_size, "sheets": [s[0] for s in sources],
                         "hourly_csv": f"csv/mrt3_hourly_{year}.csv", "hourly_rows": len(hourly),
                         "daily_csv": f"csv/mrt3_daily_{year}.csv", "daily_rows": len(daily)})
    outputs["csv/quality_summary.csv"] = csv_text(list(summaries[0]), (list(s.values()) for s in summaries))
    outputs["csv/data_quality_issues.csv"] = csv_text(ISSUES, issues)
    outputs["csv/source_manifest.json"] = json.dumps({
        "provenance": "Original Excel workbooks acquired through the Philippine Department of Transportation (DOTr) Freedom of Information (FOI) process, as supplied by the data owner.",
        "foi_request_id": None, "foi_request_url": None, "foi_acquisition_date": None,
        "encoding": "UTF-8", "delimiter": ",", "line_ending": "LF",
        "numeric_policy": "Stored Excel numeric text preserved without rounding; blanks remain empty fields.",
        "reconciliation_absolute_tolerance": str(TOLERANCE), "workbooks": manifest,
    }, indent=2) + "\n"
    for relative, content in outputs.items():
        target = ROOT / relative
        data = content.encode("utf-8")
        if check_only:
            if not target.exists() or target.read_bytes() != data:
                raise ValueError(f"Export is missing or differs from the source-derived result: {relative}")
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
    action = "Verified" if check_only else "Wrote"
    print(f"{action} {len(outputs)} files from 3 workbooks / 36 sheets.")
    print(f"Hourly rows: {sum(s['hourly_rows'] for s in summaries):,}; daily rows: {sum(s['daily_rows'] for s in summaries):,} (includes extra month-end blocks).")
    print(f"In-month hourly rows: {sum(s['in_month_hourly_rows'] for s in summaries):,}; report dates: {sum(s['in_month_days'] for s in summaries):,}.")
    print("Issue counts:", dict(Counter(row[0] for row in issues)))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="Rebuild in memory and verify every generated file without writing.")
    convert(parser.parse_args().check)
