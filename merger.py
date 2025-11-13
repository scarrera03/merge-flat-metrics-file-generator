#!/usr/bin/env python3
import argparse
import os
import sys
import re
import pandas as pd

def normalize_header(name: str) -> str:
    if not isinstance(name, str):
        name = str(name) if name is not None else ""
    s = name.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s, flags=re.UNICODE)
    return s.strip().lower()

def find_revenue_cashflow_indices(columns):
    norm = [normalize_header(c) for c in columns]
    def match_rev(s): return s == "revenue"
    def match_cf(s): return s == "cash flow" or s == "cashflow"

    try:
        start_idx = next(i for i, c in enumerate(norm) if match_rev(c))
    except StopIteration:
        return None

    try:
        end_idx = next(i for i, c in enumerate(norm) if match_cf(c))
    except StopIteration:
        return None

    if end_idx < start_idx:
        start_idx, end_idx = end_idx, start_idx
    return (start_idx, end_idx)

VIEW_RE = re.compile(r"\d{4}[ABF]")

def detect_view_header_row(raw):
    for r in range(raw.shape[0]):
        row_views = {}
        for c in range(raw.shape[1]):
            val = raw.iat[r, c]
            if isinstance(val, str) and VIEW_RE.fullmatch(val.strip()):
                row_views[c] = val.strip()
        if row_views:
            return r, row_views
    return None, {}

def collect_metric_rows(raw, header_row_idx, view_cols):
    metric_rows = []
    for r in range(header_row_idx + 1, raw.shape[0]):
        label = None
        if raw.shape[1] > 1 and isinstance(raw.iat[r, 1], str):
            label = raw.iat[r, 1].strip()
        elif isinstance(raw.iat[r, 0], str):
            label = raw.iat[r, 0].strip()
        if not label:
            continue
        row_has_num = any(pd.api.types.is_number(raw.iat[r, c]) for c in view_cols.keys())
        if row_has_num:
            metric_rows.append((r, label))
    return metric_rows

def build_flat_from_company_sheets(xlsx_path, month_fixed="December"):
    xls = pd.ExcelFile(xlsx_path)
    records = []

    for sheet in xls.sheet_names:
        raw = pd.read_excel(xlsx_path, sheet_name=sheet, header=None)
        header_row_idx, view_cols = detect_view_header_row(raw)
        if not view_cols:
            continue

        metric_rows = collect_metric_rows(raw, header_row_idx, view_cols)

        for c, view in view_cols.items():
            row = {
                "Company Name": sheet,
                "Year": int(view[:4]),
                "Month": month_fixed,
                "Version": view[-1],
                "View": view,
            }
            any_val = False

            for r, label in metric_rows:
                if "includes depreciation" in normalize_header(label):
                    continue
                val = raw.iat[r, c]
                if pd.api.types.is_number(val):
                    row[label] = float(val)
                    any_val = True

            if any_val:
                records.append(row)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    id_cols = ["Company Name", "Year", "Month", "Version", "View"]
    metric_cols = [c for c in df.columns if c not in id_cols]
    return df[id_cols + metric_cols]

def build_output_csv_path(input_path):
    base, _ = os.path.splitext(input_path)
    return base + ".csv"

def main():
    parser = argparse.ArgumentParser(
        description="Build a flat upload file from Excel company sheets with view codes (2024B/2024F/2025B) and metrics from Revenue to Cash Flow.",
        usage="%(prog)s -i <input_excel_file>"
    )
    parser.add_argument(
        "-i", "--input",
        type=str,
        required=True,
        help="Path to the Excel file (e.g., sample/Data.xlsx)"
    )

    args = parser.parse_args()
    input_path = args.input

    if not os.path.exists(input_path):
        print(f"Error: File '{input_path}' not found.")
        sys.exit(1)

    df_work = build_flat_from_company_sheets(input_path, month_fixed="December")
    if df_work.empty:
        print("Error: No view codes (e.g., 2024B/2024F) found in workbook.")
        sys.exit(1)

    idxs = find_revenue_cashflow_indices(df_work.columns)
    if idxs is None:
        print("Error: 'Revenue' or 'Cash Flow' not found.")
        sys.exit(1)

    start_idx, end_idx = idxs
    cols_to_keep = list(df_work.columns[:start_idx]) + list(df_work.columns[start_idx:end_idx + 1])
    df_final = df_work[cols_to_keep]

    output_csv = build_output_csv_path(input_path)
    df_final.to_csv(output_csv, index=False)

    print(f"Processed: {input_path}")
    print(f"Wrote CSV: {output_csv}")

if __name__ == "__main__":
    main()
