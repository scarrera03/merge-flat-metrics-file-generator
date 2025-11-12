#!/usr/bin/env python3
import argparse
import os
import sys
import re
import pandas as pd

# -------------------------------
# Utilities
# -------------------------------

def normalize_header(name: str) -> str:
    """Lowercase, collapse spaces, strip."""
    if not isinstance(name, str):
        name = str(name) if name is not None else ""
    s = name.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s, flags=re.UNICODE)
    return s.strip().lower()

def find_revenue_cashflow_indices(columns):
    """Return (start_idx, end_idx) for 'Revenue'..'Cash Flow' (inclusive) or None."""
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

# -------------------------------
# Reading "already-flat" sheets
# -------------------------------

def load_sheet_plain(xlsx_path, preferred_sheet):
    """
    Attempt to read a sheet that already has headers including 'Revenue' and 'Cash Flow'.
    If preferred_sheet is given, use that directly; else scan sheets and return the first that matches.
    Returns (df, sheet_name) or (None, None) if no plain sheet matches.
    """
    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception as e:
        print(f"Error: Failed to open '{xlsx_path}': {e}")
        sys.exit(1)

    if preferred_sheet:
        if preferred_sheet not in xls.sheet_names:
            print(f"Error: Sheet '{preferred_sheet}' not found. Available: {xls.sheet_names}")
            sys.exit(1)
        df = pd.read_excel(xlsx_path, sheet_name=preferred_sheet)
        if find_revenue_cashflow_indices(df.columns):
            return df, preferred_sheet
        return None, None

    for s in xls.sheet_names:
        df_try = pd.read_excel(xlsx_path, sheet_name=s)
        if find_revenue_cashflow_indices(df_try.columns):
            return df_try, s

    return None, None

# -------------------------------
# Building flat wide table from company sheets
# -------------------------------

VIEW_RE = re.compile(r"\d{4}[ABF]")  # 2024B, 2024F, 2025A, etc.

def detect_view_header_row(raw):
    """Return (row_index, {col_index: view_code}) or (None, {})."""
    header_row_idx = None
    view_cols = {}
    for r in range(raw.shape[0]):
        # if this row has any view code, treat it as header
        row_views = {}
        for c in range(raw.shape[1]):
            val = raw.iat[r, c]
            if isinstance(val, str) and VIEW_RE.fullmatch(val.strip()):
                row_views[c] = val.strip()
        if row_views:
            header_row_idx = r
            view_cols = row_views
            break
    if not view_cols:
        return None, {}
    return header_row_idx, view_cols

def collect_metric_rows(raw, header_row_idx, view_cols):
    """
    Return list of (row_idx, label) for metric rows beneath the header row.
    Labels are usually in column 1 (or col 0 as fallback).
    """
    metric_rows = []
    for r in range(header_row_idx + 1, raw.shape[0]):
        label = None
        if raw.shape[1] > 1 and isinstance(raw.iat[r, 1], str):
            label = raw.iat[r, 1].strip()
        elif isinstance(raw.iat[r, 0], str):
            label = raw.iat[r, 0].strip()
        if not label:
            continue
        # keep only rows that have any numeric under the view columns
        row_has_num = any(pd.api.types.is_number(raw.iat[r, c]) for c in view_cols.keys())
        if row_has_num:
            metric_rows.append((r, label))
    return metric_rows

def build_flat_from_company_sheets(xlsx_path, month_fixed="December"):
    """
    For workbooks where each sheet = a company, with a row of view codes (2024B/F/A)
    and metric labels in the first columns:
      -> Build a flat wide table:
         [Company Name, Year, Month, Version, View, <metrics as columns>]
    """
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
                # Skip "Includes Depreciation" anywhere in the label
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

    # Reorder: id columns first
    id_cols = ["Company Name", "Year", "Month", "Version", "View"]
    metric_cols = [c for c in df.columns if c not in id_cols]
    return df[id_cols + metric_cols]

# -------------------------------
# Main CLI
# -------------------------------

def build_output_csv_path(input_path):
    base, _ = os.path.splitext(input_path)
    return base + ".csv"

def main():
    parser = argparse.ArgumentParser(
        description="Extract columns from 'Revenue' to 'Cash Flow' (inclusive). Handles both a flat sheet and multi-sheet company workbooks.",
        usage="%(prog)s -i <input_excel_file> [--sheet SHEET_NAME] [--list-sheets]"
    )
    parser.add_argument("-i", "--input", type=str, required=False,
                        help="Path to the Excel file to process (e.g., data.xlsx)")
    parser.add_argument("--sheet", type=str, default=None,
                        help="Optional sheet name to read (for an already-flat sheet).")
    parser.add_argument("--list-sheets", action="store_true",
                        help="List sheet names and exit.")

    args = parser.parse_args()

    if not any([args.input, args.list_sheets]):
        parser.print_help(sys.stderr)
        sys.exit(1)

    if args.list_sheets:
        if not args.input:
            print("Error: --list-sheets requires -i/--input to be provided.")
            sys.exit(1)
        try:
            xls = pd.ExcelFile(args.input)
        except Exception as e:
            print(f"Error: Failed to open '{args.input}': {e}")
            sys.exit(1)
        print("Sheets:")
        for s in xls.sheet_names:
            print(f" - {s}")
        sys.exit(0)

    input_path = args.input
    if not input_path or not os.path.exists(input_path):
        print(f"Error: File '{input_path}' not found.")
        sys.exit(1)

    # 1) Try to load an already-flat sheet (with headers including Revenue/Cash Flow)
    df_plain, sheet_used = load_sheet_plain(input_path, args.sheet)

    if df_plain is None:
        # 2) Build flat table from company sheets
        df_flat = build_flat_from_company_sheets(input_path, month_fixed="December")
        if df_flat.empty:
            print("Error: Could not build a flat table from the workbook; no view headers found (e.g., 2024B/2024F).")
            sys.exit(1)
        df_work = df_flat
        sheet_used = "(combined from all company sheets)"
    else:
        df_work = df_plain

    # 3) Now slice columns from Revenue .. Cash Flow
    idxs = find_revenue_cashflow_indices(df_work.columns)
    if idxs is None:
        print("Error: 'Revenue' or 'Cash Flow' column not found after processing.")
        print(f"Sheet used: {sheet_used}")
        print("Detected columns:")
        for c in df_work.columns:
            print(f" - '{c}'  (normalized: '{normalize_header(c)}')")
        sys.exit(1)

    start_idx, end_idx = idxs
    cols_before_revenue = list(df_work.columns[:start_idx])  # keep metadata before Revenue
    cols_main = list(df_work.columns[start_idx:end_idx + 1]) # Revenue..Cash Flow inclusive
    cols_to_keep = cols_before_revenue + cols_main

    # Ensure Cash Flow is last; drop any metrics after
    df_final = df_work[cols_to_keep]

    # 4) Write CSV with same base name
    output_csv = build_output_csv_path(input_path)
    try:
        df_final.to_csv(output_csv, index=False)
    except Exception as e:
        print(f"Error: Failed writing CSV '{output_csv}': {e}")
        sys.exit(1)

    print(f"Processed: {sheet_used}")
    print(f"Wrote CSV: {output_csv}")

if __name__ == "__main__":
    main()