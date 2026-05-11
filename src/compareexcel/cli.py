"""Command-line interface for comparing two Excel files."""

from __future__ import annotations

import argparse
import os
import sys

import pandas as pd
from openpyxl import load_workbook

from compareexcel.core import (
    ALIGN_ONLY_NO_DATA_NOTE,
    compare_cell_formatting_sampled,
    compare_currency_totals,
    compare_data_alignment,
    compare_data_cells_alignment_and_format,
    compare_formatting,
    compare_header_alignment,
    compare_numeric_column_totals,
    sheet_data_presence_mismatch,
)
from compareexcel.report import write_report


def _ensure_utf8_stdio():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def _print_section(title: str, rows: list, describe_empty: str):
    print(title)
    if not rows:
        print(f"  {describe_empty}")
        return
    for r in rows:
        if isinstance(r, dict):
            parts = [f"{k}: {v}" for k, v in r.items() if v not in ("", None)]
            line = "  " + " | ".join(parts)
            if str(r.get("Note", "")).strip().lower() == ALIGN_ONLY_NO_DATA_NOTE.lower():
                line = f"\033[32m{line}\033[0m"
            print(line)
        else:
            print(f"  {r}")
    print()


def main():
    _ensure_utf8_stdio()

    parser = argparse.ArgumentParser(
        description=(
            "Compare two Excel workbooks: formatting, currency totals, header alignment, "
            "and data-cell alignment (full mode), or alignment + number format on data rows only "
            "(--alignment-only)."
        ),
    )
    parser.add_argument("file1", help="First workbook (.xlsx)")
    parser.add_argument("file2", help="Second workbook (.xlsx)")
    parser.add_argument("--sheet", help="Compare only this sheet (must exist in both)")
    parser.add_argument(
        "--alignment-only",
        "--ao",
        action="store_true",
        help=(
            "Header alignment and one compared cell per column (first row where both files have "
            "data in that column), including alignment and number_format. Columns with no such row "
            "are listed as 'no data' (green in report). Skips pandas totals, column summary, and "
            "sampled format pass. --sample-fraction is ignored."
        ),
    )
    parser.add_argument(
        "--output",
        "-o",
        help="Write report to this path (.xlsx or .html)",
    )
    parser.add_argument(
        "--sample-fraction",
        type=float,
        default=0.1,
        help=(
            "Full mode only: fraction of per-column rows (both sides non-blank) to check for "
            "cell number_format (default: 0.1). Ignored with --alignment-only."
        ),
    )
    parser.add_argument(
        "--numeric-column-totals",
        "-nct",
        action="store_true",
        help=(
            "Full mode only: for each shared column, sum values with pandas after excluding "
            "columns classified as date or text from Excel number_format (first data cell per "
            "column) and excluding datetime columns. Writes sheet Numeric_Column_Totals when "
            "using --output, and prints a console section. Ignored with --alignment-only."
        ),
    )

    args = parser.parse_args()

    if not os.path.isfile(args.file1):
        print(f"Error: file not found: {args.file1}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(args.file2):
        print(f"Error: file not found: {args.file2}", file=sys.stderr)
        sys.exit(1)

    if not (0 < args.sample_fraction <= 1):
        print("--sample-fraction must be between 0 and 1 (exclusive of 0).", file=sys.stderr)
        sys.exit(1)

    wb1 = load_workbook(args.file1, data_only=False)
    wb2 = load_workbook(args.file2, data_only=False)

    names1 = set(wb1.sheetnames)
    names2 = set(wb2.sheetnames)
    if args.sheet:
        if args.sheet not in names1 or args.sheet not in names2:
            print(
                f"Error: sheet {args.sheet!r} missing from one or both workbooks.",
                file=sys.stderr,
            )
            sys.exit(1)
        sheets = [args.sheet]
    else:
        sheets = sorted(names1 & names2)
        if names1 != names2:
            print("Sheet names differ between files.")
            if names1 - names2:
                print(f"  Only in file 1: {sorted(names1 - names2)}")
            if names2 - names1:
                print(f"  Only in file 2: {sorted(names2 - names1)}")
            print(f"  Comparing common sheets only: {sheets}\n")

    dfs1: dict = {}
    dfs2: dict = {}
    if not args.alignment_only:
        try:
            if args.sheet:
                dfs1 = {args.sheet: pd.read_excel(args.file1, sheet_name=args.sheet)}
                dfs2 = {args.sheet: pd.read_excel(args.file2, sheet_name=args.sheet)}
            else:
                dfs1 = pd.read_excel(args.file1, sheet_name=None)
                dfs2 = pd.read_excel(args.file2, sheet_name=None)
        except Exception as e:
            print(f"Error reading workbooks with pandas: {e}", file=sys.stderr)
            sys.exit(1)

    all_column_fmt: list[dict] = []
    all_cell_fmt: list[dict] = []
    all_currency: list[dict] = []
    all_numeric_totals: list[dict] = []
    all_header_align: list[dict] = []
    all_data_align: list[dict] = []
    all_align_and_format: list[dict] = []
    all_sheet_presence: list[dict] = []

    for sheet in sheets:
        ws1 = wb1[sheet]
        ws2 = wb2[sheet]

        if not args.alignment_only:
            if sheet not in dfs1 or sheet not in dfs2:
                continue

        pres = sheet_data_presence_mismatch(ws1, ws2, sheet)
        if pres:
            all_sheet_presence.append(pres)

        if not args.alignment_only:
            df1 = dfs1[sheet]
            df2 = dfs2[sheet]
            for row in compare_formatting(ws1, ws2):
                row = {**row, "Sheet": sheet}
                all_column_fmt.append(row)
            for row in compare_currency_totals(df1, df2, ws1, ws2):
                row = {**row, "Sheet": sheet}
                all_currency.append(row)
            if args.numeric_column_totals:
                for row in compare_numeric_column_totals(df1, df2, ws1, ws2):
                    row = {**row, "Sheet": sheet}
                    all_numeric_totals.append(row)
            for row in compare_cell_formatting_sampled(ws1, ws2, args.sample_fraction):
                row = {**row, "Sheet": sheet}
                all_cell_fmt.append(row)
            for row in compare_data_alignment(ws1, ws2):
                row = {**row, "Sheet": sheet}
                all_data_align.append(row)

        for row in compare_header_alignment(ws1, ws2):
            row = {**row, "Sheet": sheet}
            all_header_align.append(row)

        if args.alignment_only:
            for row in compare_data_cells_alignment_and_format(
                ws1, ws2, one_cell_per_column=True
            ):
                row = {**row, "Sheet": sheet}
                all_align_and_format.append(row)

    print()
    print("=== Excel comparison ===")
    print(f"File 1: {os.path.basename(args.file1)}")
    print(f"File 2: {os.path.basename(args.file2)}")
    if args.alignment_only:
        print(
            "Mode: header alignment + one paired-data cell per column (alignment & format); "
            "empty paired columns → 'no data' (green). (--alignment-only)"
        )
    elif args.numeric_column_totals:
        print(
            "Numeric column totals: enabled (--numeric-column-totals); "
            "date/text Excel formats and datetime columns are excluded."
        )
    print()

    _print_section(
        "Sheet empty vs data mismatch — (same sheet name, one workbook has no cell data)",
        all_sheet_presence,
        "None — both files have data on each compared sheet, or both are empty.",
    )

    if not args.alignment_only:
        _print_section(
            "Column format summary (first data cell per column) —",
            all_column_fmt,
            "No differences.",
        )
        _print_section(
            f"Formatting cell mismatch — (~{args.sample_fraction:.0%} sample of rows with data in both files, per column)",
            all_cell_fmt,
            "No mismatches in sample.",
        )
        _print_section(
            "Amount total mismatch — (currency columns by Excel number format)",
            all_currency,
            "No mismatches.",
        )
        if args.numeric_column_totals:
            _print_section(
                "Numeric column totals — (all eligible columns; date/text formats excluded)",
                all_numeric_totals,
                "No eligible columns after excluding date/text/datetime, or no data.",
            )

    _print_section(
        "Header alignment mismatch —",
        all_header_align,
        "No mismatches.",
    )

    if args.alignment_only:
        _print_section(
            "Data cell alignment & number format — (first row per column where both files have data; "
            "Note: no data = no such row, green in Excel/HTML)",
            all_align_and_format,
            "No issues — no mismatches on compared cells and no empty paired columns.",
        )
    else:
        _print_section(
            "Data alignment mismatch — (rows with data in both files)",
            all_data_align,
            "No mismatches.",
        )

    if args.output:
        write_report(
            args.output,
            formatting_cells_sampled=all_cell_fmt if not args.alignment_only else None,
            data_align=all_data_align if not args.alignment_only else None,
            header_align=all_header_align,
            currency_totals=all_currency if not args.alignment_only else None,
            numeric_column_totals=all_numeric_totals
            if (not args.alignment_only and args.numeric_column_totals)
            else None,
            column_formatting=all_column_fmt if not args.alignment_only else None,
            sheet_presence=all_sheet_presence,
            data_align_with_format=all_align_and_format if args.alignment_only else None,
        )
        print(f"Report written to: {args.output}")


if __name__ == "__main__":
    main()
