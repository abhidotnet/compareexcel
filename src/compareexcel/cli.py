"""Command-line interface for comparing two Excel files."""

from __future__ import annotations

import argparse
import sys

from openpyxl import load_workbook

from compareexcel.core import (
    compare_data_alignment,
    compare_formatting,
    compare_header_alignment,
)
from compareexcel.report import write_report


def _ensure_utf8_stdio():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def main():
    _ensure_utf8_stdio()

    parser = argparse.ArgumentParser(
        description="Compare two Excel workbooks for formatting and alignment differences.",
    )
    parser.add_argument("file1")
    parser.add_argument("file2")
    parser.add_argument("--sheet")
    parser.add_argument("--alignment-only", "--ao", action="store_true")
    parser.add_argument("--output", help="Output Excel report file")

    args = parser.parse_args()

    wb1 = load_workbook(args.file1)
    wb2 = load_workbook(args.file2)

    sheets = [args.sheet] if args.sheet else set(wb1.sheetnames) & set(wb2.sheetnames)

    all_format_diffs = []
    all_data_align_diffs = []
    all_header_align_diffs = []

    for sheet in sheets:
        ws1 = wb1[sheet]
        ws2 = wb2[sheet]

        fmt = compare_formatting(ws1, ws2)
        da = compare_data_alignment(ws1, ws2)
        ha = compare_header_alignment(ws1, ws2)

        for row in fmt:
            row["Sheet"] = sheet
        for row in da:
            row["Sheet"] = sheet
        for row in ha:
            row["Sheet"] = sheet

        all_format_diffs.extend(fmt)
        all_data_align_diffs.extend(da)
        all_header_align_diffs.extend(ha)

    print("\n=== SUMMARY ===")
    print(f"Formatting diffs: {len(all_format_diffs)}")
    print(f"Data alignment diffs: {len(all_data_align_diffs)}")
    print(f"Header alignment diffs: {len(all_header_align_diffs)}")

    if args.output:
        write_report(
            args.output,
            all_format_diffs,
            all_data_align_diffs,
            all_header_align_diffs,
        )
        print(f"\nReport written to: {args.output}")
