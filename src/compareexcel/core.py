"""Worksheet comparison: headers, number formats, cell alignment, and currency totals."""

from __future__ import annotations

import math
from typing import Any

import pandas as pd

ALIGN_LABELS = ("horizontal", "vertical", "rotation", "wrap_text", "shrink_to_fit", "indent")


def _header_to_column_map(ws):
    if ws.max_row < 1:
        return {}
    header_cells = list(ws.iter_rows(min_row=1, max_row=1))[0]
    return {
        str(c.value).strip(): c.column
        for c in header_cells
        if c.value is not None and str(c.value).strip()
    }


def _alignment_tuple(cell):
    if cell is None:
        return None
    al = cell.alignment
    if al is None:
        return (None, None, 0, None, None, 0)
    return (
        al.horizontal,
        al.vertical,
        al.text_rotation if al.text_rotation is not None else 0,
        al.wrap_text,
        al.shrink_to_fit,
        al.indent if al.indent is not None else 0,
    )


def format_alignment_tuple(a) -> str:
    if a is None:
        return "None (no data)"
    return "  |  ".join(f"{label}={value}" for label, value in zip(ALIGN_LABELS, a))


def _evenly_sample_rows(sorted_rows: list[int], fraction: float) -> list[int]:
    """Deterministic sample of ~fraction * len(sorted_rows) rows (at least 1 if non-empty)."""
    n = len(sorted_rows)
    if n == 0:
        return []
    k = max(1, min(n, round(n * fraction)))
    if k == n:
        return sorted_rows[:]
    step = n / k
    return [sorted_rows[int(i * step)] for i in range(k)]


def get_column_formats(ws):
    col_map = _header_to_column_map(ws)
    result = {}

    for header, col_idx in col_map.items():
        fmt = None
        for r in range(2, ws.max_row + 1):
            c = ws.cell(r, col_idx)
            if c.value not in (None, ""):
                fmt = c.number_format
                break
        result[header] = fmt if fmt is not None else "No data rows or all blank"

    return result


def _analyze_format(fmt):
    if not fmt or fmt == "No data" or fmt == "No data rows or all blank":
        return {"type": "unknown", "raw": fmt}

    f = str(fmt).lower()
    result = {"raw": fmt}

    if any(x in f for x in ["yy", "mm", "dd"]):
        result["type"] = "date"
    elif "$" in str(fmt) or "€" in str(fmt) or "£" in str(fmt):
        result["type"] = "currency"
    elif "0" in f or "#" in f:
        result["type"] = "numeric"
    elif "@" in str(fmt):
        result["type"] = "text"
    else:
        # e.g. "General" — not a structured number pattern; totals use coercion + per-column rules
        result["type"] = "unknown"

    result["decimals"] = len(f.split(".")[-1]) if "." in f else 0
    result["thousands"] = "," in f

    result["currency"] = None
    for sym in ["$", "€", "£"]:
        if sym in str(fmt):
            result["currency"] = sym
            break

    return result


def _diff_formats(f1, f2):
    a1 = _analyze_format(f1)
    a2 = _analyze_format(f2)
    diffs = []
    for k in set(a1) | set(a2):
        if a1.get(k) != a2.get(k):
            diffs.append(f"{k}: {a1.get(k)} → {a2.get(k)}")
    return diffs


def compare_formatting(ws1, ws2):
    """Compare number_format inferred from the first non-blank data cell per column."""
    f1 = get_column_formats(ws1)
    f2 = get_column_formats(ws2)

    diffs = []
    for col in set(f1) & set(f2):
        if f1[col] != f2[col]:
            diff_details = _diff_formats(f1[col], f2[col])
            diffs.append({
                "Column": col,
                "File1_Format": f1[col],
                "File2_Format": f2[col],
                "Differences": "; ".join(diff_details),
            })

    return diffs


def compare_cell_formatting_sampled(ws1, ws2, sample_fraction: float = 0.1):
    """
    Among rows where both workbooks have non-blank values in a column, take a
    deterministic ~sample_fraction sample of those rows and report number_format
    mismatches for the sampled cells only.
    """
    m1 = _header_to_column_map(ws1)
    m2 = _header_to_column_map(ws2)
    mismatches: list[dict[str, Any]] = []
    max_row = max(ws1.max_row or 1, ws2.max_row or 1)

    for col_name in sorted(set(m1) & set(m2)):
        cidx1, cidx2 = m1[col_name], m2[col_name]
        both_data_rows: list[int] = []
        for row in range(2, max_row + 1):
            v1 = ws1.cell(row, cidx1).value
            v2 = ws2.cell(row, cidx2).value
            if v1 not in (None, "") and v2 not in (None, ""):
                both_data_rows.append(row)

        for row in _evenly_sample_rows(both_data_rows, sample_fraction):
            cell1 = ws1.cell(row, cidx1)
            cell2 = ws2.cell(row, cidx2)
            if cell1.number_format != cell2.number_format:
                mismatches.append({
                    "Column": col_name,
                    "Row": row,
                    "File1_Format": cell1.number_format,
                    "File2_Format": cell2.number_format,
                })

    return mismatches


def compare_data_alignment(ws1, ws2, mismatch_limit: int = 50):
    """
    Row-by-row alignment comparison only for rows where **both** workbooks have
    non-blank values in that column (data cells only).
    Returns a flat list of mismatch records (up to mismatch_limit per column).
    """
    m1 = _header_to_column_map(ws1)
    m2 = _header_to_column_map(ws2)
    diffs: list[dict[str, Any]] = []
    max_row = max(ws1.max_row or 1, ws2.max_row or 1)

    for col_name in sorted(set(m1) & set(m2)):
        cidx1, cidx2 = m1[col_name], m2[col_name]
        stored = 0
        extra = 0

        for row in range(2, max_row + 1):
            c1 = ws1.cell(row, cidx1)
            c2 = ws2.cell(row, cidx2)
            has1 = c1.value not in (None, "")
            has2 = c2.value not in (None, "")
            if not has1 or not has2:
                continue

            a1 = _alignment_tuple(c1)
            a2 = _alignment_tuple(c2)
            if a1 == a2:
                continue

            if stored < mismatch_limit:
                diffs.append({
                    "Column": col_name,
                    "Row": row,
                    "File1_Alignment": format_alignment_tuple(a1),
                    "File2_Alignment": format_alignment_tuple(a2),
                    "Note": "",
                })
                stored += 1
            else:
                extra += 1

        if extra:
            diffs.append({
                "Column": col_name,
                "Row": "",
                "File1_Alignment": "",
                "File2_Alignment": "",
                "Note": f"... and {extra} more row(s) not listed (limit {mismatch_limit} per column)",
            })

    return diffs


ALIGN_ONLY_NO_DATA_NOTE = "no data"


def _compare_data_cells_alignment_and_format_one_per_column(ws1, ws2) -> list[dict[str, Any]]:
    """
    Alignment-only: for each shared column, use the first data row (row 2+) where
    both cells are non-blank. If none, emit one row with Note ``no data``.
    If that cell's alignment or number_format differs, emit a mismatch row.
    """
    m1 = _header_to_column_map(ws1)
    m2 = _header_to_column_map(ws2)
    diffs: list[dict[str, Any]] = []
    max_row = max(ws1.max_row or 1, ws2.max_row or 1)

    for col_name in sorted(set(m1) & set(m2)):
        cidx1, cidx2 = m1[col_name], m2[col_name]
        first_row: int | None = None
        for row in range(2, max_row + 1):
            c1 = ws1.cell(row, cidx1)
            c2 = ws2.cell(row, cidx2)
            if c1.value in (None, "") or c2.value in (None, ""):
                continue
            first_row = row
            break

        if first_row is None:
            diffs.append({
                "Column": col_name,
                "Row": "",
                "File1_Alignment": "",
                "File2_Alignment": "",
                "File1_Format": "",
                "File2_Format": "",
                "Note": ALIGN_ONLY_NO_DATA_NOTE,
            })
            continue

        c1 = ws1.cell(first_row, cidx1)
        c2 = ws2.cell(first_row, cidx2)
        a1, a2 = _alignment_tuple(c1), _alignment_tuple(c2)
        f1, f2 = c1.number_format, c2.number_format
        if a1 == a2 and f1 == f2:
            continue

        diffs.append({
            "Column": col_name,
            "Row": first_row,
            "File1_Alignment": format_alignment_tuple(a1),
            "File2_Alignment": format_alignment_tuple(a2),
            "File1_Format": f1,
            "File2_Format": f2,
            "Note": "",
        })

    return diffs


def compare_data_cells_alignment_and_format(
    ws1,
    ws2,
    mismatch_limit: int = 50,
    *,
    one_cell_per_column: bool = False,
):
    """
    For rows where both sheets have non-blank values in a column, report rows
    where alignment **or** Excel number_format differs.

    With ``one_cell_per_column=True`` (``--alignment-only``), only the first
    paired-data row per column is compared; columns with no such row are
    reported with Note ``no data``.
    """
    if one_cell_per_column:
        return _compare_data_cells_alignment_and_format_one_per_column(ws1, ws2)

    m1 = _header_to_column_map(ws1)
    m2 = _header_to_column_map(ws2)
    diffs: list[dict[str, Any]] = []
    max_row = max(ws1.max_row or 1, ws2.max_row or 1)

    for col_name in sorted(set(m1) & set(m2)):
        cidx1, cidx2 = m1[col_name], m2[col_name]
        stored = 0
        extra = 0

        for row in range(2, max_row + 1):
            c1 = ws1.cell(row, cidx1)
            c2 = ws2.cell(row, cidx2)
            if c1.value in (None, "") or c2.value in (None, ""):
                continue

            a1, a2 = _alignment_tuple(c1), _alignment_tuple(c2)
            f1, f2 = c1.number_format, c2.number_format
            if a1 == a2 and f1 == f2:
                continue

            if stored < mismatch_limit:
                diffs.append({
                    "Column": col_name,
                    "Row": row,
                    "File1_Alignment": format_alignment_tuple(a1),
                    "File2_Alignment": format_alignment_tuple(a2),
                    "File1_Format": f1,
                    "File2_Format": f2,
                    "Note": "",
                })
                stored += 1
            else:
                extra += 1

        if extra:
            diffs.append({
                "Column": col_name,
                "Row": "",
                "File1_Alignment": "",
                "File2_Alignment": "",
                "File1_Format": "",
                "File2_Format": "",
                "Note": f"... and {extra} more row(s) not listed (limit {mismatch_limit} per column)",
            })

    return diffs


def compare_header_alignment(ws1, ws2):
    m1 = _header_to_column_map(ws1)
    m2 = _header_to_column_map(ws2)

    diffs = []
    for col_name in sorted(set(m1) & set(m2)):
        a1 = _alignment_tuple(ws1.cell(1, m1[col_name]))
        a2 = _alignment_tuple(ws2.cell(1, m2[col_name]))
        if a1 != a2:
            diffs.append({
                "Column": col_name,
                "File1_Header_Alignment": format_alignment_tuple(a1),
                "File2_Header_Alignment": format_alignment_tuple(a2),
            })

    return diffs


def worksheet_has_non_blank_data(ws) -> bool:
    """True if any cell has a value other than None or a blank string."""
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if v is None:
                continue
            if isinstance(v, str) and not str(v).strip():
                continue
            return True
    return False


def sheet_data_presence_mismatch(ws1, ws2, sheet_name: str) -> dict[str, Any] | None:
    """
    When one sheet has no data (all blank cells) and the other does, return one
    summary row. If both are empty or both have data, return None.
    """
    h1 = worksheet_has_non_blank_data(ws1)
    h2 = worksheet_has_non_blank_data(ws2)
    if h1 == h2:
        return None
    return {
        "Sheet": sheet_name,
        "File_1": "has data" if h1 else "empty (no data)",
        "File_2": "has data" if h2 else "empty (no data)",
    }


def _currency_columns_from_formats(ws1, ws2) -> set[str]:
    f1 = get_column_formats(ws1)
    f2 = get_column_formats(ws2)
    out: set[str] = set()
    for col in set(f1) & set(f2):
        t1 = _analyze_format(f1[col]).get("type")
        t2 = _analyze_format(f2[col]).get("type")
        if t1 == "currency" or t2 == "currency":
            out.add(col)
    return out


def _numeric_total_eligible_columns(
    df1: pd.DataFrame, df2: pd.DataFrame, ws1, ws2
) -> list[str]:
    """
    Column names (common to both frames) to include in numeric totals.

    Excludes columns whose Excel number_format (first non-blank data cell) is
    classified as **date** or **text** on either sheet, and columns that pandas
    reads as datetime (so labeled dates are not summed).
    """
    df1 = df1.copy()
    df2 = df2.copy()
    df1.columns = [str(c).strip() if c is not None else "" for c in df1.columns]
    df2.columns = [str(c).strip() if c is not None else "" for c in df2.columns]

    f1 = get_column_formats(ws1)
    f2 = get_column_formats(ws2)
    common = sorted(set(df1.columns) & set(df2.columns))
    out: list[str] = []

    for col in common:
        if pd.api.types.is_datetime64_any_dtype(df1[col]) or pd.api.types.is_datetime64_any_dtype(
            df2[col]
        ):
            continue
        raw1 = f1.get(col, "No data rows or all blank")
        raw2 = f2.get(col, "No data rows or all blank")
        t1 = _analyze_format(raw1).get("type")
        t2 = _analyze_format(raw2).get("type")
        if t1 in ("date", "text") or t2 in ("date", "text"):
            continue
        nn1 = pd.to_numeric(df1[col], errors="coerce").notna().sum()
        nn2 = pd.to_numeric(df2[col], errors="coerce").notna().sum()
        if nn1 == 0 and nn2 == 0:
            continue
        out.append(col)

    return out


def compare_numeric_column_totals(
    df1: pd.DataFrame, df2: pd.DataFrame, ws1, ws2, tol: float = 1e-6
) -> list[dict[str, Any]]:
    """
    Sum each eligible column with ``pd.to_numeric(..., errors="coerce")``.

    Eligibility matches :func:`_numeric_total_eligible_columns` (omit date/text
    Excel formats and datetime dtypes). Returns one row per column with totals
    and whether they match within ``tol``.
    """
    df1 = df1.copy()
    df2 = df2.copy()
    df1.columns = [str(c).strip() if c is not None else "" for c in df1.columns]
    df2.columns = [str(c).strip() if c is not None else "" for c in df2.columns]

    cols = _numeric_total_eligible_columns(df1, df2, ws1, ws2)
    rows: list[dict[str, Any]] = []

    for col in cols:
        s1 = pd.to_numeric(df1[col], errors="coerce").sum()
        s2 = pd.to_numeric(df2[col], errors="coerce").sum()
        if isinstance(s1, float) and math.isnan(s1):
            s1 = 0.0
        if isinstance(s2, float) and math.isnan(s2):
            s2 = 0.0
        diff = float(s2) - float(s1)
        match = abs(diff) < tol
        rows.append({
            "Column": col,
            "File1_Total": round(float(s1), 6),
            "File2_Total": round(float(s2), 6),
            "Difference": round(diff, 6),
            "Totals_Match": "Yes" if match else "No",
        })

    return rows


def _quick_numeric_total_columns(df1: pd.DataFrame, df2: pd.DataFrame) -> list[str]:
    """Shared columns with no datetime dtype and at least one coercible numeric value (pandas only)."""
    df1 = df1.copy()
    df2 = df2.copy()
    df1.columns = [str(c).strip() if c is not None else "" for c in df1.columns]
    df2.columns = [str(c).strip() if c is not None else "" for c in df2.columns]
    common = sorted(set(df1.columns) & set(df2.columns))
    out: list[str] = []
    for col in common:
        if pd.api.types.is_datetime64_any_dtype(df1[col]) or pd.api.types.is_datetime64_any_dtype(
            df2[col]
        ):
            continue
        nn1 = pd.to_numeric(df1[col], errors="coerce").notna().sum()
        nn2 = pd.to_numeric(df2[col], errors="coerce").notna().sum()
        if nn1 == 0 and nn2 == 0:
            continue
        out.append(col)
    return out


def compare_numeric_column_totals_quick(df1: pd.DataFrame, df2: pd.DataFrame, tol: float = 1e-6):
    """
    Per-column sums using ``pd.to_numeric`` only (no openpyxl / Excel ``number_format``).

    Intended for quick checks via CLI ``--numeric-column-totals``: excludes datetime dtypes
    and columns with no coercible numeric values in either file.
    """
    df1 = df1.copy()
    df2 = df2.copy()
    df1.columns = [str(c).strip() if c is not None else "" for c in df1.columns]
    df2.columns = [str(c).strip() if c is not None else "" for c in df2.columns]

    cols = _quick_numeric_total_columns(df1, df2)
    rows: list[dict[str, Any]] = []
    for col in cols:
        s1 = pd.to_numeric(df1[col], errors="coerce").sum()
        s2 = pd.to_numeric(df2[col], errors="coerce").sum()
        if isinstance(s1, float) and math.isnan(s1):
            s1 = 0.0
        if isinstance(s2, float) and math.isnan(s2):
            s2 = 0.0
        diff = float(s2) - float(s1)
        match = abs(diff) < tol
        rows.append({
            "Column": col,
            "File1_Total": round(float(s1), 6),
            "File2_Total": round(float(s2), 6),
            "Difference": round(diff, 6),
            "Totals_Match": "Yes" if match else "No",
        })
    return rows


def compare_currency_column_totals(
    df1: pd.DataFrame, df2: pd.DataFrame, ws1, ws2, tol: float = 1e-6
) -> list[dict[str, Any]]:
    """
    Sum every shared column whose Excel format is currency on at least one sheet.

    Returns one row per such column with ``Totals_Match`` (for Excel/HTML reports).
    """
    df1 = df1.copy()
    df2 = df2.copy()
    df1.columns = [str(c).strip() if c is not None else "" for c in df1.columns]
    df2.columns = [str(c).strip() if c is not None else "" for c in df2.columns]

    currency_cols = _currency_columns_from_formats(ws1, ws2)
    rows: list[dict[str, Any]] = []

    for col in sorted(currency_cols):
        if col not in df1.columns or col not in df2.columns:
            continue
        s1 = pd.to_numeric(df1[col], errors="coerce").sum()
        s2 = pd.to_numeric(df2[col], errors="coerce").sum()
        if isinstance(s1, float) and math.isnan(s1):
            s1 = 0.0
        if isinstance(s2, float) and math.isnan(s2):
            s2 = 0.0
        diff = float(s2) - float(s1)
        match = abs(diff) < tol
        rows.append({
            "Column": col,
            "File1_Total": round(float(s1), 6),
            "File2_Total": round(float(s2), 6),
            "Difference": round(diff, 6),
            "Totals_Match": "Yes" if match else "No",
        })

    return rows


def compare_currency_totals(df1: pd.DataFrame, df2: pd.DataFrame, ws1, ws2, tol: float = 1e-6):
    """
    Rows from :func:`compare_currency_column_totals` where totals differ (``Totals_Match`` is ``No``).

    Use :func:`compare_currency_column_totals` when building a full **Excel/HTML** totals table.
    """
    return [r for r in compare_currency_column_totals(df1, df2, ws1, ws2, tol) if r["Totals_Match"] == "No"]
