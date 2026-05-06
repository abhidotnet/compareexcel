"""Worksheet comparison: headers, number formats, and cell alignment."""

from __future__ import annotations


def _header_to_column_map(ws):
    header_cells = list(ws.iter_rows(min_row=1, max_row=1))[0]
    return {
        str(c.value).strip(): c.column
        for c in header_cells if c.value
    }


def _alignment_tuple(cell):
    if cell is None or cell.alignment is None:
        return ("general", "bottom", 0, False, False, 0)

    al = cell.alignment
    return (
        al.horizontal or "general",
        al.vertical or "bottom",
        al.text_rotation or 0,
        bool(al.wrap_text),
        bool(al.shrink_to_fit),
        al.indent or 0,
    )


def _sample_alignments(ws, col_idx, sample_size=5):
    result = []
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(r, col_idx)
        if cell.value not in (None, ""):
            result.append(_alignment_tuple(cell))
            if len(result) >= sample_size:
                break
    return result


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
        result[header] = fmt or "No data"

    return result


def _analyze_format(fmt):
    if not fmt or fmt == "No data":
        return {"type": "unknown", "raw": fmt}

    f = fmt.lower()
    result = {"raw": fmt}

    if any(x in f for x in ["yy", "mm", "dd"]):
        result["type"] = "date"
    elif "$" in fmt or "€" in fmt or "£" in fmt:
        result["type"] = "currency"
    elif "0" in f or "#" in f:
        result["type"] = "numeric"
    else:
        result["type"] = "text"

    result["decimals"] = len(f.split(".")[-1]) if "." in f else 0
    result["thousands"] = "," in f

    result["currency"] = None
    for sym in ["$", "€", "£"]:
        if sym in fmt:
            result["currency"] = sym

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
                "Differences": "; ".join(diff_details)
            })

    return diffs


def compare_data_alignment(ws1, ws2):
    m1 = _header_to_column_map(ws1)
    m2 = _header_to_column_map(ws2)

    diffs = []

    for col in set(m1) & set(m2):
        a1 = set(_sample_alignments(ws1, m1[col]))
        a2 = set(_sample_alignments(ws2, m2[col]))

        if a1 != a2:
            diffs.append({
                "Column": col,
                "File1_Alignment": str(list(a1)),
                "File2_Alignment": str(list(a2)),
            })

    return diffs


def compare_header_alignment(ws1, ws2):
    m1 = _header_to_column_map(ws1)
    m2 = _header_to_column_map(ws2)

    diffs = []

    for col in set(m1) & set(m2):
        a1 = _alignment_tuple(ws1.cell(1, m1[col]))
        a2 = _alignment_tuple(ws2.cell(1, m2[col]))

        if a1 != a2:
            diffs.append({
                "Column": col,
                "File1_Header_Alignment": str(a1),
                "File2_Header_Alignment": str(a2),
            })

    return diffs
