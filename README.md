# Excel Compare (compareExcel)

**compareExcel** is a small command-line tool and Python library that compares two Excel (`.xlsx`) workbooks. For columns that exist in both files (matched by header text in row 1), it can report:

- **Column number format** — inferred from the first non-blank data cell per column (summary).
- **Cell number format (sampled)** — compares `number_format` on a deterministic sample (~10% by default) of rows where **both** workbooks have non-blank values in that column.
- **Header alignment** — alignment attributes for row-1 header cells.
- **Data cell alignment** — row-by-row alignment only for rows where **both** workbooks have non-blank values in that column; output lists up to 50 mismatches per column, plus a count of any additional mismatches.
- **Amount totals** — for columns treated as **currency** from Excel number format (e.g. `$`, `€`, `£`), sums numeric values with pandas and flags total mismatches.
- **Numeric column totals** (optional, `--numeric-column-totals`) — sums every shared column except those classified as **date** or **text** from the first data cell’s Excel format, and except **datetime** columns as read by pandas; writes a dedicated report sheet / HTML section with per-column totals and a match flag.

Console output uses readable section titles (for example “Header alignment mismatch —”, “Amount total mismatch —”). A short summary is always printed; use `--output` to write a full **Excel** or **HTML** report.

Install from PyPI with `pip install compareexcel`. The distribution and import package name is **`compareexcel`**. Installed console scripts are **`compareexcel`** and **`compareExcel`** (same entry point).

## Requirements

- Python 3.10+
- Dependencies: `pandas`, `openpyxl` (declared in `pyproject.toml`)

## Installation

From the repository root (`compareexcel/`):

```bash
pip install .
```

Editable install while developing:

```bash
pip install -e .
```

## Command-line usage

```text
compareExcel FILE1 FILE2 [--sheet SHEET] [--output PATH] [-o PATH]
    [--sample-fraction FRACTION] [--alignment-only] [--numeric-column-totals]
```

| Argument / option | Description |
|-------------------|-------------|
| `FILE1`, `FILE2` | Paths to the two `.xlsx` files to compare. |
| `--sheet` | Compare only this sheet; it must exist in both workbooks. Default: every sheet name that appears in **both** files (in sorted order). If sheet names differ, only the intersection is compared and a note is printed. |
| `--output` / `-o` | Write a report file. **`.xlsx`** → multi-sheet workbook; **`.html`** or **`.htm`** → single HTML page. Format is chosen from the file extension. |
| `--sample-fraction` | **Full mode only:** fraction (0–1] of per-column “both sides non-blank” rows to check for **cell** `number_format` mismatches. Default: `0.1` (~10%). Ignored with `--alignment-only`. |
| `--alignment-only` / `--ao` | **No pandas reads** (no currency totals). Compares **header alignment** and, for each column, the **first data row** where **both** files have a non-blank value: **alignment and `number_format`**. Columns with **no** such row are listed with Note **no data** (green text in Excel/HTML; green ANSI in the console). Skips column format summary and the sampled cell-format pass. |
| `--numeric-column-totals` / `-nct` | **Full mode only:** compute per-column sums for all eligible columns (excludes date/text formats and datetime dtypes). Adds console output and, with `--output`, sheet `Numeric_Column_Totals` or an HTML section. Ignored with `--alignment-only`. |

### Excel report sheets (when using `.xlsx`)

Sheets are created only when there is data for that category (otherwise a minimal **Summary** sheet is written):

| Sheet name | Content |
|------------|---------|
| `Column_Format_Summary` | First data cell `number_format` per column (skipped in `--alignment-only` mode when writing from the CLI). |
| `Cell_Format_Sampled` | Rows where the sampled cells’ number formats differ. |
| `Amount_Total_Mismatch` | Currency columns whose numeric totals differ. |
| `Numeric_Column_Totals` | **With `--numeric-column-totals`:** every eligible column’s `File1_Total`, `File2_Total`, `Difference`, `Totals_Match`, and `Sheet`. |
| `Header_Alignment_Mismatch` | Header row alignment differences. |
| `Data_Alignment_Mismatch` | Data row alignment differences, **full mode only** (capped per column). |
| `Data_Cell_Align_Format` | **`--alignment-only`:** one compared cell per column, or **no data** (green); yellow highlights alignment columns when they differ. |

### Examples

```bash
compareExcel workbook_a.xlsx workbook_b.xlsx --output diff_report.xlsx
compareExcel workbook_a.xlsx workbook_b.xlsx -o diff_report.html
compareExcel workbook_a.xlsx workbook_b.xlsx --sheet "Summary" -o out.xlsx --sample-fraction 0.15
compareExcel workbook_a.xlsx workbook_b.xlsx --alignment-only -o align.html
compareExcel workbook_a.xlsx workbook_b.xlsx --numeric-column-totals -o totals.xlsx
```

## Library usage

Core comparisons take **openpyxl** worksheet objects. Currency totals need **pandas** `DataFrame`s for the same sheet (column names should match the header row; the helper normalizes column labels to strings).

```python
import pandas as pd
from openpyxl import load_workbook

from compareexcel import (
    compare_cell_formatting_sampled,
    compare_currency_totals,
    compare_data_alignment,
    compare_data_cells_alignment_and_format,
    compare_formatting,
    compare_header_alignment,
    compare_numeric_column_totals,
    write_report,
)

wb1 = load_workbook("a.xlsx", data_only=False)
wb2 = load_workbook("b.xlsx", data_only=False)
sheet = wb1.sheetnames[0]
ws1, ws2 = wb1[sheet], wb2[sheet]

df1 = pd.read_excel("a.xlsx", sheet_name=sheet)
df2 = pd.read_excel("b.xlsx", sheet_name=sheet)

column_fmt = compare_formatting(ws1, ws2)
cell_fmt = compare_cell_formatting_sampled(ws1, ws2, sample_fraction=0.1)
currency = compare_currency_totals(df1, df2, ws1, ws2)
numeric_totals = compare_numeric_column_totals(df1, df2, ws1, ws2)
header_align = compare_header_alignment(ws1, ws2)
data_align = compare_data_alignment(ws1, ws2)  # optional: mismatch_limit=50

write_report(
    "out.xlsx",
    formatting_cells_sampled=cell_fmt,
    data_align=data_align,
    header_align=header_align,
    currency_totals=currency,
    numeric_column_totals=numeric_totals,
    column_formatting=column_fmt,
)
```

`write_report` chooses **Excel** vs **HTML** from the path suffix (`.html` / `.htm` → HTML; otherwise Excel). Pass `column_formatting=None` to omit the column summary section in HTML or the `Column_Format_Summary` sheet in Excel.

For an alignment-focused report matching the CLI’s `--alignment-only` behavior, call `compare_data_cells_alignment_and_format(ws1, ws2, one_cell_per_column=True)` and pass the result as `data_align_with_format=...`, and set `formatting_cells_sampled`, `data_align`, and `currency_totals` to `None` if you want those sections omitted from HTML / Excel. Use `compare_data_cells_alignment_and_format(ws1, ws2)` (default) to scan all paired-data rows with a per-column cap instead.

## Package layout

```text
compareexcel/
├── src/
│   └── compareexcel/
│       ├── __init__.py   # Public API and version
│       ├── cli.py        # Entry point and argument parsing
│       ├── core.py       # Comparison logic
│       └── report.py     # Excel / HTML report writers
├── pyproject.toml
├── README.md
├── LICENSE
└── .gitignore
```

## License

See [LICENSE](LICENSE).
