# Excel Compare (compareExcel)

**compareExcel** is a small command-line tool and Python library that compares two Excel (`.xlsx`) workbooks. It reports differences in **column number formats**, **data cell alignment** (sampled from non-empty cells), and **header row alignment** for columns that exist in both files.

Installable package name on PyPI is normalized as `compareexcel`; the installed console script is **`compareExcel`**.

## Requirements

- Python 3.10+
- Dependencies: `pandas`, `openpyxl` (declared in `pyproject.toml`)

## Installation

From the repository root (`excel-compare/`):

```bash
pip install .
```

Editable install while developing:

```bash
pip install -e .
```

## Command-line usage

```text
compareExcel FILE1 FILE2 [--sheet SHEET_NAME] [--output REPORT.xlsx] [--alignment-only]
```

| Argument / option | Description |
|-------------------|-------------|
| `FILE1`, `FILE2` | Paths to the two Excel files to compare. |
| `--sheet` | Compare only this sheet; default is all sheets that exist in **both** workbooks. |
| `--output` | Write a multi-sheet Excel report (`Formatting_Diffs`, `Data_Alignment_Diffs`, `Header_Alignment_Diffs`). |
| `--alignment-only` / `--ao` | Reserved for future use (currently parsed only). |

Example:

```bash
compareExcel workbook_a.xlsx workbook_b.xlsx --output diff_report.xlsx
```

A short summary is always printed to the console; use `--output` to persist detailed rows to an `.xlsx` file.

## Library usage

```python
from openpyxl import load_workbook
from excel_compare import (
    compare_formatting,
    compare_data_alignment,
    compare_header_alignment,
    write_report,
)

wb1 = load_workbook("a.xlsx")
wb2 = load_workbook("b.xlsx")
sheet = wb1.sheetnames[0]
fmt = compare_formatting(wb1[sheet], wb2[sheet])
data_align = compare_data_alignment(wb1[sheet], wb2[sheet])
header_align = compare_header_alignment(wb1[sheet], wb2[sheet])

write_report("out.xlsx", fmt, data_align, header_align)
```

## Package layout

```text
excel-compare/
├── src/
│   └── excel_compare/
│       ├── __init__.py   # Public API and version
│       ├── cli.py        # Entry point and argument parsing
│       ├── core.py       # Comparison logic
│       └── report.py     # Excel report writer
├── pyproject.toml
├── README.md
├── LICENSE
└── .gitignore
```

## License

See [LICENSE](LICENSE).
