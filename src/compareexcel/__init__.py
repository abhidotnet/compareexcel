"""Compare Excel workbooks for formatting and alignment differences."""

from compareexcel.core import (
    compare_data_alignment,
    compare_formatting,
    compare_header_alignment,
    get_column_formats,
)
from compareexcel.report import write_report

__version__ = "0.1.0"

__all__ = [
    "__version__",
    "compare_data_alignment",
    "compare_formatting",
    "compare_header_alignment",
    "get_column_formats",
    "write_report",
]
