"""Write comparison results to an Excel workbook."""

from __future__ import annotations

import pandas as pd


def write_report(output_file, formatting, data_align, header_align):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        if formatting:
            pd.DataFrame(formatting).to_excel(writer, sheet_name="Formatting_Diffs", index=False)

        if data_align:
            pd.DataFrame(data_align).to_excel(writer, sheet_name="Data_Alignment_Diffs", index=False)

        if header_align:
            pd.DataFrame(header_align).to_excel(writer, sheet_name="Header_Alignment_Diffs", index=False)
