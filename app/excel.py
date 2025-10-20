# app/excel.py
import io
import pandas as pd
from typing import Optional
from openpyxl.utils import get_column_letter

def build_excel_bytes(df: pd.DataFrame) -> bytes:
    """Write the labels DataFrame to an Excel file in memory and return bytes."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # write a summary sheet and a full sheet
        df.to_excel(writer, sheet_name="labels", index=False)
        # auto width columns for readability
        ws = writer.sheets["labels"]
        for i, col in enumerate(df.columns, 1):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            ws.column_dimensions[get_column_letter(i)].width = max_len
        writer.save()
    output.seek(0)
    return output.read()
