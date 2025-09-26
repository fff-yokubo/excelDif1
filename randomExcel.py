# Creating an Excel sample file with ~100 rows, many columns, and merged cells.
# The generated file will be saved to /mnt/data and the path will be printed at the end.
# This code runs in the notebook environment; you'll get a download link after execution.

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
from caas_jupyter_tools import display_dataframe_to_user

out_path = "/mnt/data/sample_excel_100rows_20cols_merged.xlsx"
n_rows = 100
n_cols = 20

# Create column names
cols = [f"Col{c+1}" for c in range(n_cols)]

# Generate random data: mixture of integers, floats, dates, and strings
rng = np.random.default_rng(12345)
data = {}
for i, col in enumerate(cols):
    t = i % 4
    if t == 0:  # integers
        data[col] = rng.integers(0, 1000, size=n_rows)
    elif t == 1:  # floats
        data[col] = (rng.random(n_rows) * 1000).round(2)
    elif t == 2:  # dates
        start = datetime(2020, 1, 1)
        data[col] = [(start + timedelta(days=int(x))).date() for x in rng.integers(0, 2000, size=n_rows)]
    else:  # random strings (short)
        words = ["東京","大阪","名古屋","札幌","福岡","横浜","神戸","京都","仙台","広島"]
        data[col] = [f"{words[int(x) % len(words)]}_{int(x)}" for x in rng.integers(0, 10000, size=n_rows)]

df = pd.DataFrame(data, columns=cols)

# Save to Excel with openpyxl engine, then manipulate workbook to add merged cells
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=1)  # leave first row for group headers
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    # Create merged header groups in the first row
    # Merge Col1-Col4 as "Group A", Col5-Col8 as "Group B", Col9-Col12 as "Group C", Col13-Col20 as "Group D"
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    worksheet.cell(row=1, column=1, value="Group A")
    worksheet.merge_cells(start_row=1, start_column=5, end_row=1, end_column=8)
    worksheet.cell(row=1, column=5, value="Group B")
    worksheet.merge_cells(start_row=1, start_column=9, end_row=1, end_column=12)
    worksheet.cell(row=1, column=9, value="Group C")
    worksheet.merge_cells(start_row=1, start_column=13, end_row=1, end_column=20)
    worksheet.cell(row=1, column=13, value="Group D")

    # Add a few merged blocks inside the data area (some horizontal, some vertical)
    # Horizontal merges: merge 2-3 rows worth of cells horizontally in random rows
    merges = [
        (5, 2, 5, 4),   # row 5, cols 2-4
        (12, 6, 12, 9), # row 12, cols 6-9
        (30, 1, 30, 3), # row 30, cols 1-3
        (50, 10, 50, 14) # row 50, cols 10-14
    ]
    for r1, c1, r2, c2 in merges:
        # Account for header offset (data starts at row 2 in Excel since startrow=1)
        worksheet.merge_cells(start_row=r1+1, start_column=c1, end_row=r2+1, end_column=c2)
        worksheet.cell(row=r1+1, column=c1, value=f"Merged_{r1}_{c1}-{c2}")

    # Vertical merges: merge multiple rows in the same column
    v_merges = [
        (20, 2, 22, 2),  # rows 20-22 in col2
        (70, 5, 75, 5)   # rows 70-75 in col5
    ]
    for r1, c1, r2, c2 in v_merges:
        worksheet.merge_cells(start_row=r1+1, start_column=c1, end_row=r2+1, end_column=c2)
        worksheet.cell(row=r1+1, column=c1, value=f"V_Merged_{r1}_{r2}")

# Display a quick preview to the user (first 8 rows)
preview = df.head(8).copy()
display_dataframe_to_user("Preview of generated Excel (first 8 rows)", preview)

print(f"Generated Excel file: {out_path}")
