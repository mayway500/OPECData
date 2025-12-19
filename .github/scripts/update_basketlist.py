#!/usr/bin/env python3
"""
Copy values from sheet named 'oilpricechart' to a sheet named 'Basketlist' inside the same Excel file.
This script preserves cell values and basic merged cell layout; it does not attempt to copy advanced styling.
"""
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def clear_sheet(ws):
    # Remove all rows by clearing cell values and unmerging
    merges = list(ws.merged_cells.ranges)
    for m in merges:
        ws.unmerge_cells(str(m))
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None


def copy_sheet_values(src, dst):
    # Clear dst first
    clear_sheet(dst)

    # copy values and merged cells
    max_row = src.max_row
    max_col = src.max_column

    # copy values
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = src.cell(row=r, column=c).value
            dst.cell(row=r, column=c).value = val

    # copy merged cells
    for m in src.merged_cells.ranges:
        dst.merge_cells(str(m))

    # copy column widths (best-effort)
    try:
        for i, col in enumerate(src.column_dimensions, start=1):
            # src.column_dimensions uses keys like 'A', 'B'
            if col in src.column_dimensions:
                width = src.column_dimensions[col].width
                if width:
                    dst.column_dimensions[col].width = width
    except Exception:
        pass


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: update_basketlist.py path/to/Opecpricechart.xlsx")
        sys.exit(2)

    xlsx = sys.argv[1]
    print(f"Opening workbook: {xlsx}")

    wb = load_workbook(xlsx)

    src_name = 'oilpricechart'
    dst_name = 'Basketlist'

    if src_name not in wb.sheetnames:
        print(f"Source sheet '{src_name}' not found in workbook.")
        sys.exit(1)

    src = wb[src_name]

    if dst_name in wb.sheetnames:
        dst = wb[dst_name]
    else:
        dst = wb.create_sheet(dst_name)

    print(f"Copying values from '{src_name}' to '{dst_name}'...")
    copy_sheet_values(src, dst)

    wb.save(xlsx)
    print("Saved workbook with updated Basketlist.")
