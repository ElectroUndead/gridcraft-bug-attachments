"""Create a minimal English-language reproduction workbook for the GridCraft header-sort bug.

Layout mirrors the real file: row 1 = summary line, row 2 = header row (AutoFilter range
starts at row 2), frozen panes ySplit=2, data from row 3.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

path = r"C:\Users\ElectroUndead\Documents\deepseek-harness\Main\gridcraft-bug\repro\GridCraft-header-sort-sample.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Payments"

ws["A1"] = "Payments from the main balance 01.10.2026-13.10.2026: 6 rows, 2,100.00 USD"

header = ["Date", "Company", "Amount"]
rows = [
    ["2026-09-30", "Zeta Cargo", 100],
    ["2026-10-01", "Alpha Logistics", 200],
    ["2026-10-02", "Gamma Shipping", 300],
    ["2026-10-03", "Beta Freight", 400],
    ["2026-10-04", "Delta Trans", 500],
    ["2026-10-05", "Epsilon Ltd", 600],
]

for c, v in enumerate(header, start=1):
    cell = ws.cell(row=2, column=c, value=v)
    # Formatting matters: this is what makes GridCraft recognise row 2 as the header row.
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor="DDEBF7")
for i, r in enumerate(rows, start=3):
    for c, v in enumerate(r, start=1):
        ws.cell(row=i, column=c, value=v)

last = 2 + len(rows)                      # row 8
ws.auto_filter.ref = f"A2:C{last}"        # filter range INCLUDES the header row (as in the real file)
ws.freeze_panes = "A3"                    # -> <pane ySplit="2" topLeftCell="A3" state="frozen"/>
ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 10

wb.save(path)
print("wrote", path)
