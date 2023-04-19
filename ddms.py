import csv
import datetime
from openpyxl import Workbook
from openpyxl.utils.exceptions import IllegalCharacterError
from openpyxl.styles import PatternFill, Font, Alignment

def remove_illegal_chars(cell_value):
    return ''.join(c for c in cell_value if c.isprintable())

now = datetime.datetime.now()
date_string = now.strftime('%Y-%m-%d')
file_name = f'Pick_List_{date_string}.xlsx'

# Read the data from the text file
data = []
with open('Reports.txt', 'r') as f:
    reader = csv.reader(f, delimiter='\t')
    for i, row in enumerate(reader):
        if i < 2:
            continue
        new_row = []
        for cell in row:
            new_row.extend(cell.split('|'))
        data.append(new_row)

# Create a new Excel workbook and add the data
wb = Workbook()
ws = wb.active

# Set the headers
headers = data.pop(0)
header_font = Font(bold=True)
header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
header_alignment = Alignment(horizontal="center", vertical="center")

for col_num, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_num, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment

# Add the remaining data to the worksheet
for row_num, row in enumerate(data, start=2):
    cleaned_row = [remove_illegal_chars(cell) for cell in row]
    for col_num, cell_value in enumerate(cleaned_row, start=1):
        ws.cell(row=row_num, column=col_num, value=cell_value)

# Save the workbook as an .xlsx file
ws.delete_cols(2, 1)
ws.delete_cols(3, 1)
ws.delete_cols(16, 1)
ws.delete_cols(20, 1)
ws.delete_cols(21, 1)
ws.delete_cols(22, 1)
ws.delete_cols(23, 1)
ws.delete_cols(27, 1)
ws.delete_cols(28, 1)
ws.delete_cols(29, 1)
ws.delete_cols(30, 1)
ws.delete_cols(31, 1)
ws.delete_cols(32, 1)
ws.delete_cols(33, 1)
wb.save(file_name)
