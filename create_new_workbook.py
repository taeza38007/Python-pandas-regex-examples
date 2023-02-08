#!/usr/local/bin/python3
from openpyxl import Workbook, load_workbook

# Assign workbook to 'wb' variable
wb = Workbook()
# Set worksheet
ws = wb.active
# Set title
ws.title = 'Data'

# Append list of python to the worksheet
ws.append(['Tim', 'Is', 'Great', '!'])
ws.append(['Tim', 'Is', 'Great', '!'])
ws.append(['end'])

# Save appended info 
wb.save('tim_example.xlsx')