# openpyxl
pip install openpyxl

Installing from local archives (Windows)
py -m pip install ./downloads/SomeProject-1.0.4.tar.gz


https://www.youtube.com/watch?v=5_cR4cwrz8E


from openpyxl import Workbook
wb = Workbook()

# grab the active worksheet
ws = wb.active

# Data can be assigned directly to cells
ws['A1'] = 42

# Rows can also be appended
ws.append([1, 2, 3])

# Python types will automatically be converted
import datetime
ws['A2'] = datetime.datetime.now()

# Save the file
wb.save("sample.xlsx")
