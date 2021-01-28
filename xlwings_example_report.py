# -*- coding: utf-8 -*-
"""
Created on Thu Jan  7 15:03:58 2021

"""

import pandas as pd
import xlwings as xw

# Import CSV file using one of the two methods below
#df = pd.read_csv(r"path_to_csv\fruit_and_veg_sales.csv")
df = pd.read_csv(r"https://raw.githubusercontent.com/Nishan-Pradhan/xlwings_report/master/fruit_and_veg_sales.csv")

# Open new Excel Workbook
wb = xw.Book()

# Define Sheet and change Sheet Name
sht = wb.sheets["Sheet1"]
sht.name = "fruit_and_veg_sales"

# DataFrame to cell A1 in Excel
sht.range("A1").options(index=False).value = df

# Select all data range
all_data_range = sht.range("A1").expand("table")

# Set Row height and Column width
all_data_range.row_height = 22.5
all_data_range.column_width = 12

# Format colors, font, alignment and wrap text
all_data_range.color = (208,206,206)
all_data_range.api.Font.Name = "Arial"
all_data_range.api.Font.Size = 8
all_data_range.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
all_data_range.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
all_data_range.api.WrapText = True

# Format headers
header_range = sht.range("A1").expand("right")
header_range.color = (112,173,71)
header_range.api.Font.Color = 0xFFFFFF
header_range.api.Font.Bold = True
header_range.api.Font.Size = 9

# Format first column
id_column_range = sht.range("A2").expand("down")
id_column_range.color=(198,224,180)

# Add borders only around data
data_ex_headers_range = sht.range("A2").expand("table")
for border_id in range(7,13):
    data_ex_headers_range.api.Borders(border_id).Weight = 2
    data_ex_headers_range.api.Borders(border_id).Color = 0xFFFFFF

# Bonus - change tab colour
sht.api.Tab.Color = 0x70AD47

# Save your Excel file
wb.save(r"\folder_path\fruit_and_veg_report.xlsx")
