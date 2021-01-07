# -*- coding: utf-8 -*-
"""
Created on Thu Jan  7 15:03:58 2021

"""

import pandas as pd
import xlwings as xw

df = pd.read_csv(r"C:\Users\Nishan\Documents\Medium\fruit_and_veg_sales.csv")


wb = xw.Book()

sht = wb.sheets["Sheet1"]
sht.name = "fruit_and_veg_sales"

sht.range("A1").options(index=False).value = df


all_data_range = sht.range("A1").expand('table')

all_data_range.row_height = 22.5
all_data_range.column_width = 12

all_data_range.color = (208,206,206)
all_data_range.api.Font.Name = 'Arial'
all_data_range.api.Font.Size = 8
all_data_range.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
all_data_range.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
all_data_range.api.WrapText = True

header_range = sht.range("A1").expand('right')
header_range.color = (112,173,71)
header_range.api.Font.Color = 0xFFFFFF
header_range.api.Font.Bold = True
header_range.api.Font.Size = 9

id_column_range = sht.range("A2").expand('down')
id_column_range.color=(198,224,180)

data_ex_headers_range = sht.range("A2").expand('table')
for border_id in range(7,13):
    data_ex_headers_range.api.Borders(border_id).Weight = 2
    data_ex_headers_range.api.Borders(border_id).Color = 0xFFFFFF

wb.save(r"C:\Users\Nishan\Documents\Medium\fruit_and_veg_report.xlsx")







