import openpyxl
from openpyxl.chart import (
    BarChart, Series, Reference
)

wb = openpyxl.Workbook()
sheet = wb.active

rows = [
    ('Number', 'Batch 1', 'Batch 2'),
    (2, 10, 30),
    (3, 40, 60),
    (4, 50, 70),
    (5, 20, 10),
    (6, 10, 40),
    (7, 50, 30),
]


for row in rows:
    sheet.append(row)

refObj = Reference(sheet, min_col=2, max_col=3, min_row=1, max_row=7)
cats = Reference(sheet, min_col=1, min_row=2, max_row=7)
#
chartObj = BarChart()
chartObj.type = 'bar' # 横向bar
# chartObj.type = 'col'  # 竖向bar
chartObj.style = 25
chartObj.title = 'DEMO'
chartObj.y_axis.title = 'Number'
chartObj.x_axis.title = 'Sample'
chartObj.add_data(refObj, titles_from_data=True)
chartObj.set_categories(cats)
#
sheet.add_chart(chartObj, 'A10')

wb.save('./chart.xlsx')