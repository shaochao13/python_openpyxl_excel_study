import openpyxl
from openpyxl.chart import AreaChart, Reference

def create_chart(file_path):
    wb = openpyxl.Workbook()
    ws = wb.active

    rows=[
      ['Number','Batch 1', 'Batch 2'],
      [2, 40, 30],
      [3, 40, 25],
      [4, 50, 30],
      [5, 30, 10],
      [6, 25, 5],
      [7, 30, 40]]

    for row in rows:
        ws.append(row)

    chart = AreaChart()
    chart.style = 10

    data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=7)
    labels = Reference(ws, min_col=1, min_row=2, max_row=7)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    ws.add_chart(chart, anchor='D2')

    wb.save(file_path)


if __name__ == "__main__":
    create_chart('./demo06.xlsx')