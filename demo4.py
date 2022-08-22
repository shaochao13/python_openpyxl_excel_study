from faker import Faker
import openpyxl
from openpyxl.chart import (
    BarChart, Reference
)


def init_test_data():
    faker = Faker(locale='zh_CN')
    data = [
        ('月份', faker.name(), faker.name(), faker.name())
    ]

    for i in range(1, 13):
        data.append((f'{i}月', faker.random_number(digits=5), faker.random_number(digits=5), faker.random_number(digits=5)))

    return data


def create_bar_chart(file_path):
    wb = openpyxl.Workbook()
    ws = wb.active

    datas = init_test_data()
    for data in datas:
        ws.append(data)

    data = Reference(ws, min_col=2, max_col=4, min_row=1, max_row=13)
    labels = Reference(ws, min_col=1, min_row=2, max_row=13)

    bar_chart = BarChart()
    bar_chart.title = '销售人员业绩表(2020年)'
    bar_chart.x_axis.title = '月份'
    bar_chart.y_axis.title = '销售额(万元)'

    bar_chart.add_data(data, titles_from_data=True)
    bar_chart.set_categories(labels)
    bar_chart.width = 34
    bar_chart.height = 15
    ws.add_chart(bar_chart, 'F2')

    wb.save(file_path)


if __name__ == '__main__':
    create_bar_chart('./bar_chart.xlsx')