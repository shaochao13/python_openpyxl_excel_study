import openpyxl
from openpyxl.chart import LineChart, BarChart, Reference


# 将测试数据插入到工作表中
def insert_test_data_to_sheet(ws):
    rows = [
        ['月份', '木材', '糖'],
        ['1月', 10000, 10],
        ['2月', 15002, 30],
        ['3月', 30000, 50],
        ['4月', 9500, 35],
        ['5月', 21000, 25],
        ['6月', 12352, 45],
        ['7月', 32009, 90],
    ]
    for r in rows:
        ws.append(r)


# 绘制拆线图
def create_line_chart(ws):
    # 创建一个拆线图LineChart对象
    line = LineChart()

    # 创建数据选区， 木材数据用拆线图显示
    data = Reference(ws, min_col=2, min_row=1, max_row=7)
    # 将数据添加到图形上
    line.add_data(data, titles_from_data=True)

    # 创建横轴labels 选区
    labels = Reference(ws, min_col=1, min_row=2, max_row=7)
    # 将labels添加到图形上
    line.set_categories(labels)

    # 让拆线图的y轴显示刻度位于图形的右侧  'autoZero', 'max', 'min'
    line.y_axis.crosses = 'max'
    # 设置拆线图y 轴显示的title
    line.y_axis.title = '木材产量(棵)'


    return line


# 绘制柱形图
def create_bar_chart(ws):
    # 创建一个柱形图BarChart对象
    bar = BarChart()

    # 创建数据选区， 糖 数据用拆线图显示
    data = Reference(ws, min_col=3, min_row=1, max_row=7)
    # 将数据添加到图形上
    bar.add_data(data, titles_from_data=True)

    # # 创建横轴labels 选区
    # labels = Reference(ws, min_col=1, min_row=2, max_row=7)
    # # 将labels添加到图形上
    # bar.set_categories(labels)

    # 显示第二个y轴，如果不设置，将不会显示
    bar.y_axis.axId = 200
    # 设置网络线不显示
    bar.y_axis.majorGridlines = None
    # 设置柱形图 y 轴显示的title
    bar.y_axis.title = '糖产量(万吨)'

    return bar


def create_line_bar_chart(file_path):
    wb = openpyxl.Workbook()
    ws = wb.active

    insert_test_data_to_sheet(ws)

    line_chart = create_line_chart(ws)
    bar_chart = create_bar_chart(ws)

    # 让两个图形叠加在一起
    line_chart += bar_chart

    # 设置图形的宽、高
    line_chart.width = 30
    line_chart.height = 15

    # 设置图形的title
    line_chart.title = 'Demo'
    # 设置图形的横轴title
    line_chart.x_axis.title = '月份'

    # 将叠加后的图形添加到工作表中
    ws.add_chart(line_chart, anchor= 'E2')

    # 保存workbook
    wb.save(file_path)


if __name__ == "__main__":
    create_line_bar_chart('./pExcel07.xlsx')