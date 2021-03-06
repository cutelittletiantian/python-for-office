import os
import openpyxl
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

orderPath = "./resources/ex4_4/Michael/doc"
# 主要操作的表预先存着
mainSheet = "销售订单数据"

# 遍历各表文件
orderItems = os.listdir(path=orderPath)

for orderItem in orderItems:
    # 完整路径
    orderItemPath = os.path.join(orderPath, orderItem)

    # 读表
    wb = openpyxl.load_workbook(filename=orderItemPath, data_only=True) # type: Workbook
    # 加载主要操作的表
    orderSheet = wb[mainSheet]  # type: Worksheet

    # 遍历主表各行
    for rowData in orderSheet.iter_rows(min_row=2):
        # 获取当前行的完整内容
        rowContent = [something.value for something in rowData]
        # 获取其中的商品名
        productName = rowContent[2]
        # 转到相应的表中添加数据
        wb[productName].append(rowContent)

    # 保存
    wb.save(filename=orderItemPath)