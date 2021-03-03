import openpyxl
import os

from openpyxl import Workbook

os.chdir(path="./resources/ex4_2/Michael/")

fromPath = "doc"
toPath = "revise"

if not os.path.exists(path=toPath):
    os.mkdir(path=toPath)

orderDocs = os.listdir(fromPath)
for orderDoc in orderDocs:
    orderDocPath = os.path.join(fromPath, orderDoc)
    orderBook: Workbook = openpyxl.load_workbook(filename=orderDocPath)

    for orderSheet in orderBook.worksheets:
        # 继续遍历每一行中的内容
        for rowData in orderSheet.iter_rows(min_row=2):
            if rowData[2].value == "有点酸可乐":
                rowData[2].value = "有点酸甜可乐"

    # 保存
    targetPath = os.path.join(toPath, orderDoc)
    orderBook.save(filename=targetPath)