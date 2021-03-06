import openpyxl
import os

path = "./resources/ex3_1/"

allItems = os.listdir(path)

for item in allItems:
    itemPath = os.path.join(path, item)
    wb = openpyxl.load_workbook(itemPath)
    if "十二月销售订单数据" in wb.sheetnames:
        print(item)