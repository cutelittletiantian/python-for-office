import os

import openpyxl
from openpyxl.utils import cell
from openpyxl.worksheet.worksheet import Worksheet

resourcePath = "./resources/ex3_7/doc/"
os.chdir(resourcePath)


def profit_analysis(excelFile, month, monthlyCost):
    # 读表（算出公式结果）
    orderSheet: Worksheet = openpyxl.load_workbook(excelFile, data_only=True)["销售订单数据"]

    # 当月收入
    total_amount = 0

    # 遍历行（跳过表头）
    for rowData in orderSheet.iter_rows(min_row=2):
        # 读一行的总收入
        amount = rowData[
            cell.column_index_from_string("I") - 1
            ].value
        # 将总的进行加和
        total_amount += amount

    # 判断盈利与亏损
    if total_amount > monthlyCost:
        print(f"{month}月是盈利的。盈利{total_amount - monthlyCost}元。")


# 处理12个文件（range是左开右闭区间）
for month in range(1, 13):
    excelFileName = f"2019年{month}月销售订单.xlsx"
    profit_analysis(excelFile=excelFileName, month=month, monthlyCost=8500)
