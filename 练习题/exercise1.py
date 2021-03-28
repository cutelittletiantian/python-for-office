import os
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet



# 工作区确定
os.chdir(path="./resources/exer_1")

# 作者投稿登记表文件夹所在路径
registerPath = "sub"

# 读取["18年", "19年", "20年"]的数据，（可指定查找年份）
years = ["18年", "19年", "20年"]
for year in years:
    # 组装相应年表格文件的路径
    registerFileName = f"作者投稿-{year}.xlsx"
    registerFilePath = os.path.join(registerPath, registerFileName)

    # 打开相应工作簿
    registerBook = openpyxl.load_workbook(filename=registerFilePath)  # type: Workbook
    # 逐个扫描每一个工作表
    for registerSheet in registerBook.worksheets:
        # 逐行遍历，跳过表头
        for rowData in registerSheet.iter_rows(min_row=2):
            if rowData[1].value == "小夜"\
                    and rowData[0].value == "索引" \
                    and rowData[3].value == "16798429@yequ.com":
                # 输出查询结果
                print(registerFileName, rowData)