import openpyxl
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

gatherPath = "./resources/ex4_5/data/汇总.xlsx"

# 读文件
productBook = openpyxl.load_workbook(filename=gatherPath)  # type: Workbook
# 汇总表
summarySheet = productBook["口香糖"]  # type: Worksheet


# 处理数据
def process_platform_sheet(platform_code: str):
    platform_sheet = productBook[f"{platform_code}平台"]
    for rowData in platform_sheet.iter_rows(min_row=2):
        # 找到口香糖产品
        if "口香糖" in rowData[0].value:
            # 复制下来一行的内容
            row_content = [content.value for content in rowData]
            # 末尾再贴上平台
            row_content.append(f"{platform_code}平台")
            # 添加到汇总表中
            summarySheet.append(row_content)


# 依次处理ABC三个平台
for platformCode in ["A", "B", "C"]:
    process_platform_sheet(platform_code=platformCode)

# 保存并关闭文件
productBook.save(filename=gatherPath)
productBook.close()