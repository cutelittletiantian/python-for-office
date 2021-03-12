import os
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

workPath = "./resources/review6_1/word_write"
os.chdir(path=workPath)

# 表格成绩
scoreTablePath = "夜曲大学英语考试成绩.xlsx"

# 读表
scoreBook = openpyxl.load_workbook(scoreTablePath)  # type: Workbook
scoreSheet = scoreBook["汇总"]  # type: Worksheet

# 遍历行
for rowData in scoreSheet.iter_rows(min_row=2):
    # 学院
    rowData[2].value = rowData[2].value.replace("计算机学院", "计算机科学与技术学院")
    rowData[2].value = rowData[2].value.replace("机械科学与技术学院", "机械制造科学与技术学院")

# 保存
scoreBook.save(filename="新-夜曲大学英语考试成绩.xlsx")