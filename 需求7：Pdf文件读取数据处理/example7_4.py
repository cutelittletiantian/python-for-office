import os
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import pdfplumber

# 工作路径
os.chdir(path="./resources/ex7_4/")
# 各大银行所在路径
bankListPath = "bankBalance"
# 工作簿路径
bankBookPath = "年底合集.xlsx"

# 新建工作表
bankBook = openpyxl.Workbook()  # type: Workbook
# 删除默认工作表
bankBook.remove(worksheet=bankBook["Sheet"])

# 提取银行列表
bankList = os.listdir(path=bankListPath)
for bankItem in bankList:
    # 组装完整路径
    bankPath = os.path.join(bankListPath, bankItem)
    # 分离的文件名
    bankName = os.path.splitext(p=bankItem)[0]
    # 根据文件名，创建工作表
    bankSheet = bankBook.create_sheet(title=bankName)  # type: Worksheet

    # 打开pdf文档
    bankFile = pdfplumber.open(path_or_fp=bankPath)
    # 获取表格所在页
    page = bankFile.pages[0]
    # 提取表格
    table = page.extract_tables()[0]
    # 打印看看里面有些什么？
    # print(table)
    # Hmm，原来其中的内容是pdf表中的一行行内容构成的列表，既然这样那就知道怎么写excel表了
    for rowData in table:
        bankSheet.append(rowData)
    # 保存一波工作表
    bankBook.save(filename=bankBookPath)
