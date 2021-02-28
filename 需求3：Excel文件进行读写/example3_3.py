import os
import shutil
import openpyxl

booksPath = "./resources/ex3_3/books"
csPath = "./resources/ex3_3/计算机科学"
excelPath = "./resources/ex3_3/长颈鹿图书馆.xlsx"

if not os.path.exists(csPath):
    os.mkdir(csPath)

# 找excel
ws = openpyxl.load_workbook(excelPath)["长颈鹿图书馆书籍清单"]

csList = []

for rowData in ws.rows:
    bookName = rowData[1].value
    bookType = rowData[4].value

    if bookType == "计算机科学":
        csList.append(bookName)

bookItems = os.listdir(booksPath)
for book in csList:
    bookFileName = f"{book}.docx"
    itemPath = os.path.join(booksPath, bookFileName)
    shutil.move(itemPath, csPath)