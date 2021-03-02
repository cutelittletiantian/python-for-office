import openpyxl

# 读文件、读表
workFile = openpyxl.load_workbook("./resources/ex3_2/进楼登记.xlsx")
wb = workFile["体温表"]

feverList = []

# 遍历行数据
for rowData in wb.rows:
    name = rowData[0].value
    # 跳过表头
    if name == "姓名":
        continue
    temperature = rowData[1].value
    # 判断
    if float(temperature) > 37:
        feverList.append(name)

print(feverList)