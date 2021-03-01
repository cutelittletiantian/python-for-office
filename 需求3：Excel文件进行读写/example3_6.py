import openpyxl

phonePath = "./resources/ex3_6/电话本.xlsx"


def queryPhone(queryName) -> list:
    # 结果是一个列表
    resultPhone = []

    # 读文件、读表
    querySheet = openpyxl.load_workbook(phonePath)["电话簿"]

    # 遍历行（跳过line1）
    for rowData in querySheet.iter_rows(min_row=2):
        name = rowData[0].value
        if name == queryName:
            resultPhone.append(rowData[1].value)

    return resultPhone


# 测试用例举例
print(queryPhone("周倩文"))
