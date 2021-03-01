import openpyxl
from openpyxl import Workbook
from openpyxl.utils import cell
from openpyxl.worksheet.worksheet import Worksheet

excelPath = "./resources/ex3_4/2019年1月销售订单.xlsx"

# 读文件（需要函数）
wb: Workbook = openpyxl.load_workbook(excelPath, data_only=True)
# 读表
orderSheet: Worksheet = wb["销售订单数据"]

# 商品及对应销售额
productSale = dict()

# 遍历行
for dataRow in orderSheet.iter_rows(min_row=2):
    # 名字
    productName = dataRow[
        cell.column_index_from_string("C") - 1
    ].value
    # 总价
    productTotal = int(dataRow[
        cell.column_index_from_string("I") - 1
    ].value)
    productSale[productName] = productSale.get(productName, 0) + productTotal

maxPrice = 0
maxName = ""

for k, v in productSale.items():
    if v > maxPrice:
        maxPrice = v
        maxName = k

print(f"明星产品是：{maxName}，一月总共销售出{maxPrice}元")