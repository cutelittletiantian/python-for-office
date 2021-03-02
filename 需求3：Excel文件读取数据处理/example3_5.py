# 总金额
import openpyxl

# 文档的路径
import rmbTrans
from openpyxl.utils import cell
from openpyxl.worksheet.worksheet import Worksheet

productFile = "./resources/ex3_5/开口哭牌供货列表.xlsx"

totalMoney = 0

# 读文件、读表
productSheet: Worksheet = openpyxl.load_workbook(filename=productFile)["6月供货"]

# 遍历行数据，跳过表头
for rowData in productSheet.iter_rows(min_row=2):
    # 读总金额
    sTotal: str = rowData[
        cell.column_index_from_string("G")-1
    ].value
    # 转数字(用rmbTrans第三方库，将人民币大写汉字转数值)
    dTotal: float = rmbTrans.trans(sTotal)
    totalMoney += dTotal

print(f"开口哭牌供应商6月采购总金额为{totalMoney}元。")