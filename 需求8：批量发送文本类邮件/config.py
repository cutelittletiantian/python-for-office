import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# 这个文件夹里面，配置一些私密的数据信息，例如发送方（自己）的用户名、密码
# 注意：所谓“密码”，不是真正的qq邮箱密码，而是由qq邮箱SMTP服务生成的授权码
sender_info = {
    "nickname": "指定自己发送邮件的发件人名",
    "email": "请在此填写补全自己发送端的邮箱",
    "password": "请申请到授权码，并填写于此"
}

# 存储邮箱联系方式的资源路径
resourcesPath = "./resources"
# 存储邮箱联系方式的文件名
contactBookName = "邮件联系方式表.xlsx"
# 存储邮箱联系方式的工作表名
contactSheetName = "邮箱"


# 定义函数，一次性读出excel中所有的联系人信息
def read_excel(path="./", sheet_name="") -> list:
    # 打开、选表
    contact_sheet = openpyxl.load_workbook(filename=path)[sheet_name]  # type: Worksheet
    # 取数据
    contact_list = [contact_item for contact_item in contact_sheet.iter_rows(min_row=2, values_only=True)]
    return contact_list


# 测试函数read_excel能否正常输出
# import os
# contactList = read_excel(path=os.path.join(resourcesPath, contactBookName), sheet_name=contactSheetName)
# print(contactList)
