import os
import openpyxl
import shutil
from openpyxl import Workbook
from openpyxl.utils import cell
from openpyxl.worksheet.worksheet import Worksheet

# 设置工作区
os.chdir(path="./resources/exer_5")

# 待整理文件所在路径
teacherPath = "teacher"
# 汇总工作簿路径
summaryPath = "汇总.xlsx"
# 汇总工作表名
summarySheetName = "面试教师名单"


################################################
def make_directory(path: str = None):
    """
    创建一个文件夹path，如果存在则无动作
    :param path: 需要创建的文件路径，需要提前指定
    :return: None
    """
    if not os.path.exists(path=path):
        os.mkdir(path=path)

    return
################################################


# 从excel中清理出需要整理的文件信息
# 文件夹名所在的列表
targetFileNameList = []
# 读表
selectedSheet = openpyxl.load_workbook(filename=summaryPath)[summarySheetName]  # type: Worksheet
# 扫描行
for rowData in selectedSheet.iter_rows(min_row=2, values_only=True):
    # 组装文件名（为了让大家看清楚传入的参数都是个什么结构，我把这里的括号分行写的）
    fileName = "-".join(
        [
            rowData[0],
            rowData[cell.column_index_from_string("I") - 1]
        ]
    )
    targetFileNameList.append(fileName.lower())


# 分类teacher文件夹下的内容
# 获取所有文件构成列表
teacherItemList = os.listdir(path=teacherPath)
# 逐个遍历文件
for teacherItem in teacherItemList:
    # 组装完整路径
    teacherItemPath = os.path.join(teacherPath, teacherItem)

    # 分离出文件名
    teacherItemName = os.path.splitext(p=teacherItem)[0].lower()
    # 不在需要整理的列表当中的文件不予分类
    if teacherItemName not in targetFileNameList:
        continue

    # 执行分类
    # 分离出职称（辅导员、讲师、研究院之类的）
    job = teacherItemName.split("-")[1]
    # 组装目标路径
    targetPath = os.path.join(teacherPath, job)
    # 创建相应的文件夹（自定义函数判断是否存在）
    make_directory(path=targetPath)
    # 执行移动
    shutil.move(src=teacherItemPath, dst=targetPath)