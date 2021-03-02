import os
import shutil
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import cell
from openpyxl.worksheet.worksheet import Worksheet

os.chdir("./resources/challenge/Michael/")
studentInfoPath = "学生资料"
studentRegionPath = "学生地区.xlsx"


# 学生:地区
student_region = dict()
# 提取excel中信息
regionSheet: Worksheet = openpyxl.load_workbook(filename=studentRegionPath)["地区表"]
# 遍历行数据
for rowData in regionSheet.iter_rows(min_row=2):
    # 姓名
    name = rowData[0].value
    # 地区
    region = rowData[1].value
    # 添加键值对
    student_region[name] = region


studentInfoItems = os.listdir(studentInfoPath)
for studentInfo in studentInfoItems:
    '''
    :param studentInfo: 文件名（格式：姓名.docx）
    '''
    name = os.path.splitext(studentInfo)[0]
    # 对应地区
    region = student_region.get(name, None)

    # 非空前提下
    if region is None:
        continue
    else:
        regionPath = os.path.join(studentInfoPath, region)
        if not os.path.exists(regionPath):
            os.mkdir(regionPath)

        # 移动文件
        sourcePath = os.path.join(studentInfoPath, studentInfo)
        shutil.move(sourcePath, regionPath)