import os
import openpyxl
from openpyxl.utils import cell
from openpyxl.worksheet.worksheet import Worksheet

# 全班同学的姓名列表
studentName = [
    '吕婷婷', '肖槐', '魏皖怡', '胡轶轩', '包印雪',
    '谭彦', '周宇', '吴琪', '龚静雯', '张思思',
    '潘婷', '夏乐群', '朱佩玉', '隋胜男', '朱薇',
    '唐梅', '罗勇', '林贸然', '张丽', '张可馨',
    '汪洋', '韩明希', '杜宇琳', '胡连群', '岳海洋',
    '李雅梦', '蔡林芮', '孟轩', '项文彦', '苏培坤',
    '惠红', '顾洪彬', '卡欣', '陶振英', '子顺',
    '成洛', '邹德浩', '王思亮', '叶家有', '王德滋',
    '杨少帆', '谢福顺', '刘军', '李多钰'
]

allFormsPath = "./resources/review3/回执/"

# 找出提交名单
submitStudent = []
allItems = os.listdir(allFormsPath)
for item in allItems:
    submitStudent.append(
        os.path.splitext(item)[0].split(sep="-")[1]
    )

# 创建表
submitBook = openpyxl.Workbook()
# 访问工作表
submitSheet: Worksheet = submitBook["Sheet"]
submitSheet.title = "七年级六班"
# 遍历全班学生
for index, student in enumerate(studentName, start=1):
    submitSheet[f"A{index}"].value = student
    if student not in submitStudent:
        submitSheet[f"B{index}"].value = "x"

# 最后一步千万别忘了，保存文件
submitBook.save(filename=os.path.join(allFormsPath, "夜曲中学七年级六班回执提交情况.xlsx"))