import os
import shutil

regPath = "./resources/review5_1/registration"
targetPath = "./resources/review5_1/不合格"

# 可竞选科目
subjects = ["语文", "数学", "英语", "体育", "音乐", "安全教育", "美术"]

# 可竞选年级
grades = ["一年级", "二年级", "三年级", "四年级", "五年级", "六年级"]

if not os.path.exists(targetPath):
    os.mkdir(path=targetPath)

regItems = os.listdir(path=regPath)
# 输出报名表的总数量
print(len(regItems))

for regItem in regItems:
    regItemName = os.path.splitext(regItem)[0]
    regInfo: list = regItemName.split(sep="-")
    if len(regInfo) == 3 and (regInfo[1] in subjects) and (regInfo[2] in grades):
        continue
    else:
        # 完整路径
        regItemPath = os.path.join(regPath, regItem)
        # 完整目标路径
        itemTargetPath = os.path.join(targetPath, regItem)
        # 移动
        shutil.move(src=regItemPath, dst=itemTargetPath)

# 输出不合格数量
print(len(os.listdir(regPath)))