import os
import shutil

# 报告所在的路径
projectPath = "./resources/ex1_6"

# 获取文件夹下的文件
studentFiles = os.listdir(projectPath)

for studentFile in studentFiles:
    # 组装完整路径
    studentFilePath = os.path.join(projectPath, studentFile)

    # 判断这个路径是不是文件夹
    if os.path.isdir(studentFilePath):
        continue

    # 分离出文件名
    studentFileName = os.path.splitext(studentFile)[0].lower()

    # 分离学生姓名
    student = studentFileName.split("_")[1].strip()

    # 组装目标路径
    targetPath = os.path.join(projectPath, student)

    # 判断目标路径如果不存在要创建
    if not os.path.exists(targetPath):
        os.mkdir(targetPath)

    # 移动文件
    shutil.move(studentFilePath, targetPath)