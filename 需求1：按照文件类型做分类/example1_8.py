import os

projectPath = "./resources/ex1_8/project"

# 目录下的所有文件夹
allStudents = os.listdir(projectPath)

# 待交付物
assignments = ["开题报告", "中期报告", "任务书", "指导记录表"]

# 遍历所有的文件夹
for student in allStudents:
    # 组装该文件夹的完整路径
    studentPath = os.path.join(projectPath, student)
    # 该学生所有的交付文件
    allHandouts = os.listdir(studentPath)

    # 对于每一种交付件
    for assignment in assignments:
        # 交付件名组装
        shouldHandout = f"{assignment}_{student}.docx"
        if shouldHandout not in allHandouts:
            print(f"{student}没有交{assignment}")