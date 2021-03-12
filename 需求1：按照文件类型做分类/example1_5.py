import os
import shutil

# 所有答题卡路径
examPath = "./resources/ex1_5/exam"

# 试卷列表
allPapers = os.listdir(examPath)

for paper in allPapers:
    # 当前试卷的完整路径
    paper_path = os.path.join(examPath, paper)
    # 判断这个路径是文件夹还是文件
    if os.path.isdir(paper_path):
        continue

    # 当前试卷的文件名
    paper_filename = os.path.splitext(paper)[0].lower()
    # 当前试卷所在的班级
    paper_class = paper_filename.split("-")[0].strip()

    # 试卷将要存储的完整路径
    targetPath = os.path.join(examPath, paper_class)

    # 判断一下这个路径是否已经存在
    if not os.path.exists(targetPath):
        os.mkdir(targetPath)

    # 移动当前试卷
    shutil.move(paper_path, targetPath)