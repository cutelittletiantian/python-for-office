import os
import shutil

# 所有文件路径
sourcePath = "./resources/ex1_7/source"

# 目标路径
targetPath = "./resources/ex1_7/animal"
# 确保创建
if not os.path.exists(targetPath):
    os.mkdir(targetPath)

# 待查文件列表
searchList = ['东北虎.jpg', '非洲最美猎豹.jpg', '非洲最美长颈鹿.jpg', '几维鸟.jpg']

allSources = os.listdir(sourcePath)

for source in allSources:
    source = source.lower()
    if source in searchList:
        # 组装完整源文件路径
        filePath = os.path.join(sourcePath, source)
        shutil.move(filePath, targetPath)