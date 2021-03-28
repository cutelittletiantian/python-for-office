import PIL
from PIL import Image
import os


# 设置工作区
os.chdir(path="./resources/exer_4")
# 待修改文件所在路径
beforePath = "before"
# 修改后文件所在路径
afterPath = "after"
if not os.path.exists(path=afterPath):
    os.mkdir(path=afterPath)

# 获取待剪裁照片的列表
photoList = os.listdir(path=beforePath)
# 遍历、改尺寸
for photo in photoList:
    # 组装完整路径
    photoPath = os.path.join(beforePath, photo)
    # 读取照片
    imgBefore = PIL.Image.open(fp=photoPath)  # type: Image.Image
    # 两寸（413*626px）照片修改到一寸（295*413px）
    photoSize = (295, 413)
    imgAfter = imgBefore.resize(size=photoSize)
    # 组装完整目标路径
    targetPath = os.path.join(afterPath, photo)
    # 保存
    imgAfter.save(fp=targetPath)