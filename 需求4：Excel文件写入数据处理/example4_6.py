import ezexif
import openpyxl
import os

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

os.chdir(path="./resources/ex4_6/")
photoPathList = "photo"
recordPath = "照片参数.xlsx"

# 取照片
allPhotos = os.listdir(path=photoPathList)
# 开工作表格
photoParamBook = openpyxl.load_workbook(filename=recordPath)  # type: Workbook
photoParamSheet = photoParamBook["示例"]  # type: Worksheet

# 遍历照片
for photo in allPhotos:
    # 组装路径
    photoPath = os.path.join(photoPathList, photo)
    # 加载照片属性
    photoProperties = ezexif.process_file(filename=photoPath)
    # 如果你忘了有哪些键，不妨先输出，查一下
    # print(photoProperties)
    # 组装一行数据
    rowContent = (
        photo,
        photoProperties["Image ExifOffset"],
        photoProperties["EXIF ExposureProgram"],
        photoProperties["EXIF DateTimeOriginal"],
        photoProperties["EXIF MeteringMode"],
        photoProperties["EXIF Flash"],
        photoProperties["EXIF ExposureMode"]
    )
    # 工作表添加数据
    photoParamSheet.append(iterable=rowContent)


# 另存
photoParamBook.save(filename="照片参数-新.xlsx")