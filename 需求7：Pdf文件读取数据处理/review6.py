# review4 用非函数方法实现
import os

docPath = "./resources/review6"

allItems = os.listdir(docPath)


chapterList = []

for item in allItems:
    itemName = os.path.split(p=item)[0]
    chapterCode = itemName.split(sep="-")[1]
    chapterList.append(chapterCode)

# 执行排序
chapterList.sort(key=int)

# 空白列表
file = []
# 还原文件名
for chapter in chapterList:
    for item in allItems:
        newChapter = str(chapter)
        sourceName = item.split(sep="-")[1]
        newName = sourceName.split(sep=".")[0]

        if newChapter == newName:
            file.append(item)

print(file)