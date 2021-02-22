import os
searchPath = "./resources/"
searchFileName = "杜子鄂实习合同.docx"
allItems = os.listdir(searchPath)
fileExists = False  # 表示是否存在
for item in allItems:
    if item == searchFileName:
        fileExists = True
        break

# 查看结果（非核心代码）
if fileExists:
    print("文件找到")
else:
    print("文件未找到")