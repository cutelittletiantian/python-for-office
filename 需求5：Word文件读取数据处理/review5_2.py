import os
import shutil

regPath = "./resources/review5_2/registration"

allForms = os.listdir(path=regPath)
print(allForms)

for form in allForms:
    # 完整路径
    formPath = os.path.join(regPath, form)
    if os.path.isdir(formPath):
        continue
    # 提取文件名
    formName = os.path.splitext(form)[0]
    # 提取相应科目
    formSubject = formName.split(sep="-")[1]
    # 目标路径
    targetPath = os.path.join(regPath, formSubject)
    if not os.path.exists(path=targetPath):
        os.mkdir(path=targetPath)
    # 移动
    shutil.move(src=formPath, dst=targetPath)

print(os.listdir(regPath))