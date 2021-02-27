import os

searchPath = "./resources/ex1_2"
allItems = os.listdir(searchPath)

cnt_vid = 0

for item in allItems:
    # 下面这里其实做了3件事情
    # 第一件事情，os.path.splitext()中传入item（文件名是【杜子鄂的个人简历.DOCX】）
    # 这时，返回的结果是列表：["杜子鄂的个人简历", ".DOCX"]
    # 第二件事情，[1]表示选择上面列表中，下标为1的元素，即选出".DOCX"
    # 第三件事情，lower()函数将".DOCX"中的大写英文字母统一转为小写
    # 综上所述，这时的extension赋值的结果为".docx"，字母是小写的嗷~
    extension = os.path.splitext(item)[1].lower()
    if extension in [".avi", ".mp4", ".wmv", ".mov", ".flv"]:
        cnt_vid += 1


print("一共有" + str(cnt_vid) + "个视频文件")
