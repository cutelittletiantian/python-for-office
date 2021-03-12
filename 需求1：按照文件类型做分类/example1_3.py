import os

# 要查询的文件列表
queryList = ["大象洗澡.mp4", "心动.MP3", "河马洗澡.mp4", "长颈鹿洗澡.mp4", "猎豹.jpg"]

# 将阿文的下载文件夹路径 /Users/yequ/Downloads 赋值给变量downloadPath
downloadPath = ".resources/ex1_3/Downloads"

allItems = os.listdir(downloadPath)

for item in queryList:
    if item in allItems:
        print(item)