import os

docPath = "./resources/review4"

allItems = os.listdir(docPath)


def sort_file(file_name):
    source_name = os.path.splitext(p=file_name)[0]
    file_code = source_name.split(sep="-")[1]
    return int(file_code)


# 执行排序(key传入函数：对函数的形参排序，对其中的返回值规则进行排序)
allItems.sort(key=sort_file)
print(allItems)