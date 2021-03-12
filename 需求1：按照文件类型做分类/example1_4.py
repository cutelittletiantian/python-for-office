import os

# 所有的学生名字列表
studentList = ["James", "Harry", "Maggie", "Joan", "Judy", "Fred", "Roy", "Billy", "Louis", "Tony", "Kevin", "Tracy",
               "Vincent", "Jay", "Dean", "Neil", "Faye", "Evan", "Dana", "Kelly"]

submit_names = []
notSubList = []

photoPath = "./resources/ex1_4/tiantian/证件照"
photos = os.listdir(photoPath)
for photo in photos:
    filename = os.path.splitext(photo)[0].lower()
    submit_names.append(filename)

for student in studentList:
    if student.lower() not in submit_names:
        notSubList.append(student)

print(notSubList)