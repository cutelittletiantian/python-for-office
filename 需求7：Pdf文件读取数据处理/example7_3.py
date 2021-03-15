import docx
import pdfplumber
from docx.document import Document
from docx.text.paragraph import Paragraph

# 配置参数
# pdf名单
nameListPath = "./resources/ex7_3/名单.pdf"
# 优秀教师名单
teacherDocPath = "./resources/ex7_3/优秀教师.docx"
# 名单
nameList = []

# 打开pdf文档
nameListFile = pdfplumber.open(path_or_fp=nameListPath)

# 逐页遍历文档页面，取表格
# 第一页有表头，特殊情况特殊处理
pageFirst = nameListFile.pages[0]
# 取表
tableFirst = pageFirst.extract_tables()[0]
# 跳过表头，range从1开始
for rowIndex in range(1, len(tableFirst)):
    nameList.append(tableFirst[rowIndex][0])

# 接着读取剩余的表格
for page in nameListFile.pages[1:]:
    table = page.extract_tables()[0]
    for rowIndex in range(len(table)):
        nameList.append(table[rowIndex][0])

# 添加到word文档中
teacherDoc = docx.Document(docx=teacherDocPath)  # type: Document
# 末尾添加行
teacherDoc.add_paragraph(text=" ".join(nameList))
# 输出调试一下看
print(" ".join(nameList))
# 保存
teacherDoc.save(path_or_stream=teacherDocPath)