import docx
import pdfplumber
import os

# 文件夹路径
from docx.document import Document

pdfListPath = "./resources/ex7_1"
# 保存路径
saveDocPath = "./resources/ex7_1/Python文稿.docx"

# 遍历pdf文件
pdfList = os.listdir(path=pdfListPath)


# 指定排序规则函数
def pdf_sort(pdf_file: str):
    # 分离出文件名
    pdf_name = os.path.splitext(p=pdf_file)[0]
    # 整数比较
    return int(pdf_name)


# 对文档顺序进行个排序
pdfList.sort(key=pdf_sort, reverse=False)

# 新建存储用的Word文档
newDoc = docx.Document()  # type: Document

# 逐个遍历pdf文档
for pdfFile in pdfList:
    # 组装完整路径
    pdfPath = os.path.join(pdfListPath, pdfFile)
    # 读取pdf文件
    pdf = pdfplumber.open(path_or_fp=pdfPath)

    # 逐页遍历
    for page in pdf.pages:
        textData = page.extract_text()
        # 替换
        textData = textData.replace("PYTHON", "Python")
        # 加入文档
        newDoc.add_paragraph(text=textData)
        # 添加分页
        newDoc.add_page_break()

    # 完成一页的添加
    print(f"{pdfFile}提取完成")

# 最终保存
newDoc.save(path_or_stream=saveDocPath)