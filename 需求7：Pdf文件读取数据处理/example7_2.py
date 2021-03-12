import os
import pdfplumber
import docx
from docx.document import Document

os.chdir(path="./resources/ex7_2")

# PDF文件夹路径
pdfListPath = "Python"
# 目标路径
targetPath = "稿件文档"
# 创建目标路径（如果不存在）
if not os.path.exists(targetPath):
    os.mkdir(path=targetPath)

pdfList = os.listdir(path=pdfListPath)


# 对pdf按章节排序规则函数
def sort_pdf(pdf_file: str):
    pdf_name = os.path.splitext(p=pdf_file)[0]
    return int(pdf_name)


# 执行排序
pdfList.sort(key=sort_pdf)

# 遍历pdf文档，转换
for pdfFile in pdfList:
    # 组装完整文件名
    pdfFilePath = os.path.join(pdfListPath, pdfFile)
    # 单独拎出来文件名，去掉.pdf
    pdfName = os.path.splitext(pdfFile)[0]
    # 创建相应的docx文档
    doc = docx.Document()  # type: Document
    # 打开pdf文档
    pdf = pdfplumber.open(path_or_fp=pdfFilePath)

    # 逐页遍历
    for page in pdf.pages:
        doc.add_paragraph(text=page.extract_text())
        # 分页
        doc.add_page_break()
    # 保存doc文档
    doc.save(path_or_stream=os.path.join(targetPath, f"{pdfName}.docx"))
    print(f"{pdfName}.docx提取完成")