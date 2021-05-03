import pdfplumber
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph

# 注意：book.pdf暂时还没找文本素材，这个地方大家自己随便编点东西，做个pdf文档然后尝试吧
sourcePath = "./resources/ex7_6/book.pdf"
targetPath = "./resources/ex7_6/book.docx"

# 启动pdf文档
sourceFile = pdfplumber.open(path_or_fp=sourcePath)
# 启动Word文档
targetFile = docx.Document()  # type: Document

# 遍历第3页至第10页
for pageNum, page in enumerate(sourceFile.pages[2:10], start=3):
    # 取文本
    text = page.extract_text()
    # 写入word文档
    targetFile.add_paragraph(text=text)
    # 输出完成一页的标记
    print(f"第{pageNum}页完成")

# 保存word文档
targetFile.save(path_or_stream=targetPath)
