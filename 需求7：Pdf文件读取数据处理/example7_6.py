import pdfplumber
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph

sourcePath = "/Users/yequ/Desktop/Translation/book.pdf"
targetPath = "/Users/yequ/Desktop/book.docx"

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