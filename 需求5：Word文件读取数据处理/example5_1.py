import docx
from docx.document import Document
from docx.styles.style import BaseStyle
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

toeflPath = "./resources/ex5_1/TOEFL/托福口语-万能理由和答案.docx"

# 可显示的样式
styleShow = ["Title", "Heading 1", "Heading 2"]

# 加载文件
toeflDoc: Document = docx.Document(docx=toeflPath)
# 读取每一段
for para in toeflDoc.paragraphs:
    paraStyle: BaseStyle = para.style
    if paraStyle.name in styleShow:
        print(para.text)