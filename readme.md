# 前排提示

内容尚未建设完成，同志们持续关注，隔一段时间更新一次。

# 简介

身为学委，催班里的同学们交作业不是一件容易的事情。班里有同学经常晚交、不交作业；交了的同学文件也经常喜欢瞎**不按照格式命名，清名单巨麻烦，总之就是拼命折腾学委。\_(:з」∠)\_

这个项目，通过Python的方法，结合一些特定的需求，解决一些办公过程中常见的、繁琐的且通用的问题。

持续建设中，大家可以关注，平日里要打工，不定期更新（提前为自己的咕咕咕找好借口）

## 配置要求

> OS: ``Windows 10``或者``MacOS``均可（注意：``Windows 10``或者``MacOS``在某些方面，比如文件路径的表示上可能略有不同）
>
> Python版本: 本项目采用``3.8.x 64-bit``版本（更高版本应该也能向下兼容吧，还没试过）
>
> 推荐的编辑器：``Pycharm``或``Jupyter``，看你喜好啦~推荐使用虚拟环境（venv），在为本项目进行必要的第三方库配置时，不影响其它Python项目的配置，不污染系统Python的环境。

## 需要用到的Python库

``os``：Python自带库，无需安装。

> 导入方式：

```python
import os
```

***

``shutil``：Python自带库，无需安装。

> 导入方式

```python
import shutil
```

***

``ezexif``：第三方库，打开cmd（或其它操作系统的终端），使用``pip install ezexif``进行安装。

> 用途：处理读取一些照片文件的信息等。
>
> 导入方式：

```python
import ezexif
```

****

``openpyxl``：第三方库，打开cmd（或其它操作系统的终端），使用``pip install openpyxl``进行安装。

> 用途：对excel表格进行读写操作。
>
> 导入方式：原则上采用``import openpyxl``便足以满足开发需要，但是为了尽可能发挥代码提示的功能，推荐如下方式导入

```python
import openpyxl  # 最主要的导入
from openpyxl.utils import cell
from openpyxl import Workbook  # 这个导入和type注释会共同起作用
from openpyxl.worksheet.worksheet import Worksheet  # 这个导入和type注释会共同起作用
```

另注：为了显示代码提示，建议在对工作簿、工作表变量赋值后，用注释在同一行指明类型（格式：``# type: 类型名``），例如

```python
import openpyxl  # 最主要的导入
from openpyxl.utils import cell
from openpyxl import Workbook  # 这个导入和type注释会共同起作用
from openpyxl.worksheet.worksheet import Worksheet  # 这个导入和type注释会共同起作用

workBook = openpyxl.load_workbook(filename=recordPath)  # type: Workbook
workSheet = photoParamBook["示例"]  # type: Worksheet
```

***

``rmbTrans``：第三方库，打开cmd（或其它操作系统的终端），使用``pip install rmbTrans``进行安装。

> 用途：将中文大写的人民币值（例如：``肆仟贰佰壹拾贰元伍角伍分``经过转换变成数值的``4212.55``，单位：元）。
>
> 导入方式：

```python
import rmbTrans
```

****

``python-docx``：第三方库，打开cmd（或其它操作系统的终端），使用``pip install python-docx``进行安装。注意：这个库在导入时使用``import docx``，而不是导入全名。

> 用途：对word文档进行读写操作。
>
> 导入方式：原则上采用``import docx``（注意导入的库名不是python-docx）便足以满足开发需要，但是为了尽可能发挥代码提示的功能，推荐如下方式导入

```python
import docx  # 最主要的导入
from docx.document import Document  # 这个导入和type注释会共同起作用
from docx.table import Table, _Cell  # 这个导入和type注释会共同起作用
from docx.text.paragraph import Paragraph  # 这个导入和type注释会共同起作用
from docx.text.run import Run  # 这个导入和type注释会共同起作用
```
另注：为了显示代码提示，建议在对文档变量赋值后，或者**单独**对段落类型变量、样式块类型变量或表格类型变量赋值后，用注释在同一行指明类型（格式：``# type: 类型名``），例如

```python
import docx  # 最主要的导入
from docx.document import Document  # 这个导入和type注释会共同起作用
from docx.table import Table, _Cell  # 这个导入和type注释会共同起作用
from docx.text.paragraph import Paragraph  # 这个导入和type注释会共同起作用
from docx.text.run import Run  # 这个导入和type注释会共同起作用

doc = docx.Document()  # type: Document
para1 = doc.paragraphs[0]  # type: Paragraph
run1 = para1.runs[0]  # type: Run

table1 = doc.tables[0]  # type: Table
```


****

``pdfplumber``：第三方库，打开cmd（或其它操作系统的终端），使用``pip install pdfplumber``进行安装。

> 用途：对pdf文字类型文件进行操作。
>
> 注意：pdf类文件有扫描类和文字类，这个库针对的是文字类pdf文档，即可用光标选中字符的这类文档。扫描类pdf文档内容，在办公场景下需要我们去提取的概率并不大，故不在此赘述。
>
> 导入方式：

```python
import pdfplumber
```

****

``smtplib``、 ``email``：Python自带库，无需安装。

> 用途：通过SMTP协议（邮件传输协议），设置邮件的发送。
>
> 导入方式：

```python
# 导入smtplib模块(SMTP协议在Python中对应的模块)
import smtplib
# 导入邮件正文内容数据处理模块
from email.mime.text import MIMEText
# 导入邮件协议的协议头模块
from email.header import Header
```

****

## 请注意

* 当您在运行Python的样例时，请预先安装好所需的第三方库，**并关闭resources中所涉及的文件**（如果有），以免因为文件占用导致程序不能自动化处理打开着的文件。
* 使用库的时候，不要死记硬背，常见的思路记住即可，细节的东西可以随机应变或者查文档进行处理。
* 导入第三方库时（尤其是``docx``和``openpyxl``时），部分第三方库的类单独导入进来，结合形如``# type: ...``格式的注释，有利于充分利用好编辑器的代码填充提示，使第三方库的封装更有意义。

## 快捷传送门

本项目中的内容

``Python``官方文档

>  https://docs.python.org/3/

``openpyxl``第三方库官方文档

>  https://openpyxl.readthedocs.io/en/stable/index.html

``python-docx``第三方库官方文档

>  https://python-docx.readthedocs.io/en/latest/

``pdfplumber``第三方库开源代码及文档

> https://github.com/jsvine/pdfplumber
