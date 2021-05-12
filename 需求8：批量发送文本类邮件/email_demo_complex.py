import os
# 发送者的基本配置文件
import config
# 导入SMTP协议库
import smtplib
# 邮件正文内容数据处理模块
from email.mime.text import MIMEText
# 邮件附件处理模块
from email.mime.application import MIMEApplication
# 邮件混合内容数据（既有图片、又有文字还有其它诸如附件这样的文件）的整合处理模块（容器）
from email.mime.multipart import MIMEMultipart
# 邮件头模块
from email.header import Header


"""批量发送邮件应用（综合版）

这里的邮件，不仅有简单的文本，还考虑到含有其它图片、附件等情形。
"""

# 读取联系人的邮件列表（excel文件数据请自备，在resources目录下
# 记得留一个表头，函数逻辑是自动跳过表头读取后面的行
# 每个数据的格式：[收件人姓名, 收件人部门, 收件人邮箱]
# 可以根据个人需求进行更改，但是后面的receiver也要记得做相应调整
contactList = config.read_excel(
    path=os.path.join(config.resourcesPath, config.contactBookName),
    sheet_name=config.contactSheetName
)
# 测试
print("已读取的联系人信息数据如下：")
print(contactList)
print()

# 选择邮箱服务器，登录自己的qq邮箱，得到一个SMTP协议服务对象
# 邮箱服务器设置
mailHost = "smtp.qq.com"
# 选择SMTP协议qq服务，它的端口号是465（官方指定默认端口）
smtpObj = smtplib.SMTP_SSL(host=mailHost, port=465)
# 登录邮箱
smtpObj.login(
    user=config.sender_info["email"],
    password=config.sender_info["password"]
)

# 扫描所有的联系人名单，根据相应的信息组合对应的邮件内容
for receiver in contactList:
    try:
        """组装邮件内容（这里需要逐个添加图片、正文和附件）"""

        # 1. 创建一个打包各种不同邮件内容的容器
        # 默认：mailContent = MIMEMultipart(_subtype="mixed")
        mailContent = MIMEMultipart(_subtype="mixed")
        # ################## 容器的创建内容到这里结束

        # 2. 往邮件内容中添加图片（注意：不是附件），采取html文本添加
        # Step 1: 构造正文图片的html文本
        picHtmlText = """<p><img 
        src="图片的链接请你自己想办法去找......"
        alt="（图片）如果图片显示不出来，这段文字就会顶上去，你可以适当编一下" height="200px"></p>"""
        # Step 2: 创建html文本MIME对象
        picContent = MIMEText(_text=picHtmlText, _subtype="html", _charset="utf-8")
        # Step 3: 添加到multipart容器
        mailContent.attach(picContent)
        # ################## 图片html添加到此结束

        # 3. 往邮件中添加和用户相关的文字文本，由于要显示图片，全文都应当采取html形式
        # Step 1: 构造正文普通文字的内容
        # receiver是正在遍历的列表，数据格式[姓名, 部门, 邮箱]
        wordHtmlText = f"<p>文案请自拟" \
                       f"</p>"
        # Step 2: 创建文字文本MIME对象
        txtContent = MIMEText(_text=wordHtmlText, _subtype="html", _charset="utf-8")
        # Step 3: 添加到multipart容器
        mailContent.attach(txtContent)
        # ################## 和用户相关的普通文字(plain)文本添加到此结束

        # 添加附件（附件素材都在resources/attach目录下）
        # 4. 往邮件中添加附件1
        # Step 1: 构造图片数据流（图片是二进制文件，这种方式可以把图片以二进制数据形式加载进来）
        attachFileOne = open(file="resources/attach/演示附件1.docx", mode="rb").read()
        # Step 2: 创建附件MIME对象
        attachContentOne = MIMEApplication(_data=attachFileOne, _subtype="base64", _charset="utf-8")
        attachContentOne.add_header(_name="Content-Type", _value="application/msword")
        attachContentOne.add_header(
            _name="Content-Disposition",
            _value="attachment",
            filename="演示附件1.docx"
        )
        # Step 4: 添加到multipart容器
        mailContent.attach(attachContentOne)

        # 5. 往邮件中添加附件2
        # 类似的，模仿一下？
        attachFileTwo = open(file="resources/attach/演示附件2.pptx", mode="rb").read()
        attachContentTwo = MIMEApplication(_data=attachFileTwo, _subtype="base64", _charset="utf-8")
        attachContentTwo.add_header(_name="Content-Type", _value="application/vnd.ms-powerpoint")
        attachContentTwo.add_header(
            _name="Content-Disposition",
            _value="attachment",
            filename="演示附件2.pptx"
        )
        mailContent.attach(attachContentTwo)

        """邮件内容组装完成"""

        # ################## 邮件头设置
        # 邮件主题（你可以自拟）
        mailContent["Subject"] = Header("《办公》应用之批量邮件发送测试")
        # 邮件发件人名称，格式“发件人名(可自定义) <发件人邮箱>”（中间一个空格）
        mailContent["From"] = Header(f"{config.sender_info['nickname']} <{config.sender_info['email']}>")
        # 邮件收件人名称，格式“收件人名(可自定义) <收件人邮箱>”（中间一个空格）
        mailContent["To"] = Header(f"{receiver[0]} <{receiver[2]}>")
        # ################## 邮件头设置内容到这里结束

        # ################## 发邮件
        # 使用sendmail(发送人，收件人，message.as_string())发邮件
        smtpObj.sendmail(
            from_addr=config.sender_info["email"],
            to_addrs=[receiver[2]],
            msg=mailContent.as_string()
        )

        # 获取姓名输出“xx邮件发送成功”
        print(f"{receiver}的邮件发送情况：成功")

    except Exception:
        print(f"Error: {receiver}的邮件发送情况：失败")
        continue
