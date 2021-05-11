import os
# 发送者的基本配置文件
import config
# 导入SMTP协议库
import smtplib
# 邮件正文内容数据处理模块
from email.mime.text import MIMEText
# 邮件头模块
from email.header import Header

"""批量发送邮件应用（简化版）

这里的邮件，只有简单的文本，暂时没有考虑其它图片、附件等情形。
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


# 选择邮箱服务器，登录自己的qq邮箱，得到一个SMTP协议服务对象
# 邮箱服务器设置
mailHost = "smtp.qq.com"
# 选择SMTP协议qq服务，它的端口号是465（官方指定）
smtpObj = smtplib.SMTP_SSL(host=mailHost, port=465)
# 登录邮箱
smtpObj.login(
    user=config.sender_info["email"],
    password=config.sender_info["password"]
)

# 扫描所有的联系人名单，根据相应的信息组合对应的邮件内容
for receiver in contactList:
    try:
        # 邮件正文设置（创建一个MIMEText对象）
        message = MIMEText(
            _text=f"在这个地方输入正文的内容（支持'\\n'换行）",
            _subtype="plain",
            _charset="utf-8"
        )

        # 邮件头设置
        # 邮件主题
        message["Subject"] = Header("在这个地方输入邮件的主题")
        # 邮件发件人名称，格式“发件人名(可自定义) <发件人邮箱>”（中间一个空格）
        message["From"] = Header(f"{config.sender_info['nickname']} <{config.sender_info['email']}>")
        # 邮件收件人名称，格式“收件人名(可自定义) <收件人邮箱>”（中间一个空格）
        message["To"] = Header(f"{receiver[0]} <{receiver[2]}>")

        # 使用sendmail(发送人，收件人，message.as_string())发邮件
        smtpObj.sendmail(
            from_addr=config.sender_info["email"],
            to_addrs=[receiver[2]],
            msg=message.as_string()
        )

        # 获取姓名输出“xx邮件发送成功”
        print(f"{receiver[0]}的邮件发送情况：成功")

    except:
        print(f"Error: {receiver[0]}的邮件发送情况：失败")
        continue
