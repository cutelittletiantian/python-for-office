# 使用SMTP和MIME协议批量发送含正文、图片、音视频、附件等的邮件（Python语言实现）

> 本文所述的SMTP协议，确切地说，应该是SMTPs协议（SMTP-over-SSL），即基于SSL加密后的一种SMTP协议变种。
>
> 超详细！教科书级别的实践参考！
>
> 如有斟误或补充，请联系作者。





## 1. SMTP和MIME协议简介

**Note：如果您以前没有接触过此类应用，可以在研读完“Python语言实现邮件批量发送”部分的代码后回到此部分进行梳理，如果学有余力，可以继续阅读《计算机网络》相关教科书及其它文献、博客等，深入学习该应用中的牵涉的各种理论知识。**



### 	1.1 SMTP协议

***

#### 		1.1.1 概述

* 中文全名：简单邮件传输协议
* 英文拼写：**S**imple **M**ail **T**ransfer **P**rotocal



**基于文本**的一个协议，指定消息的一个或多个接收者（即：收件人和抄送人），然后消息文本就会传输。

> 最早的SMTP协议只支持传输可打印的7位ASCII字符邮件，二进制文件（例如图片、音视频文件等）都不能传，此外还有各种各样的瓶颈和不足。于是有了MIME协议对SMTP协议做扩展。



``附：SMTPs协议简述``

SMTP协议**基于SSL安全协议**之上的一个变种，安全性更高，防止邮件在传输过程中被截取泄密，还能防止**发送者抵赖（发送者发送邮件以后删除内容，然后拒不承认自己发过这份邮件）**。

本文后面进行实现时，采用的SMTP协议，就是指这个更安全的SMTPs协议。

***

#### 		1.1.2 SMTP协议邮件内容组成部分

SMTP规定：邮件由**邮件头部（header）**和**邮件主体（body）**组成。

> 如果把邮件比作一封信，那么邮件头部就相当于信封，上面有主题、还指定了这封信从何而来（发送方信息）、到哪里去（接收方信息）等这样一些内容。我们可以这样子去进行类比。



1. **邮件头部（header）**

|   字段   |           含义            | 备注 |
| :------: | :--------------------------------------: | :--------------------------------- |
|   **From**   |     邮件发送方的信息      | 常用、**可由系统自动填入**自己信息或自定义。<br />格式：``发送方名称(可自定义) <发送方的邮箱>``（中间有一个空格）<br />格式样例：``define_your_name <your_email_name@the_host_name.com>`` |
|    **To**    |     邮件接收方的信息      | 常用、且**最重要**。<br />格式：``接收方名称(可自定义) <接收方的邮箱>``（中间有一个空格）<br />格式样例：同前**From**字段所述。 |
| **Subject** |        邮件的主题         |                常用、且**最重要**。<br />内容可自定义。                |
|   Date   |         发信日期          |                  常用、**通常由系统自动填入。**                  |
| Reply-To | 接收方回信时所采用的地址  | 较少用。<br />回信采用地址可以与发送方不同。<br />使用场景：借自己同学邮箱发送邮件后，想让回复的邮件到自己的邮箱中来，可以预设这个属性。<br />格式样例：同前**From**字段所述。 |
|    Cc    | 抄送（Carbon copy）方信息 |                         较少用。<br />格式：``抄送方名称(可自定义) <抄送方的邮箱>``（中间有一个空格）。<br />格式样例：同前**From**字段所述。                         |
| Bcc | 暗送（Blind carbon copy）方信息 | 较少用。<br />格式：``抄送方名称(可自定义) <抄送方的邮箱>``（中间有一个空格）。<br />格式样例：同前**From**字段所述。 |

Note：

* 大家**可以先只着重了解From，To和Subject**三个字段，后面实现邮件批量发送的演示代码，也只用到了这三个。

* 有关**直接发送、抄送和暗送的区别**这里不进行赘述，也不是这篇文档关注的主要矛盾，我们考虑另起文档进行实验报告。



2. **邮件主体（body）**

主体内容即邮件里面写了什么东西，这个部分可以让**用户“自由”编辑**。

为什么“自由”两个字打了引号呢？因为SMTP协议只支持传输7位ASCII码字符，假如说只有SMTP协议，那么最后的结果就是：

* 中文不支持、非英语国家的字符不支持
* 不能传送二进制流数据的文件（例如图片、音视频、附件等）
* 除此之外还有其它很多很多的问题，这里不赘述了

所以说这种“自由”，仅限于7位ASCII码支持的字符数据可自由编辑，其它的数据没法传呀......

**于是后来有了MIME协议，解决了上述的局限，所谓的“用户‘自由’编辑”，终于可以把“自由”的引号摘掉了。我们接着往下看。**

***



### 	1.2 MIME协议

***

#### 		1.2.1 概述

* 中文全名：多用途互联网邮件扩展协议
* 英文拼写：**M**ultipurpose **I**nternet **M**ail **E**xtensions



**MIME协议是对SMTP协议的扩展协议。**



MIME协议中，邮件是一个或者多个“板块”（或者说**“报文”**）进行封装组装，每一个报文都支持它的数据被解析为某一种特定的内容类型**Content-Type**（如普通文本、html文本、图像、音视频、**multipart容器**等等），其中**multipart容器**非常有用，允许单个报文里面嵌套多个彼此独立的子报文。每个子报文还可以继续添加数据，并设置自己独立的类型、编码等。

上面这段话说了很多内容，简言之，**直接现象就是，MIME协议规定的邮件，可以支持各种不同的语言，可以同时放图片、音视频文件、附件，可以内嵌html资源等等**，内容更加丰富、邮件也更灵活。

> 始终牢记一点：MIME协议是SMTP协议的扩展，MIME协议是SMTP协议的扩展，MIME协议是SMTP协议的扩展。脑海中默念3次。

***

#### 		1.2.2 MIME协议邮件内容组成部分

> MIME协议是SMTP的扩充，所以MIME协议同样支持SMTP协议的7为ASCII编码字符，同样包含了SMTP协议指定的邮件头部字段。
>
> 不过既然是扩展，MIME协议相较SMTP协议，还增加了一些辅助字段。

与SMTP协议一样，MIME协议大体还是分为了**邮件头部（header）**和**邮件主体（body）**两个部分，不过在SMTP协议的<u>原有基础</u>上，做了扩充。



1. **邮件头部（header）**

> 由于MIME协议支持一个容器报文下嵌套多个不同类别子报文，每个报文都会有扩展的头部字段加以标识和区分。

|           字段            |                             含义                             | 备注                                                         |
| :-----------------------: | :----------------------------------------------------------: | :----------------------------------------------------------- |
|  **SMTP协议原设有字段**   |                           见1.1.2                            | 见文档1.1.2部分“SMTP协议邮件内容组成”邮件头部（header）<br />原有的字段在MIME协议邮件头部依然适用<br />且**From，To，Subject**属性仍非常重要 |
|       MIME-Version        |                         MIME协议版本                         | 常见版本1.0，**无需设置**，按默认来即可。                    |
|      **Content-Id**       |                        MIME邮件标识符                        | 如果采用**multipart类型容器**：<br />仅从邮件中**标识一个子报文**。<br /><br />如果邮件是non-multipart（非容器）：<br />标识一个完整的邮件。<br /><br />常用情景：<br />将本地资源嵌入邮件，需要使用multipart容器，<br />并构造子报文和对应子报文的Content-Id，<br />这个Content-Id将用于给html文本子报文调用显示。<br /><br />**较常用**字段之一（尤其是添加**内嵌资源**的时候）。 |
|     **Content-Type**      | 邮件传输数据<br />是什么样的文件类型<br />（或者说：**MIME类型**） | 指定传输来的**数据解析成什么类型的文件**。（eg：文本文件？还是ppt文件之类的？）<br />可选参数：charset，说明邮件主体内容**以什么编码呈现**（防止乱码）<br />（通常建议取**utf-8**）<br />**最常用**字段之一。<br /><br />Content-Type有很多类型，这种情况下要能查表指定。<br />Content-Type查询表：<br />参考1：[MIME 参考手册 (w3school.com.cn)](https://www.w3school.com.cn/media/media_mimeref.asp)<br />参考2：[HTTP content-type \|菜鸟教程 (runoob.com)](https://www.runoob.com/http/http-content-type.html)<br />具体的做法，可以结合第2部分2.5.2节，学习查询中的**“渔”**<br /><br />如果邮件采用**multipart类型容器**：<br />仅说明邮件中**一个子报文**传输数据的类型。<br />如果邮件是non-multipart（非容器）：<br />说明的就是一个完整邮件所传输数据的类型。<br /><br />格式：``Content-Type: 父类型/子类型[; charset=内容呈现的编码字符集（非必选）]``<br />格式举例：``Content-Type: text/html; charset=utf-8``（注意冒号、分号后都有一空格） |
| Content-Transfer-Encoding |                         内容传送编码                         | 描述在**传送过程中**邮件的主体是如何进行编码的。<br />（默认值为**base64**，通常取默认值就好了，传送过程的编码我们无需操心，放心交给协议去做就好） |
|    Content-Description    |                           内容描述                           | 可读字符串，辅助描述邮件内容。<br />可自定义、**非必选**。   |

**另一个对MIME协议扩展的字段**

下面的这个字段并不是MIME协议指定的字段，而是HTTP协议响应头（response header）的一个字段，但可以用于MIME协议邮件头部进行扩展。

|        字段         |      含义      | 备注                                                         |
| :-----------------: | :------------: | ------------------------------------------------------------ |
| Content-Disposition | 附件和别名配置 | 主要在需要**指定邮件附件**的场景下进行设置。<br />可以用在multipart子报文中。<br /><br />指示浏览器以何种形式展示附件。<br /><br />可选参数：``filename``，如果设置了，<br />表示浏览器下载附件时的默认文件名，<br />如果不设置可能就是一堆乱码。 |

备注中所述的“附件的展示形式”有两种形式：

* attachment：下载展示。设置这个值以后，消息体应该被下载到本地；大多数浏览器会呈现一个“保存为”的对话框。
  假如将 `filename` 的值预填，浏览器显示这个附件的名字就是``filename``，下载后的附件名默认就是这个``filename``。
* inline：内联展示。浏览器响应体（response body）会以页面的一部分或者整个页面的形式展示。

**下面是Content-Disposition字段的使用格式**

* 使用格式：

``Content-Disposition: <展示形式>[; filename="自定义文件名"]``

* 格式样例：

```python
Content-Disposition: inline  # 内联展示形式
Content-Disposition: attachment  # 下载展示形式
Content-Disposition: attachment; filename="filename.jpg"  # 下载展示，同时浏览器默认显示、保存附件文件名为"filename.jpg"
```

**需要注意的是：Content-Disposition这个扩展字段，在邮件发送的情境下，主要作用是设置附件在接收方浏览器中显示的文件名。不指定这个字段也可以，但是接收方显示的附件文件名可能会是乱码，不过不影响附件内容。**



> Note：
>
> * 同是编码，Content-Type字段中的charset和Content-Transfer-Encoding是有区别的
>   * Content-Type中的charset：描述的是数据中**实际内容**应当采取什么编码呈现，确保**用户不会看到乱码**（通常为**utf-8**）
>   * Content-Transfer-Encoding：描述的是数据在**网络传输当中**按照什么方式压缩和传输邮件主体的（通常取为**base64**）
>
> * 关于字段设置的注意事项
>   * 用到multipart容器组合报文：仅容器需要设置SMTP原有字段（收发方信息、主题等），内部子报文无需设置；其它字段按需照常设置
>   * 未用multipart容器组合报文：该报文就是邮件的内容，需要设置SMTP原有字段
> * 有关字段具体设置的实例，后面会有例子具体说明的。



2. **邮件主体（body）**

主体内容即邮件里面写了什么东西，这个部分可以让**用户自由编辑**。

不同于SMTP协议中的“自由”需要打上引号，这里的自由是指：

* 支持不同语种字符
* 支持嵌入图片、音视频等资源
* 支持各种类型附件

只需要注意，用户在构造上述数据时，需要遵守MIME协议，**文本文件传入文本数据到MIME报文主体部分、二进制文件传入数据二进制流到MIME报文主体部分**，即可完成必要的设计。

***



### 1.3 SMTP协议和MIME协议的关系

下面这张图，非常清晰地表达了MIME协议和SMTP协议的关系。（图片摘自：谢希仁《计算机网络（第7版）》第292页，图6-18）

<img src="https://bl3302files.storage.live.com/y4mOUfTrojFFbgd2hGUBichoXnU0ga0MonXW9yTde5DO_bc-dfst705HZoNGDOMVhrUmrZDVFtqBGw95COQrckTuq_K2VC6dEihOu0Ah7SV-Df8RB-pF50PKQJFtqulwk-Vz8xsZx9G488c93kSmv-JnKvXOlC4wrUbkIpitVQqR1o0R1wW314RBGhMfBELB3_G?width=364&height=213&cropmode=none" width="364" height="213" alt="图：MIME和SMTP的关系（摘自：谢希仁《计算机网络（第7版）》第292页，图6-18）" />

MIME协议没有取代或者更改SMTP协议，只是在SMTP协议的基础上进行了扩充，所以SMTP可识别的字段在MIME协议下也是可识别的。

* 发送时，不管用户编辑的邮件内容是否为7位ASCII字符、是否为二进制文件数据流，MIME协议都会将这些内容转换为SMTP协议支持的7位ASCII码（大多数可能是乱码）再进行传输；
* 接收时，MIME协议会通过内部机制，将SMTP协议传输而来的7位ASCII码（大多数可能是乱码）转换为二进制文件数据流或者其它编码集的字符，使文本、图片、音视频、附件等都能按发送方期盼原样呈现。

***





## 2. 发送邮件的整体思路、过程和实现方法

结合1.3部分中描述的“SMTP协议和MIME协议的关系图”，我们可以知道，**发送邮件的过程中，MIME协议和SMTP协议都用在了哪里**：

* **站在用户的视角，我们邮件内容（报文）的准备工作应当以MIME协议为规则**，不过在设置MIME协议的一些字段的时候，其实有的字段（From，To，Subject）这样的也算是SMTP协议的字段，只不过MIME协议基于此有对字符编码和内容丰富性的扩展而已。

* **想让邮件发送，我们只需调用SMTP协议有关接口，指定收发人和待传送报文**。此后MIME协议按照内部机制，将数据报内容转换为SMTP协议可以接受的7位ASCII码，而SMTP协议则按照既定的传输机制（建立连接、邮件传输、连接释放）将邮件从发送方服务器传输到接收方服务器。



### 2.1 发送的总体流程

***

> 这个板块的内容，建议你看完后面的例子以后再回来回顾一下梗概过程，这个就是我们第3部分代码演示的整体流程。

1. 登录邮件服务器，需要指定以下内容
   * 邮件服务器的域名（例如：QQ邮件系统服务器smtp.qq.com，网易邮件系统服务器：smtp.163.com等）
   * 邮件服务端口号
   * 登录用户名：发送方的邮箱
   * 登录授权码：**这里的授权码可能并不是登录邮箱的密码，而是通过邮箱服务器规定的方式获取的授权码**，需要结合相应邮件系统的文档操作（例如接下来的例子中，qq邮箱就需要结合官方文档的引导来获取授权码，详见3.2节）

2. 准备邮件的内容报文：结合MIME协议准备数据、拼装MIME报文

用户可以选择直接使用nonmultipart类型的MIME报文，指定邮件头部和数据以后直接发送，但**为了让解决方案具有通用性，以下只描述使用multipart类型的MIME报文时需要考虑的数据准备情景**。根据具体的应用需求，在选择混合类型容器报文后，您可能会选择下面其它情景中的一个或多个。

* 混合类型容器报文（Content-Type: multipart/mixed）：这个容器报文将包含邮件的所有下述其它子报文，发送邮件前，将子报文按顺序添加进容器报文，最后发送容器报文即可。
* 文本数据报文（text）：其中还可以分为两种子类别
  * 普通文本数据报文（Content-Type: text/plain）：编辑什么文本内容就显示什么内容
  * html文本数据报文（Content-Type: text/html）：编辑的文本按照html解析，可以引入想要嵌入的资源，在邮件系统进行显示
* 内嵌资源数据报文：想在邮件正文里内嵌显示的图片、音视频等资源
* 附件资源数据报文：想在邮件里通过下载呈现的任何类型文件

**所有的子报文在创建完成后，要按照次序添加进容器报文。**

> 准备这些报文的详细过程，请阅读2.2—2.5。可以不用一次性全读，会很晕；可以在看了第3部分的实现案例以后，再回过头来看下面的内容。

3. 设置multipart容器报文的**头部**，指定邮件的发送方（From）、接收方（To）以及邮件主题（Subject）

4. 将添加好了各种子报文的**multipart容器报文作为整体发送**，发送前给出收发方邮箱的参数，发送完毕随后可断开连接。



**基于上述步骤，以下是Python语言实现时主要需要导入的库和实现的梗概过程**

> 不要嫌弃下面的内容比较长，代码就几行，其它都是注释......😂

```python
# SMTP协议封装库，负责登录和发送
import smtplib
# MIME协议邮件头模块，专门负责设置邮件头部，后面主要拿它生成收、发方和主题的信息（From, To, Subject）
from email.header import Header
# MIME协议各种邮件体模块，专门负责面向用户准备传输数据的报文(MIMExxxx表示各种类型的报文对象)
from email.mime.xxxx import MIMExxxx
# email.mime模块直接支持的MIMExxx报文主类如下：（具体说明见2.4.3）
# from email.mime.base import MIMEBase  # 所有MIME报文类的基类，用于实现多态，显然这个不是我们要用的
# from email.mime.multipart import MIMEMultipart
# from email.mime.nonmultipart import MIMENonMultipart
# from email.mime.message import MIMEMessage  # 这个也非常少用
# from email.mime.text import MIMEText
# from email.mime.image import MIMEImage
# from email.mime.application import MIMEApplication
# from email.mime.audio import MIMEAudio
# 其它不直接支持的主类，依然可以通过MIMEApplication设置主类从而继续支持


# 1. 与SMTP邮件服务器建立连接、登录邮件服务器

# 建立连接
# 例如：QQ邮箱的SMTP邮件服务器，host="smtp.qq.com", port=465
smtpObj = smtplib.SMTP_SSL(host="填入SMTP邮件服务器的名称", port=填写端口号，一个整数)

# 执行登录
# 注意：password可能并不是真正的密码，可能是你的邮箱系统中生成的授权码，这个需要由你自主获取并复制过来。
# 在3.2节中，提供了QQ邮箱中获得授权码的方法。
smtpObj.login(
    user="填入自己的邮箱",
    password="填入获取的授权码"
)


# 2. 准备邮件内容（MIME报文的拼装）
# 构造MIME报文容器。
multiMsg = MIMEMultipart()
# 随后读入其它数据一个个往报文容器里添加
# ——这个过程在本节略去，从2.2~2.5节详细说明了这个阶段的各种报文的准备方法，请根据需要选择性查看并动手实验
# ......
# 总之，邮件内容的MIME报文拼装完成后，存在了multiMsg当中


# 3. 给MIME报文容器multiMsg的邮件头部添加发送方（From）、接收方（To）以及邮件主题（Subject）字段
# 注意：发件人和收件人头部的格式
#
# 名称<邮箱>
# ——举例：小添添<emailname@hostname.com>
#
# 邮箱两端应使用一对尖括号“<>”括起来。
multiMsg["Subject"] = Header(s="这里自定义填写主题名称", charset="utf-8")
multiMsg["From"] = Header(s="发件人名称(可自定义为姓名或者昵称)<发件人的邮箱>", charset="utf-8")
multiMsg["To"] = Header(s="收件人名称(可自定义为姓名或者昵称)<收件人的邮箱>", charset="utf-8")

# 以下方式的设置与上述等价，可按照个人习惯选择一种方式进行设置
# multiMsg.add_header(_name="Subject", _value="这里自定义填写主题名称", charset="utf-8")
# multiMsg.add_header(_name="From", _value="发件人名称(可自定义为姓名或者昵称)<发件人的邮箱>", charset="utf-8")
# multiMsg.add_header(_name="To", _value="收件人名称(可自定义为姓名或者昵称)<收件人的邮箱>", charset="utf-8")


# 4. 发出邮件，主要形参：
# from_addr: 发件人的邮箱(两端不带尖括号)
# to_addrs: 收件人的邮箱(两端不带尖括号)
# msg: 第2步中拼装完成的MIME报文multiMsg
# 
# 注意——msg发出的内容，要把MIME报文做转换，转换有2种形式：
# as_string()和as_bytes()
# 两种方式转化报文的形式不尽相同，但皆可行，您可以试着输出观察二者的区别:
# print(multiMsg.as_string())
# print(multiMsg.as_bytes())
smtpObj.sendmail(
    from_addr="填入发件人的邮箱",
    to_addrs="填入收件人的邮箱",
    msg=multiMsg.as_string()
)
# 发送完毕，即可退出与SMTP服务器之间的连接
smtpObj.quit()
```

***



### 2.2 Multipart容器报文的准备

***

大致分如下步骤准备。

1. 创建MIME容器报文，进行如下配置

* 容器报文头
  * 指定Content-Type：**主类为multipart，子类取默认值mixed**

``Content-Type: multipart/mixed``

2. 在其它子报文创建完成后，依次地附加（attach）到这里的容器报文主体中来。

> 关于multipart类型容器MIME报文的子类型特殊说明

容器报文（主类为**multipart**）的Content-Type，**子类**主要有3种：mixed，related，alternative。**默认情况下选择mixed即可，无需想太多，因为mixed类型支持的子报文种类最多。**

* 在邮件中要添加附件，必须定义multipart/mixed段；

* 如果存在内嵌资源，**至少**要定义multipart/related段；

* 如果纯文本与超文本共存，**至少**要定义multipart/alternative段。

上述内容中，什么是**“至少”**？举个例子说，如果只有纯文本与超文本正文，那么在邮件头中将类型扩大化，定义为multipart/related，甚至multipart/mixed，都是允许的。这张图片就较为形象地揭示了mixed，relative，alternative之间的关系。

<img src="https://bl3302files.storage.live.com/y4m658CCr1Zr4OydP5j6gGNGfZEkWDtf5z_jqRbWDW9cpKYB3OnP3K7_6p7UKRz0mwea7HEwUHVzx-TaF0-BPa-kJyDo8IZGYkLGy6Aw1dercUwOaUk6luzZpHLxKqKgKXLUBqNCv3XZzXNbQmADqQ382LUL0EkTW1p3UvZivIiBDpw6Q9taSv7bfrHIV6jen94?width=499&height=363&cropmode=none" width="499" height="363" alter="multipart报文各种子类之间的关系" />



**以下是一种Python实现方式（可结合第3部分程序演示案例加深印象）**

```python
from email.mime.multipart import MIMEMultipart

# 构造MIME报文容器
# multiMsg = MIMEMultipart(_subtype="mixed")
# 构造函数默认_subtype为mixed了，所以也可以不指定
multiMsg = MIMEMultipart()
```

这一步完成了以后，实际上就是创建了一个邮件内容报文，**指定了邮件头部（header）的Content-Type为multipart/mixed**，稍后它的邮件主体（body）里会添加各种子报文数据。

``补充说明``

上面的创建，等价于下面的方式：

```python
from email.mime.multipart import MIMEMultipart

# 构造MIME报文容器
multiMsg = MIMEMultipart()
# 设置MIME报文头部MIME类型（Content-Type）的父类/子类值为multipart/mixed
multiMsg.replace_header(_name="Content-Type", _value="multipart/mixed")
```

**这种创建方式，其它类型的MIME报文也是类似的，后面就不再赘述这个了**。这种写法可以自由指定<u>任何</u>想要设置的MIME协议头字段，更灵活。下一次见到这种写法，是在2.4内嵌资源和2.5附件资源数据报文准备部分，我们有需要用类似的方式来**添加**（``add_header(_name="...", _value="...")``）其它的头字段。

***



### 2.3 文本数据报文的准备

***

#### 2.3.1 普通文本数据报文

大致分如下步骤准备。

1. 准备好自定义的文本数据。
2. 创建MIME子报文，进行如下配置：

* 子报文头
  * 指定Content-Type：**主类为text，子类为plain**，通常设置编码字符集为utf-8

``Content-Type: text/plain; charset: utf-8``

* 子报文体：加入之前准备好的文本数据。

3. 将子报文添加到multipart容器报文中



**以下是一种Python实现方式（可结合第3部分程序演示案例加深印象）**

```python
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 准备MIME容器报文
multiMsg = MIMEMultipart()

# 1. 准备文本数据
txtData = "就在这里吧，输入进来你想要的文本内容，随便输入，哈哈哈......也可以从文本文件中以文本方式读入数据"

# 2. 创建文本子报文
# _text: 添加准备好的数据
# _subtype: 指定MIME类型子类为plain。父类已经为text了，这是由MIMEText默认设置好的
# _charset: 指定文本以什么编码字符集显示，通常设为utf-8，否则可能乱码
txtMsg = MIMEText(_text=txtData, _subtype="plain", _charset="utf-8")

# 3. 文本子报文加入容器报文
multiMsg.attach(payload=txtMsg)
```

***

#### 2.3.2 html文本数据报文

**Note: HTML文本的语法不是本文档讨论的重点，因此接下来提到HTML文档，都是直接给出的，如有不解可自行查找相关资料，了解即可。**

大致分如下步骤准备。

1. 准备好自定义的**html文本**数据（可以在支持html的编辑器先编辑好再读进来相应的文本内容，或者复制粘贴😶）。
2. 创建MIME子报文，进行如下配置：

* 子报文头
  * 指定Content-Type：**主类为text，子类为html**，通常设置编码字符集为utf-8

``Content-Type: text/html; charset: utf-8``

* 子报文体：加入之前准备好的html文本数据。

3. 将子报文添加到multipart容器报文中

> html文本数据报文，往往是用于引用一些外部资源，或者设置样式让邮件形式更丰富。
>
> 参考2.4节通过html文本数据报文，结合外部资源嵌入显示更丰富的邮件内容。



**以下是一种Python实现方式（可结合第3部分程序演示案例加深印象）**

```python
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 准备MIME容器报文
multiMsg = MIMEMultipart()

# 1. 准备文本数据（下面只是很简单地复制过来了一段html模板文本）
htmlData = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Title</title>
</head>
<body>
    <p>test</p>
</body>
</html>"""
# 也可以选择编辑好html文件了以后，再把相应的html文件数据读进来
# htmlData = open(file="resources/text/htmlDemo.html", mode="r", encoding="utf-8").read()

# 2. 创建文本子报文
# _text: 添加准备好的数据
# _subtype: 指定MIME类型子类为html。父类已经为text了，这是由MIMEText默认设置好的
# _charset: 指定文本以什么编码字符集解析，通常设为utf-8，否则可能乱码
txtMsg = MIMEText(_text=htmlData, _subtype="html", _charset="utf-8")

# 3. 文本子报文加入容器报文
multiMsg.attach(payload=txtMsg)
```

***



### 2.4 内嵌资源数据报文的准备

> 这个过程需要联合html文本进行显示

***

#### 2.4.1 内嵌第三方链接资源

大致分如下步骤准备。

1. 准备好需要嵌入的**外部资源**（音频、视频、图片等）**链接**（后不妨简称为“外链”）
2. 准备html文本，写入外链
3. 创建MIME子报文，进行如下配置：
   * 子报文头
     * 指定Content-Type：**主类为text，子类为html**，通常设置编码字符集为utf-8
       * ``Content-Type: text/html; charset: utf-8``

   * 子报文体：加入之前准备好的html文本数据。

3. 将子报文添加到multipart容器报文中

> 关于嵌入资源链接来源的特别说明

**链接嵌入到邮件html文本中才会显示出来，所以链接准备要合理，否则邮件中可能无法显示，或者邮件服务系统可能会拦截。**

* 资源链接不能是本地路径
* 资源链接来自互联网的搜索结果时，这个链接的内容要能够支持外链引用
* 本地资源可以上传至一个云平台的文件服务器，且这个文件服务器能够生成用于外链引用的资源链接
  例如：
  * OneDrive云（微软），绝大多数外链都适合生成
  * 适合用于图片外链生成的“图床”：[Image Upload - SM.MS - Simple Free Image Hosting](https://sm.ms/)

**引用外链时，要适当考虑外链服务器的稳定性，否则外部资源可能会时而能显示，时而又不能显示。**



**以下是一种Python实现方式（可结合第3部分程序演示案例加深印象）**

这里引用的外部资源，是武汉理工大学南湖图书馆的照片。

来源：

https://bl3302files.storage.live.com/y4mTM2sZADmrTf1IXMUhdAOShJoHNduCqoGVp1Sek2DY2M9DoOlXQuucmFATpcjjfuDgyEgMK1dqmwflD6c9Cip7BSHHvL3THe5ZTTQjhi8bPjpQDPKe-xeO9vbhqrzzQ1erS_pvopTF5ZJnGnRLijaDjJiRsXZYcuYZswbTQJAHyzTTGHkecplIm5P9Bx2zV-f?width=660&height=440&cropmode=none

如下图所示。

<img src="https://bl3302files.storage.live.com/y4mV_yck_wisbER9GaLbC6mCr5-qfJQ8RNyq-4d4g8rk_tksQb8dmm8I6R2TS8ONpJr6Me4FYXSqxJXMznkL8ZNEsX6I4uiz9HojUcxYVP3-8kaYgjIbZuSFm-GiOdR5xm8v1uW9-MgiA1Ezlp-vxt25afV_DCNbJKj3Ej8UfzX4SGf1Q4uF7AnFkDhJ6sb64nx?width=660&height=440&cropmode=none" width="660" height="440" />

> 注意：嵌入的资源，如果是外链，这个外链必须是该资源文件服务器的**共享权限**外链，且邮件中的外链嵌入资源**通常是图片**外链，其它类型文件外链可能不支持。（亲测）

```python
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 准备MIME容器报文
multiMsg = MIMEMultipart()

# 1. 准备图片外链
externalLink = "https://bl3302files.storage.live.com/y4mTM2sZADmrTf1IXMUhdAOShJoHNduCqoGVp1Sek2DY2M9DoOlXQuucmFATpcjjfuDgyEgMK1dqmwflD6c9Cip7BSHHvL3THe5ZTTQjhi8bPjpQDPKe-xeO9vbhqrzzQ1erS_pvopTF5ZJnGnRLijaDjJiRsXZYcuYZswbTQJAHyzTTGHkecplIm5P9Bx2zV-f?width=660&height=440&cropmode=none"

# 2. 准备html文本，加入外链（可以提前编辑好html文本再把需要的标签复制过来）
htmlText = \
    f'''
    <div>
        <img src={externalLink} width="660" height="440" />
    </div>'''
# print(htmlData)

# 3. 创建html文本子报文
# _text: 添加准备好的数据
# _subtype: 指定MIME类型子类为html。父类已经为text了，这是由MIMEText默认设置好的
# _charset: 指定文本以什么编码字符集显示，通常设为utf-8，否则可能乱码
txtMsg = MIMEText(_text=htmlText, _subtype="html", _charset="utf-8")

# 4. 文本子报文加入容器报文
multiMsg.attach(payload=txtMsg)
```

***

#### 2.4.2 内嵌本地资源

大致分如下步骤准备。

1. 将本地资源文件以<u>二进制</u>方式**打开并读入**
2. 创建内嵌资源的MIME子报文，进行如下配置：
   * 子报文头（两件事情）
     * 根据**文件的类型（后缀名）**查表，指定其**对应的Content-Type**的主类和子类，由于是二进制文件，不用设置编码字符集
       * 这里所说的“查表”，是指查文件后缀名与MIME类型（Content-Type）映射表，这里给出两个参考表链接：[MIME 参考手册 (w3school.com.cn)](https://www.w3school.com.cn/media/media_mimeref.asp)，[HTTP content-type | 菜鸟教程 (runoob.com)](https://www.runoob.com/http/http-content-type.html)
         查表不难，会查字典也就会用表查Content-Type，详细说明见2.5.2部分
     * 头部添加**Content-Id**值，指定一个**cid链接**（形如：\<resource\>这种格式的，两边用尖括号框起来），这个链接后续会用在html文本中进行引用
   * 子报文体：加入第1步读进来的二进制数据
3. 将创建好的内嵌资源子报文添加到容器报文中
4. 按照2.3.2节的步骤，创建html文本报文，在适当的位置用cid链接进行嵌入
5. 将创建好的html子报文添加到multipart容器报文中

> 特别说明

**使用此方式会让邮件变得很庞大（如果内嵌资源数据过于庞大，发送出去要花很长时间），但可以嵌入的资源更丰富**。所以，使用此法嵌入资源的时候，要注意控制文件大小。



**以下是一种Python实现方式（可结合第3部分程序演示案例加深印象）**

这里引用的资源文件，是一段把水倒入杯子中的视频（来源：[Pouring Water in Drinking Glass · Free Stock Video (pexels.com)](https://www.pexels.com/video/pouring-water-in-drinking-glass-853752/) ，非商用素材，使用前视频已下载到本地，并在本地读取嵌入）

> 注意：嵌入的资源，如果来自本地，相比外链可嵌入的内容会更加丰富（亲测）。例如下面的这个例子中，我们在邮件正文内嵌入了视频。

```python
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# 准备MIME容器报文
multiMsg = MIMEMultipart()

# 1. 本地资源数据文件二进制方式读入
# 项目文件中，这个开源文件的相对目录是"resources/video/Pouring Water in Drinking Glass.mp4"，请根据资源所在具体位置进行调整
resourceData = open(file="resources/video/Pouring Water in Drinking Glass.mp4", mode="rb").read()
# 2. 创建内嵌资源的MIME子报文
resourceMsg = MIMEApplication(_data=resourceData)
# 查表，找到.mp4类型的文件的“主类/子类”为"video/mpeg4"。设置Content-Type
resourceMsg.replace_header(_name="Content-Type", _value="video/mpeg4")
# 增加Content-Id字段，创建cid，这里的cid就是<water>了
cid = "water"
resourceMsg.add_header(_name="Content-Id", _value=f"<{cid}>")

# 3. 内嵌资源报文添加进容器
multiMsg.attach(payload=resourceMsg)

# 4. 准备html文本，加入cid（可提前编辑好html文本再复制粘贴过来）
htmlText = \
    f'''
    <div>
        <video src="cid:{cid}"
               alt="video: Pouring Water in Drinking Glass.mp4"
               loop="loop" width="30%" autoplay="autoplay">
               视频：Pouring Water in Drinking Glass
        </video>
    </div>
    '''
# print(htmlData)

# 创建html文本子报文
# _text: 添加准备好的数据
# _subtype: 指定MIME类型子类为html。父类已经为text了，这是由MIMEText默认设置好的
# _charset: 指定文本以什么编码字符集显示，通常设为utf-8，否则可能乱码
txtMsg = MIMEText(_text=htmlText, _subtype="html", _charset="utf-8")

# 5. html文本子报文加入容器报文
multiMsg.attach(payload=txtMsg)
```

***

#### 2.4.3 补充说明：Content-Type与email.mime模块的联系

``关于Content-Type和Python的email.mime模块的对应关系，补充一点``

注意：Python的email.mime模块只特别定义了**一些常见Content-Type主类（主类/子类）的专用MIME报文类**（创建这样的报文类后只用指定子类，如不指定子类则取默认值）。

email.mime模块所定义的常见Content-Type主类的专用MIME报文类有如下所述：

* MIMEBase：所有MIME报文类的父类，通常不会直接用它
  * MIMEMultipart：multipart**容器**MIME报文类。**Content-Type主类为multipart**。
  * MIMENonMultipart：**非容器**MIME报文类的父类，通常不会直接使用它。
    * MIMEMessage：主要针对.nws后缀的MIME报文类，很少用。
    * MIMEText：主要针对**文本**文件（.txt，.html等）配置的MIME报文类。**Content-Type主类为text**。
    * MIMEImage：主要针对**图片**文件（.jpg，.png，.bmp等）配置的MIME报文类。**Content-Type主类为image**。
    * MIMEAudio：主要针对**音频**文件（.mp3，.mp4等）配置的MIME报文类。**Content-Type主类为audio**。
    * MIMEApplication：可针对**各种文件**（包括上述这些文件，以及没有被email.mime模块现有类直接囊括的文件）配置的MIME报文类。**Content-Type主类默认为application，但是这个可以改**。

上述这些MIME报文类的层级关系，其实也是彼此之间存在的**继承和多态**的关系。

> MIMEApplication类特别说明

对于其它的Content-Type文件（如视频主类video），**email.mime模块没有专门的MIME报文类**（比如没有email.mime.video.MIMEVideo类）的，可以**统一使用email.mime.application.MIMEApplication创建报文，并修改Content-Type字段为"video/xxx"**（xxx是什么可以去查表），例如针对某视频文件（主类为video），可以这样构造它的MIME子报文。

```python
resourceMsg = MIMEApplication(_data=resourceData)
resourceMsg.replace_header(_name="Content-Type", _value="video/mpeg4")
```

> Note：若Application主类报文没有指定_maintype参数，则它的默认值也为application，这点与其它的MIME报文对象是一致的。
>
> 而且这个世界上，大多数MIME报文的Content-Type主类还是application的。

此外，**即使是有专用MIME报文类的Content-Type主类，也能采用上面这种方式构造**。我们在构造MIME对象时，可以按照喜好自由选择。

***



### 2.5 附件资源数据报文的准备

***

#### 2.5.1 准备步骤

大致分如下步骤准备。

1. 将本地资源文件以<u>二进制</u>方式**打开并读入**
2. 创建MIME子报文，进行如下配置：
   * 子报文头部（2件事情）
     * 根据**文件的类型（后缀名）**查表，指定其**对应的Content-Type**的主类和子类，由于是二进制文件数据，不用设置编码字符集
     * 头部设置MIME扩展协议的字段**Content-Disposition**，设置显示模式为附件**attachment**，同时设置显示文件名**filename**（否则接收用户看到浏览器显示的文件名是乱码）。
   * 子报文主体：加入第1步读进来的二进制数据
3. 将创建好的子报文添加到multipart容器报文中



**以下是一种Python实现方式（可结合第3部分程序演示案例加深印象）**

这里读取的附件文件有一个word文件和一个ppt文件，均在本地。

```python
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# 准备MIME容器报文
multiMsg = MIMEMultipart()

# 1. 二进制方式读入word附件
attachData = open(file="resources/attach/演示word附件1.docx", mode="rb").read()

# 2. 创建MIME子报文
# _data: 添加读入的二进制数据
# _subtype: 查表，设置好word文档的子类型
attachMsg = MIMEApplication(_data=attachData, _subtype="msword")
# 添加扩展字段Content-Disposition：设置显示方式和浏览器默认呈现的附件名称
attachMsg.add_header(_name="Content-Disposition", _value="attachment", filename="演示word附件1.docx")

# 3. 将创建好的报文添加到multipart容器报文中
multiMsg.attach(payload=attachMsg)

# 类似的过程添加ppt附件
attachData = open(file="resources/attach/演示ppt附件2.pptx", mode="rb").read()
attachMsg = MIMEApplication(_data=attachData, _subtype="vnd.ms-powerpoint")
attachMsg.add_header(_name="Content-Disposition", _value="attachment", filename="演示ppt附件2.pptx")
multiMsg.attach(payload=attachMsg)
```

***

#### 2.5.2 补充说明：遇到陌生后缀名文件的Content-Type，怎么查表

提供2种文件**后缀名-MIME类型对照表**，可按需选择性查看：

* [MIME 参考手册 (w3school.com.cn)](https://www.w3school.com.cn/media/media_mimeref.asp)
* [HTTP content-type | 菜鸟教程 (runoob.com)](https://www.runoob.com/http/http-content-type.html)

对照表中，列举出了“文件后缀名”与“MIME类型（Content-Type）”的对应关系，表的结构大致均为如下：

| 文件后缀名 | MIME类型（Content-Type） |
| :--------: | :----------------------: |
|    .txt    |        text/plain        |
|   ......   |          ......          |

若想要知道某一种文件后缀名对应的MIME类型是什么，结合表对照即可。例如，2.5.1中使用到了word文档的附件（.docx或.doc）以及ppt文档的附件（.pptx或.ppt），遇到陌生的后缀名不知道Content-Type的情况下，查表得到：

| 文件后缀名 |     MIME类型（Content-Type）      |
| :--------: | :-------------------------------: |
|    .txt    |            text/plain             |
|   ......   |              ......               |
|  **.doc**  |      application/**msword**       |
|   ......   |              ......               |
|  **.ppt**  | application/**vnd.ms-powerpoint** |

这时，才得出，二者的Content-Type分别为application/**msword**和application/**vnd.ms-powerpoint**，于是上述2.5.1的演示代码中，创建**MIMEApplication**报文对象，分别指定两种附件的Content-Type的**_subtype**取值为**msword和vnd.ms-powerpoint**。

注意：您可能会发现.ppt文件的子类型有vnd.ms-powerpoint和x-ppt两种，**任选其一**即可。

> 遇到不知道该解析成什么Content-Type的文件后缀名时该怎么办

假如遇到了查表没有找到对应Content-Type的文件后缀（比如可以试试从上述两个文档中查.dat的文件后缀对应MIME类型，貌似是找不到的），这时可以使用通用的二进制流 MIME类型**application/octet-stream**，万物皆可application/octet-stream（二进制流）。

```python
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# 万物皆可application/octet-stream（通用MIME类型）

multiMsg = MIMEMultipart(_subtype="mixed")
# 未知后缀文件
with open(file="这里改成文件的路径", mode="rb") as unknownFile:
    # 读入数据
    unknownData = unknownFile.read()
    # unknownMsg = MIMEApplication(_data=unknownData, _subtype="octet-stream")
    # 默认_subtype为octet-stream，所以也可以不写_subtype参数
    unknownMsg = MIMEApplication(_data=unknownData)

multiMsg.attach(payload=unknownMsg)
```

***





## 3. 程序演示案例（邮件批量发送技术垂直原型）

> 完整源码可前往Github链接自取：
>
> 如果想在本地跑起来验证效果，请根据3.2的步骤提前准备好预备的数据，填入代码空白处。

本案例是基于QQ邮箱SMTP服务器（smtp.qq.com）实现的，其它邮箱的SMTP服务器可类比实现。**源码覆盖到了第2部分中所有的数据准备情景，可对照注释和第2部分中提到的步骤查看。**



### 3.1 需求分析（精简版）

***

某高校的一个院级组织经常需要给自己的组员群发邮件通知消息，但是邮件内容因组员而异，每次编辑邮件都要重新编辑，非常麻烦。

现希望**基于QQ邮箱服务**，开发一个邮件自动发送程序，可以**同时支持文本、内嵌资源、附件等**内容发送；在有组员必要信息的前提下，可以**支持一种模板，批量发送**。

1. 需要支持的邮件内容包含：
   * 希望支持的文本类型：普通文本和html文本
   * 希望支持的内嵌资源文件类型：图片、视频
   * 希望支持的附件文件类型：各种各样后缀的文件都能发出
2. 批量发送说明：
   * 邮件整体有模板，只需要根据组员信息在模板文件中填空即可
   * 将组员信息整合到Excel表格（后缀名.xlsx）
   * Excel表格包含表头
   * Excel表头必留字段：组员的ID、姓名、邮箱；除此之外表头往后延伸其它字段——例如延伸出昵称、部门、性别、~~男/女朋友~~（bushi...——时，程序可自动读取追加的字段，而无需修改程序本身。ID是唯一字段，不能重复，可能是工号、学号或者默认的自增编号（从1开始）。

***



### 3.2 预备工作

***

#### 3.2.1 找到QQ邮箱的SMTP服务器和服务端口号的方法

1. 进入QQ邮箱，点击**“设置”**，选择**“账户”**选项卡。

<img src="https://bl3302files.storage.live.com/y4m9368N37L6JA12R_DuauTrJqfaX2DT54QoPeh1zDcipeVXkVvXszXB2ksSolu2l-dTbaNtzVjGhTNSl2M6hy9NPwfyWrUYf-CqgU0tSisxUk5GBs_hEdXDQdxvtw4Z_oyxmCJdlBbjnzP5kUirRAzI3wSK_HYos-MDXVWDoOn8rSRgbGW4UuJjTIjuHK3c9ZU?width=660&height=139&cropmode=none" width="660" height="139" />

2. 下拉，找到各种服务这一栏选项，点击**“如何设置”**这个链接，进入帮助文档界面。

<img src="https://bl3302files.storage.live.com/y4mYsZscWCYp_ymPQUX2GD3Gyueuy3de7pAI0L5xBYoFqtqVMcBWODjjFQxMrjqu32qsrJgrtEVZjSwxZqI6RO4v2Cw4E7-1uKahhgCw5YzK_KvTul03ClBheMVFMdQZuHgGdvsNvHbxB5G8rvefvbjujtEk3osTkQi1S1mF9x_RMk2OnXmMxnwBSpATowjb8_s?width=660&height=180&cropmode=none" width="660" height="180" />

3. 点进去以后，第一眼就能找到服务器的域名和提供服务的端口号。本应用针对发送邮件，因此只用关注发送邮件服务器栏。

<img src="https://bl3302files.storage.live.com/y4mysGmDUAJH3_fR5Cf828QfR45-pYPApRdRZVyVG-x74VWmqsJwPHH_szJ-17qFBl-efKzD7kZROt-FBuI3t9mU9Q1vhv2-MJ5JvMdwzEXwFadbNoj86nWrtdfZ1UdssB3b0Yx4SO9HZ_oV6ow2NEbLvUrikNqmpQDDMbDpKcsMf0rg5fm5m57xYECCE1XmldJ?width=660&height=199&cropmode=none" width="660" height="199" />

于是找到了发送服务器的一些信息。

* 服务器：smtp.qq.com
* 端口号：465或587，这里我们选择465。
* 发送服务器要求使用SSL安全协议建立连接，所以调用smtplib协议库的时候要找一找有没有相关支持（其实是有的）

#### 3.2.2 登录用的授权码的获取过程

1. 回到3.2.1节的第2步这个位置，点击**“生成授权码”**

<img src="https://bl3302files.storage.live.com/y4mvjOq0ilPuJZIM7UinQrL_H-krSKaWG6f-f2iVcGuzOar4ScoiAos8GKKV__XG8vJB6H2nFnkVwglB-0XMtKM8CxXox_OdaUZDw6635OPW9hALs7PavNu1HDNPEtXiHxIWWPDpii4PgXq0Ib_BZVpMdxO-MH1pxUB--JtI-xD79QlzDxmcukv7SKeshzkEKYg?width=660&height=186&cropmode=none" width="660" height="186" />

2. 这个时候QQ邮箱会提醒你需要输入**手机令牌**

<img src="https://bl3302files.storage.live.com/y4mAaedzF8DZgjKaA6l9elrt3SpJShpbbhOXPJ_YMHCM8jekLYwUMzIhG2hyy70RY66GwMfcaMwqoUHdJ0Vzm18FASqFfxtXyFv3bJYO8AFKPONJx3wNG0Ykm7qKHrhoc2q8KxqRV9MmP77mX_7xVBb6rDzjwXrcISfSkq4-SAubp4gIJvc2tCczUFXhM_fEOYz?width=320&height=219&cropmode=none" width="320" height="219" />

3. 这个令牌在哪里找：**下载QQ安全中心**，找到屏幕中显示的**流水码**。

<img src="https://bl3302files.storage.live.com/y4majCgaB19YobQJF9VV5Dg6TwjeWN8Evh2eZ5wfQxIUCnbV6Hl5qprzBN9OFLMahPLxbxpTCgAKJzbNQkG2H8YAUqdxo1tRCdDJSRxBpqeUDrSeKbW9igo-Fo99PNHnnhQefvVPOCJrSrW5fFWVndZd7ubyeJxJ3_-9b4lW6wVPOB2fdXbKSBnDO6zUME4SWHu?width=323&height=660&cropmode=none" width="323" height="660" />

4. 在流水码刷新的时间内，填入当前显示的流水码到手机令牌中。

<img src="https://bl3302files.storage.live.com/y4m4SU37OswwC9-aG4ctXbEAn4PzxiQN9xd5ECCqZ_VZaNNyIxkgEgwi5cm9pwL0jtWNpfnEj6194ya7DH0slGp_6vkoIjDjF2hU4zeOFZ1iDui4gonuyf-nN_-dkdnfzc666IY0bY5pSeJgli3SGeI-b7m4SCQjNUo9DnIFEKySj_xWsgU-yhFer5KAsCe6Z1-?width=660&height=466&cropmode=none" width="660" height="466" />

5. QQ邮箱继续验证，得到一个16位**授权码**。这个授权码将**代替QQ邮箱的真实密码**写在程序当中，用于登录SMTP邮件服务器。

<img src="https://bl3302files.storage.live.com/y4mjKDem4dzXOkhMx-HqJoiFe8I9bWcvVRt1-NcMUZ4xEYA4H8wI-K8NvsZf-_GsqPV41GXETQa8Lm8EdpouL5JGTeKbP47IBVVjtFaZXb5yqE0Ygz36bvNN85P7fcJh3EbzxElADNOeUQ_Ufwq0FnRxCyWwfP3ADKBYCwfFlCkxqIK6WFeFMvfr6v48GjNgxSo?width=660&height=347&cropmode=none" width="660" height="347" />

#### 3.2.3 为实现批量发送需要准备的Excel表结构

| ID（必填） | 姓名（必填） | 邮箱（必填） | 其它字段（根据业务继续补充） |
| :--------: | :----------: | :----------: | :--------------------------: |
|            |              |              |            ......            |

Excel表文件设置

* 文件名：成员联系方式表.xlsx
* 工作表名：联系方式
  * 演示需要：扩展字段两个可选字段
    * 部门
    * 性别

Excel演示数据表展示（隐去邮箱）



***



### 3.3 程序框架

***

#### 3.3.1 邮件内容大致外观



#### 3.3.2 MIME报文结构形式展示



#### 3.3.3 批量处理



***



### 3.4 完整代码

***



***



## 4. 验证邮件的“真面目”

> 收到邮件后，阅读邮件原文（由一些MIME邮件头部和一堆乱码组成。这里的“乱码”其实是邮件体中SMTP协议传来的7位ASCII码，还未经MIME协议处理显示给用户的那种）
>
> 观察其中哪些内容是我们根据MIME协议设置过的，体会一下成果。



### 4.1 打开邮件原文

***



***



### 4.2 查看由我们设置过的核心部分

***



***





# 附：参考文献及资源链接

* 谢希仁编著《计算机网络（第7版）》—第6章 应用层—6.5 电子邮件
* [邮件传输协议SMTP和SMTPS - 楼兰胡杨 - 博客园 (cnblogs.com)](https://www.cnblogs.com/east7/p/13406089.html)
* [MIME 参考手册 (w3school.com.cn)](https://www.w3school.com.cn/media/media_mimeref.asp)
* [HTTP content-type | 菜鸟教程 (runoob.com)](https://www.runoob.com/http/http-content-type.html)
* [header中Content-Disposition的作用与使用方法 - wq9 - 博客园 (cnblogs.com)](https://www.cnblogs.com/wq-9/articles/12165056.html)
* [Content-Disposition - HTTP | MDN (mozilla.org)](https://developer.mozilla.org/zh-CN/docs/Web/HTTP/Headers/Content-Disposition)
* [王道考研 计算机网络 P72 6.4 电子邮件_哔哩哔哩 (゜-゜)つロ 干杯~-bilibili](https://www.bilibili.com/video/BV19E411D78Q?p=72)
* [MIME中的Multipart/mixed - 蹉跎错，消磨过，最是光阴化浮沫 - ITeye博客](https://www.iteye.com/blog/gaojunwei-1939073)
* [Free stock videos · Pexels Videos](https://www.pexels.com/videos/)
* [python自动发邮件(嵌入图片，带附件，html内容)-注释详细_清风徐来-CSDN博客](https://blog.csdn.net/weixin_42223090/article/details/105630310)
* [python3之模块SMTP协议客户端与email邮件MIME对象 - Py.qi - 博客园 (cnblogs.com)](https://www.cnblogs.com/zhangxinqi/p/9113859.html)

