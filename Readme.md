# QQ Mail Homework Collector (QQ 邮箱作业自动收集助手)

这是一个基于 Python 的自动化工具，专为助教和老师设计。它可以全自动从 QQ 邮箱的指定文件夹下载所有邮件附件，并根据文件夹名称（邮件主题）智能提取**学号**和**姓名**，最后生成一份 Excel 统计报表。

## ✨ 功能特点

*   **自动下载**：通过 IMAP 协议连接 QQ 邮箱，下载指定标签/文件夹下的所有附件。
*   **智能路径识别**：自动处理 QQ 邮箱复杂的 IMAP 文件夹路径（如 `&UXZO.../25TA`），你只需输入简短的文件夹名（如 `25TA`）。
*   **隐私安全**：使用 `.env` 文件管理账号和授权码，避免敏感信息硬编码在代码中。
*   **智能归档**：以“邮件主题”为名创建文件夹，自动清洗文件名中的非法字符。
*   **统计报表**：自动扫描下载目录，智能解析“学号+姓名+作业名”的各种组合格式，生成 `作业统计表.xlsx`。
*   **中文支持**：完美解决 Windows 控制台乱码和 emoji 报错问题。

## 🛠️ 环境依赖

*   Python 3.6+
*   依赖库：`python-dotenv`, `pandas`, `openpyxl`

## 🚀 快速开始

### 1. 安装依赖

在项目目录下打开终端（Terminal），运行以下命令安装必要的库：

```bash
pip install python-dotenv pandas openpyxl
```

或者使用：

```bash
pip install -r requirements.txt
```

### 2. 获取 QQ 邮箱授权码

1. 登录QQ邮箱网页版。
2. 点击 设置 -> 账户。
3. 向下滚动，找到 POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务。
4. 开启IMAP/SMTP服务。
5. 获取16位授权码（这不是你的QQ密码，请妥善保存）。

### 3. 配置环境变量

在项目根目录下创建一个名为`.env`的文件（注意前面有个点，没有后缀名），填入以下内容：

```bash
# 你的 QQ 邮箱地址
QQ_EMAIL=12345678@qq.com

# 你的 16 位 IMAP 授权码
QQ_PASSWORD=abcdefghijklmnop

# 你在邮箱里归档作业的文件夹名称（如：25TA，无需关心乱码前缀）
TARGET_FOLDER=25TA

# 附件保存的本地目录名
SAVE_DIR=downloaded_attachments
```

### 4. 运行程序

#### 第一步：下载附件

运行下载脚本，程序会自动寻找真实路径并开始下载：

```bash
python DownloadQQAttachments.py
```

_注：请将文件名替换为你实际保存的 Python 脚本文件名_

#### 第二步：统计信息

下载完成后，运行统计脚本：

```bash
python StatisticsAttachmentDetails.py
```

运行结束后，你将在根目录下看到`作业统计表.xlsx`。

## 命名格式支持

统计脚本目前支持识别以下类型的命名组合（顺序不限）：

+ 2021001 张三 第一次作业
+ 2021001+张三+第一次作业
+ 张三-2021001-补交
+ 2021001_张三