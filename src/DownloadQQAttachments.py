import imaplib
import email
from email.header import decode_header
import os
import re
import sys
import io
from dotenv import load_dotenv  # 导入dotenv库

# ================= 配置加载区域 =================
# 1. 加载 .env 文件
load_dotenv()

# 2. 读取环境变量
EMAIL_USER = os.getenv('QQ_EMAIL')
EMAIL_PASS = os.getenv('QQ_PASSWORD')
TARGET_FOLDER_KEYWORD = os.getenv('TARGET_FOLDER')
SAVE_DIR = os.getenv('SAVE_DIR', 'downloaded_attachments') # 如果没填，默认使用后面的值

# 3. 检查配置是否读取成功
if not EMAIL_USER or not EMAIL_PASS or not TARGET_FOLDER_KEYWORD:
    print("❌ 错误：未读取到配置信息。")
    print("请确保你已创建 '.env' 文件，并包含 QQ_EMAIL, QQ_PASSWORD, TARGET_FOLDER 字段。")
    sys.exit(1)
# ===========================================

# 解决 Windows 控制台打印乱码问题
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def clean_filename(filename):
    """清洗文件名，去除非法字符"""
    if not filename: return "unknown"
    return re.sub(r'[\\/*?:"<>|]', "", filename).strip()

def decode_str(s):
    """解码邮件字符串"""
    if s is None: return ""
    value, charset = decode_header(s)[0]
    if charset:
        try:
            return value.decode(charset)
        except:
            try: return value.decode('gbk')
            except: return value.decode('utf-8', errors='ignore')
    else:
        if isinstance(value, bytes):
            return value.decode('utf-8', errors='ignore')
        return str(value)

def find_real_folder_path(mail, keyword):
    """
    核心功能：遍历所有文件夹，寻找包含关键字的真实路径
    """
    print(f"正在服务器上查找包含 '{keyword}' 的文件夹...")
    status, folders = mail.list()
    
    match_folder = None
    
    for f in folders:
        try:
            f_str = f.decode('utf-8')
        except:
            f_str = str(f)
            
        if keyword in f_str:
            # 提取双引号中的内容作为真实路径
            match = re.search(r'"([^"]+)"$', f_str)
            if match:
                full_path = match.group(1)
                # 再次确认 keyword 确实在路径里
                if keyword in full_path:
                    match_folder = full_path
                    break 
    
    return match_folder

def download_attachments():
    print(f"正在连接 QQ 邮箱服务器 (用户: {EMAIL_USER})...")
    try:
        mail = imaplib.IMAP4_SSL("imap.qq.com")
        mail.login(EMAIL_USER, EMAIL_PASS)
        print("登录成功！")
    except Exception as e:
        print(f"登录失败: {e}")
        print("请检查 .env 文件中的账号和授权码是否正确。")
        return

    # --- 第一步：自动寻找真实文件夹路径 ---
    real_folder_path = find_real_folder_path(mail, TARGET_FOLDER_KEYWORD)
    
    if real_folder_path:
        print(f"✅ 找到文件夹！")
        print(f"   输入关键字: {TARGET_FOLDER_KEYWORD}")
        print(f"   真实路径: {real_folder_path}")
        
        try:
            # 尝试选中该文件夹
            resp, _ = mail.select(f'"{real_folder_path}"')
            if resp != 'OK':
                print(f"❌ 选中文件夹失败，服务器返回: {resp}")
                return
        except Exception as e:
            print(f"❌ 选中文件夹出错: {e}")
            return
    else:
        print(f"❌ 未找到包含 '{TARGET_FOLDER_KEYWORD}' 的文件夹。")
        print("请检查 .env 中的 TARGET_FOLDER 设置。")
        return

    # --- 第二步：搜索邮件 ---
    print(f"正在搜索 '{real_folder_path}' 中的所有邮件...")
    status, messages = mail.search(None, "ALL")
    
    if status != "OK" or not messages[0]:
        print("该文件夹下没有邮件。")
        return

    email_ids = messages[0].split()
    print(f"共找到 {len(email_ids)} 封邮件。开始下载...")

    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    # --- 第三步：遍历下载 ---
    for mail_id in email_ids:
        try:
            _, msg_data = mail.fetch(mail_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject = decode_str(msg["Subject"])
                    subject = clean_filename(subject)
                    
                    if not subject: subject = f"无标题邮件_{mail_id.decode()}"

                    # 创建邮件同名文件夹
                    mail_folder = os.path.join(SAVE_DIR, subject)
                    
                    processed_log = False 

                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart': continue
                        if part.get('Content-Disposition') is None: continue

                        filename = part.get_filename()
                        if filename:
                            if not processed_log:
                                print(f"处理邮件: {subject}")
                                if not os.path.exists(mail_folder):
                                    os.makedirs(mail_folder)
                                processed_log = True

                            filename = decode_str(filename)
                            filename = clean_filename(filename)
                            filepath = os.path.join(mail_folder, filename)
                            
                            if not os.path.exists(filepath):
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"  |-- 下载附件: {filename}")
                            else:
                                print(f"  |-- 跳过重复: {filename}")
                    
        except Exception as e:
            print(f"  ! 处理邮件出错: {e}")
            continue

    mail.close()
    mail.logout()
    print("\n所有任务完成！")

if __name__ == "__main__":
    download_attachments()