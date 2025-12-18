import imaplib
import email
from email.header import decode_header
import os
import re
import sys
import io
import json
from datetime import datetime
from dotenv import load_dotenv
from email_content_parser import extract_email_body, combine_extraction_results, extract_info_from_subject, extract_info_from_body, extract_info_from_sender

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

def parse_email_date(date_str):
    """解析邮件日期字符串"""
    if not date_str:
        return datetime.now()
    
    try:
        # 尝试解析各种邮件日期格式
        from email.utils import parsedate_to_datetime
        dt = parsedate_to_datetime(date_str)
        return dt if dt else datetime.now()
    except:
        try:
            # 备用解析方法
            return datetime.strptime(date_str[:20], '%a, %d %b %Y %H:%M:%S')
        except:
            return datetime.now()

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

def save_metadata(folder_path, metadata):
    """保存邮件元数据到JSON文件"""
    metadata_file = os.path.join(folder_path, 'email_metadata.json')
    try:
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2, default=str)
    except Exception as e:
        print(f"  ! 保存元数据失败: {e}")

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

                    # 解析邮件日期
                    email_date = parse_email_date(msg["Date"])
                    
                    # 提取邮件正文
                    email_body = extract_email_body(msg)
                    
                    # 智能解析学生信息
                    sender_info = decode_str(msg["From"])
                    subject_result = extract_info_from_subject(subject)
                    body_result = extract_info_from_body(email_body)
                    filename_result = {}  # 暂时没有文件名信息
                    sender_result = extract_info_from_sender(sender_info)
                    
                    # 合并解析结果
                    student_info = combine_extraction_results(subject_result, body_result, filename_result, sender_result)
                    
                    # 如果解析成功，使用解析后的信息作为文件夹名
                    if student_info["confidence"] > 30:  # 置信度阈值
                        # 构建文件夹名，确保有意义
                        parts = []
                        if student_info['student_id']:
                            parts.append(student_info['student_id'])
                        if student_info['name']:
                            parts.append(student_info['name'])
                        if student_info['assignment']:
                            parts.append(student_info['assignment'])
                        
                        if parts:
                            folder_name = "_".join(parts)
                        else:
                            folder_name = subject  # 如果解析结果为空，使用原标题
                        
                        folder_name = clean_filename(folder_name)
                        if not folder_name.strip():
                            folder_name = subject  # 如果清理后为空，使用原标题
                    else:
                        folder_name = subject
                    
                    # 创建邮件同名文件夹
                    mail_folder = os.path.join(SAVE_DIR, folder_name)
                    
                    # 准备增强元数据
                    metadata = {
                        "邮件ID": mail_id.decode(),
                        "原始主题": subject,
                        "文件夹名称": folder_name,
                        "发件人": sender_info,
                        "收件人": decode_str(msg["To"]),
                        "发送时间": email_date.isoformat(),
                        "接收时间": parse_email_date(msg["Received"]).isoformat() if msg["Received"] else "",
                        "邮件正文": email_body if email_body else "",  # 保存完整正文
                        "解析信息": student_info,
                        "附件数量": 0,
                        "附件列表": []
                    }
                    
                    # 先创建文件夹（即使没有附件也要创建）
                    print(f"处理邮件: {subject}")
                    if not os.path.exists(mail_folder):
                        os.makedirs(mail_folder)
                    
                    processed_log = False 

                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart': continue
                        if part.get('Content-Disposition') is None: continue

                        filename = part.get_filename()
                        if filename:
                            if not processed_log:
                                processed_log = True

                            filename = decode_str(filename)
                            filename = clean_filename(filename)
                            filepath = os.path.join(mail_folder, filename)
                            
                            # 保存附件信息到元数据
                            attachment_info = {
                                "文件名": filename,
                                "大小": len(part.get_payload(decode=True)),
                                "类型": part.get_content_type(),
                                "创建时间": email_date.isoformat()
                            }
                            metadata["附件列表"].append(attachment_info)
                            metadata["附件数量"] += 1
                            
                            if not os.path.exists(filepath):
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"  |-- 下载附件: {filename}")
                            else:
                                print(f"  |-- 跳过重复: {filename}")
                    
                    # 保存元数据文件（总是保存，即使没有附件）
                    save_metadata(mail_folder, metadata)
                    
        except Exception as e:
            print(f"  ! 处理邮件出错: {e}")
            continue

    mail.close()
    mail.logout()
    print("\n所有任务完成！")
    print("💾 已为每个邮件文件夹创建了元数据文件 (email_metadata.json)")

if __name__ == "__main__":
    download_attachments()