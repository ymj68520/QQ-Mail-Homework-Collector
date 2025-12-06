import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import threading
import sys
import os
import io
import re
import imaplib
import email
from email.header import decode_header
import pandas as pd
from dotenv import load_dotenv

# ================= 工具类：重定向输出到UI =================
class IORedirector(object):
    """把 print 的内容重定向到 Text 控件中"""
    def __init__(self, text_area):
        self.text_area = text_area

    def write(self, str_val):
        # 在主线程更新UI
        self.text_area.after(0, self._insert_text, str_val)

    def _insert_text(self, str_val):
        self.text_area.configure(state='normal')
        self.text_area.insert(tk.END, str_val)
        self.text_area.see(tk.END) # 自动滚动到底部
        self.text_area.configure(state='disabled')

    def flush(self):
        pass

# ================= 主程序逻辑类 =================
class QQMailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QQ邮箱作业自动收集助手 v1.0")
        self.root.geometry("700x600")
        
        # 加载配置
        load_dotenv()
        self.config = {
            "QQ_EMAIL": tk.StringVar(value=os.getenv("QQ_EMAIL", "")),
            "QQ_PASSWORD": tk.StringVar(value=os.getenv("QQ_PASSWORD", "")),
            "TARGET_FOLDER": tk.StringVar(value=os.getenv("TARGET_FOLDER", "")),
            "SAVE_DIR": tk.StringVar(value=os.getenv("SAVE_DIR", "downloaded_attachments"))
        }

        self._init_ui()

    def _init_ui(self):
        # 1. 配置区域 Frame
        config_frame = ttk.LabelFrame(self.root, text="⚙️ 参数设置", padding=10)
        config_frame.pack(fill="x", padx=10, pady=5)

        # 邮箱账号
        ttk.Label(config_frame, text="QQ邮箱:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.config["QQ_EMAIL"], width=30).grid(row=0, column=1, padx=5)

        # 授权码
        ttk.Label(config_frame, text="授权码:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.config["QQ_PASSWORD"], width=20, show="*").grid(row=0, column=3, padx=5)

        # 文件夹/标签
        ttk.Label(config_frame, text="标签/文件夹:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.config["TARGET_FOLDER"], width=30).grid(row=1, column=1, padx=5)
        ttk.Label(config_frame, text="(输入如 '25TA')").grid(row=1, column=2, sticky="w")

        # 保存路径
        ttk.Label(config_frame, text="保存目录:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(config_frame, textvariable=self.config["SAVE_DIR"], width=30).grid(row=2, column=1, padx=5)
        ttk.Button(config_frame, text="浏览...", command=self.browse_folder).grid(row=2, column=2, padx=5)

        # 保存配置按钮
        ttk.Button(config_frame, text="💾 保存配置到 .env", command=self.save_env).grid(row=3, column=1, pady=10)

        # 2. 操作区域 Frame
        action_frame = ttk.LabelFrame(self.root, text="🚀 执行操作", padding=10)
        action_frame.pack(fill="x", padx=10, pady=5)

        self.btn_download = ttk.Button(action_frame, text="📥 开始下载附件", command=self.start_download_thread)
        self.btn_download.pack(side="left", expand=True, fill="x", padx=10)

        self.btn_analyze = ttk.Button(action_frame, text="📊 生成统计表格", command=self.start_analyze_thread)
        self.btn_analyze.pack(side="left", expand=True, fill="x", padx=10)

        self.btn_open_excel = ttk.Button(action_frame, text="📂 打开统计表", command=self.open_excel_file)
        self.btn_open_excel.pack(side="left", expand=True, fill="x", padx=10)

        # 3. 日志区域 Frame
        log_frame = ttk.LabelFrame(self.root, text="📝 运行日志", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state='disabled', bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)

        # 重定向输出
        sys.stdout = IORedirector(self.log_text)
        sys.stderr = IORedirector(self.log_text)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.config["SAVE_DIR"].set(folder)

    def save_env(self):
        """保存配置到 .env 文件"""
        content = (
            f"QQ_EMAIL={self.config['QQ_EMAIL'].get()}\n"
            f"QQ_PASSWORD={self.config['QQ_PASSWORD'].get()}\n"
            f"TARGET_FOLDER={self.config['TARGET_FOLDER'].get()}\n"
            f"SAVE_DIR={self.config['SAVE_DIR'].get()}\n"
        )
        try:
            with open(".env", "w", encoding="utf-8") as f:
                f.write(content)
            print("✅ 配置已保存到 .env 文件！")
        except Exception as e:
            print(f"❌ 保存失败: {e}")

    def open_excel_file(self):
        """尝试打开生成的 Excel 文件"""
        save_dir = self.config["SAVE_DIR"].get()
        # 这里的路径逻辑需要与 analyze 中保持一致
        base_dir = os.path.dirname(save_dir) if os.path.dirname(save_dir) else "."
        output_file = os.path.join(base_dir, "作业统计表.xlsx")

        if not os.path.exists(output_file):
            print(f"❌ 找不到文件: {output_file}")
            print("请先点击“生成统计表格”生成文件。")
            return

        print(f"📂 正在打开: {output_file}")
        try:
            os.startfile(output_file)
        except Exception as e:
            print(f"❌ 无法打开文件: {e}")

    # ================= 业务逻辑：下载 =================
    def start_download_thread(self):
        self.btn_download.config(state="disabled")
        self.btn_analyze.config(state="disabled")
        t = threading.Thread(target=self.run_download_logic)
        t.start()

    def run_download_logic(self):
        try:
            # 获取当前配置
            user = self.config["QQ_EMAIL"].get()
            pwd = self.config["QQ_PASSWORD"].get()
            keyword = self.config["TARGET_FOLDER"].get()
            save_dir = self.config["SAVE_DIR"].get()

            if not user or not pwd or not keyword:
                print("❌ 请先填写完整配置信息！")
                return

            print(f"\n--- 开始任务: 连接邮箱 {user} ---")
            mail = imaplib.IMAP4_SSL("imap.qq.com")
            mail.login(user, pwd)
            print("登录成功！正在搜索文件夹路径...")

            # 寻找真实路径逻辑
            status, folders = mail.list()
            real_path = None
            for f in folders:
                f_str = f.decode('utf-8', 'ignore') if isinstance(f, bytes) else str(f)
                if keyword in f_str:
                    match = re.search(r'"([^"]+)"$', f_str)
                    if match and keyword in match.group(1):
                        real_path = match.group(1)
                        break
            
            if not real_path:
                print(f"❌ 未找到包含 '{keyword}' 的文件夹。")
                mail.logout()
                return
            
            print(f"✅ 锁定目标文件夹: {real_path}")
            mail.select(f'"{real_path}"')

            status, messages = mail.search(None, "ALL")
            if status != "OK" or not messages[0]:
                print("该文件夹下没有邮件。")
                mail.logout()
                return

            email_ids = messages[0].split()
            print(f"共发现 {len(email_ids)} 封邮件。开始处理...")

            if not os.path.exists(save_dir):
                os.makedirs(save_dir)

            for mail_id in email_ids:
                try:
                    _, msg_data = mail.fetch(mail_id, "(RFC822)")
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            subject = self._decode_str(msg["Subject"])
                            subject = self._clean_filename(subject)
                            if not subject: subject = f"无标题_{mail_id.decode()}"

                            # 创建文件夹
                            mail_folder = os.path.join(save_dir, subject)
                            processed_log = False
                            
                            for part in msg.walk():
                                if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                                    continue
                                
                                filename = part.get_filename()
                                if filename:
                                    if not processed_log:
                                        print(f"处理: {subject}")
                                        if not os.path.exists(mail_folder): os.makedirs(mail_folder)
                                        processed_log = True
                                    
                                    filename = self._clean_filename(self._decode_str(filename))
                                    filepath = os.path.join(mail_folder, filename)
                                    if not os.path.exists(filepath):
                                        with open(filepath, "wb") as f:
                                            f.write(part.get_payload(decode=True))
                                        print(f"  -> 下载: {filename}")
                except Exception as e:
                    print(f"  ! 错误: {e}")

            mail.close()
            mail.logout()
            print("\n🎉 下载任务全部完成！")

        except Exception as e:
            print(f"❌ 发生严重错误: {e}")
        finally:
            self.root.after(0, lambda: self.btn_download.config(state="normal"))
            self.root.after(0, lambda: self.btn_analyze.config(state="normal"))

    # ================= 业务逻辑：统计 =================
    def start_analyze_thread(self):
        self.btn_download.config(state="disabled")
        self.btn_analyze.config(state="disabled")
        t = threading.Thread(target=self.run_analyze_logic)
        t.start()

    def run_analyze_logic(self):
        save_dir = self.config["SAVE_DIR"].get()
        output_file = os.path.join(os.path.dirname(save_dir) if os.path.dirname(save_dir) else ".", "作业统计表.xlsx")
        
        print(f"\n--- 开始统计: 扫描目录 {save_dir} ---")
        if not os.path.exists(save_dir):
            print("❌ 目录不存在，请先下载附件。")
            self._reset_buttons()
            return

        student_map = {}
        no_id_list = []

        for folder in os.listdir(save_dir):
            folder_path = os.path.join(save_dir, folder)
            if os.path.isdir(folder_path):
                # 解析逻辑
                clean_text = re.sub(r'[+\-_,，\.。=]', ' ', folder)
                info = {"id": "", "name": "", "other": ""}
                
                # 学号
                id_match = re.search(r'\d{6,}', clean_text)
                if id_match:
                    info["id"] = id_match.group()
                    clean_text = clean_text.replace(info["id"], ' ')
                
                # 姓名
                name_match = re.search(r'[\u4e00-\u9fa5]{2,4}', clean_text)
                if name_match:
                    info["name"] = name_match.group()
                    clean_text = clean_text.replace(info["name"], ' ')
                
                info["other"] = re.sub(r'\s+', ' ', clean_text).strip()
                
                files = os.listdir(folder_path)
                
                # 计算大小
                total_size = 0
                for f in files:
                    fp = os.path.join(folder_path, f)
                    if os.path.isfile(fp):
                        total_size += os.path.getsize(fp)
                
                sid = info["id"]
                if sid:
                    if sid not in student_map:
                        student_map[sid] = {
                            "学号": sid,
                            "姓名": info["name"],
                            "作业备注": set(),
                            "来源文件夹": [],
                            "附件数量": 0,
                            "附件总大小Bytes": 0,
                            "附件列表": []
                        }
                    entry = student_map[sid]
                    if len(info["name"]) > len(entry["姓名"]):
                        entry["姓名"] = info["name"]
                    if info["other"]:
                        entry["作业备注"].add(info["other"])
                    entry["来源文件夹"].append(folder)
                    entry["附件数量"] += len(files)
                    entry["附件总大小Bytes"] += total_size
                    entry["附件列表"].extend(files)
                else:
                    no_id_list.append({
                        "学号": "",
                        "姓名": info["name"],
                        "作业备注": info["other"],
                        "来源文件夹": folder,
                        "附件数量": len(files),
                        "附件总大小": self._format_size(total_size),
                        "附件列表": "; ".join(files)
                    })

        final_data = []
        for sid, data in student_map.items():
            final_data.append({
                "学号": sid,
                "姓名": data["姓名"],
                "作业备注": "; ".join(sorted(list(data["作业备注"]))),
                "来源文件夹": "; ".join(data["来源文件夹"]),
                "附件数量": data["附件数量"],
                "附件总大小": self._format_size(data["附件总大小Bytes"]),
                "附件列表": "; ".join(data["附件列表"])
            })
        final_data.extend(no_id_list)

        if not final_data:
            print("⚠️ 未找到任何文件夹记录。")
        else:
            try:
                df = pd.DataFrame(final_data)
                df = df.sort_values(by="学号")
                
                # 使用 ExcelWriter 启用筛选功能
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    worksheet = writer.sheets['Sheet1']
                    worksheet.auto_filter.ref = worksheet.dimensions
                
                print(f"✅ 统计完成！共 {len(df)} 条数据(学生)。")
                print(f"📄 Excel 已保存至: {os.path.abspath(output_file)}")
            except Exception as e:
                print(f"❌ 保存 Excel 失败: {e}")
                print("请关闭正在打开的 Excel 文件后重试。")
        
        self._reset_buttons()

    def _reset_buttons(self):
        self.root.after(0, lambda: self.btn_download.config(state="normal"))
        self.root.after(0, lambda: self.btn_analyze.config(state="normal"))

    def _format_size(self, s):
        if s == 0: return "0B"
        units = ("B", "KB", "MB", "GB")
        i = 0
        while s >= 1024 and i < len(units) - 1:
            s /= 1024.0
            i += 1
        return f"{s:.2f} {units[i]}"

    # ================= 辅助函数 =================
    def _clean_filename(self, filename):
        if not filename: return "unknown"
        return re.sub(r'[\\/*?:"<>|]', "", filename).strip()

    def _decode_str(self, s):
        if s is None: return ""
        value, charset = decode_header(s)[0]
        if charset:
            try: return value.decode(charset)
            except: 
                try: return value.decode('gbk')
                except: return value.decode('utf-8', errors='ignore')
        else:
            if isinstance(value, bytes): return value.decode('utf-8', errors='ignore')
            return str(value)

if __name__ == "__main__":
    root = tk.Tk()
    # 尝试设置图标（如果有的话）
    # root.iconbitmap("icon.ico") 
    app = QQMailApp(root)
    root.mainloop()