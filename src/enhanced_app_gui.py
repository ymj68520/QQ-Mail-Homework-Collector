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
import subprocess

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
class EnhancedQQMailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QQ邮箱作业自动收集助手 v2.0 - 增强版")
        self.root.geometry("800x700")
        
        # 加载配置
        load_dotenv()
        self.config = {
            "QQ_EMAIL": tk.StringVar(value=os.getenv("QQ_EMAIL", "")),
            "QQ_PASSWORD": tk.StringVar(value=os.getenv("QQ_PASSWORD", "")),
            "TARGET_FOLDER": tk.StringVar(value=os.getenv("TARGET_FOLDER", "")),
            "SAVE_DIR": tk.StringVar(value=os.getenv("SAVE_DIR", "downloaded_attachments"))
        }

        # 分析模式选择
        self.analysis_mode = tk.StringVar(value="basic")
        
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

        # 2. 下载操作区域 Frame
        download_frame = ttk.LabelFrame(self.root, text="📥 下载操作", padding=10)
        download_frame.pack(fill="x", padx=10, pady=5)

        self.btn_download_basic = ttk.Button(download_frame, text="📥 标准下载附件", command=self.start_download_thread)
        self.btn_download_basic.pack(side="left", expand=True, fill="x", padx=5)

        self.btn_download_enhanced = ttk.Button(download_frame, text="📥 增强下载附件(含元数据)", command=self.start_enhanced_download_thread)
        self.btn_download_enhanced.pack(side="left", expand=True, fill="x", padx=5)

        # 3. 分析操作区域 Frame
        analysis_frame = ttk.LabelFrame(self.root, text="📊 分析模式", padding=10)
        analysis_frame.pack(fill="x", padx=10, pady=5)

        # 分析模式选择
        mode_frame = ttk.Frame(analysis_frame)
        mode_frame.pack(fill="x", pady=5)
        
        ttk.Label(mode_frame, text="选择分析模式:").pack(side="left", padx=5)
        
        modes = [
            ("基础统计", "basic"),
            ("模式一：一次作业多次提交分析", "multi_submission"),
            ("模式二：多个作业综合分析", "multi_assignment"),
            ("全部模式", "all")
        ]
        
        for text, value in modes:
            ttk.Radiobutton(mode_frame, text=text, variable=self.analysis_mode, 
                           value=value).pack(side="left", padx=10)
        
        # 解析模式选择
        parse_frame = ttk.Frame(analysis_frame)
        parse_frame.pack(fill="x", pady=5)
        
        self.parse_mode = tk.StringVar(value="smart")
        ttk.Label(parse_frame, text="解析模式:").pack(side="left", padx=5)
        
        parse_modes = [
            ("智能解析（推荐）", "smart"),
            ("传统解析", "traditional")
        ]
        
        for text, value in parse_modes:
            ttk.Radiobutton(parse_frame, text=text, variable=self.parse_mode, 
                           value=value).pack(side="left", padx=10)

        # 分析按钮
        button_frame = ttk.Frame(analysis_frame)
        button_frame.pack(fill="x", pady=5)
        
        self.btn_analyze = ttk.Button(button_frame, text="📊 开始分析", command=self.start_analyze_thread)
        self.btn_analyze.pack(side="left", expand=True, fill="x", padx=5)

        self.btn_preview = ttk.Button(button_frame, text="👁️ 预览结果", command=self.preview_results)
        self.btn_preview.pack(side="left", expand=True, fill="x", padx=5)

        # 4. 功能说明区域
        info_frame = ttk.LabelFrame(self.root, text="ℹ️ 功能说明", padding=10)
        info_frame.pack(fill="x", padx=10, pady=5)
        
        info_text = """
• 基础统计：生成传统的作业统计表
• 模式一：分析同一作业的多次提交情况，识别初交、补交、最终版等
• 模式二：按学生维度分析多个作业的完成情况，生成完成矩阵和统计报告
• 全部模式：同时运行所有分析模式，生成完整的分析报告
        """
        
        info_label = ttk.Label(info_frame, text=info_text.strip(), justify="left")
        info_label.pack(anchor="w")

        # 5. 日志区域 Frame
        log_frame = ttk.LabelFrame(self.root, text="📝 运行日志", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', 
                                                  bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10))
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

    def start_download_thread(self):
        self.btn_download_basic.config(state="disabled")
        self.btn_download_enhanced.config(state="disabled")
        self.btn_analyze.config(state="disabled")
        t = threading.Thread(target=self.run_download_logic, args=(False,))
        t.start()

    def start_enhanced_download_thread(self):
        self.btn_download_basic.config(state="disabled")
        self.btn_download_enhanced.config(state="disabled")
        self.btn_analyze.config(state="disabled")
        t = threading.Thread(target=self.run_download_logic, args=(True,))
        t.start()

    def run_download_logic(self, enhanced=False):
        try:
            # 获取当前配置
            user = self.config["QQ_EMAIL"].get()
            pwd = self.config["QQ_PASSWORD"].get()
            keyword = self.config["TARGET_FOLDER"].get()
            save_dir = self.config["SAVE_DIR"].get()

            if not user or not pwd or not keyword:
                print("❌ 请先填写完整配置信息！")
                return

            script_name = ".\EnhancedDownloadQQAttachments.py" if enhanced else "DownloadQQAttachments.py"
            print(f"\n--- 开始任务: 使用 {script_name} 连接邮箱 {user} ---")
            
            # 运行相应的下载脚本
            try:
                result = subprocess.run([sys.executable, script_name], 
                                      capture_output=False, text=True, 
                                      cwd=os.getcwd())
                if result.returncode == 0:
                    print("✅ 下载任务完成！")
                else:
                    print(f"❌ 下载脚本执行失败，返回码: {result.returncode}")
            except Exception as e:
                print(f"❌ 执行下载脚本时出错: {e}")

        except Exception as e:
            print(f"❌ 发生严重错误: {e}")
        finally:
            self.root.after(0, self._reset_buttons)

    def start_analyze_thread(self):
        self.btn_download_basic.config(state="disabled")
        self.btn_download_enhanced.config(state="disabled")
        self.btn_analyze.config(state="disabled")
        t = threading.Thread(target=self.run_analyze_logic)
        t.start()

    def run_analyze_logic(self):
        try:
            mode = self.analysis_mode.get()
            parse_mode = self.parse_mode.get()
            save_dir = self.config["SAVE_DIR"].get()
            
            if not os.path.exists(save_dir):
                print("❌ 下载目录不存在，请先下载附件。")
                self._reset_buttons()
                return

            # 设置解析模式环境变量
            os.environ['PARSE_MODE'] = parse_mode
            
            print(f"\n--- 开始分析: 模式 = {mode}, 解析模式 = {parse_mode} ---")
            
            if mode == "basic":
                self.run_script(".\StatisticsAttachmentDetails.py")
            elif mode == "multi_submission":
                self.run_script(".\MultiSubmissionAnalyzer.py")
            elif mode == "multi_assignment":
                self.run_script(".\MultiAssignmentAnalyzer.py")
            elif mode == "all":
                print("🔄 运行所有分析模式...")
                scripts = [
                    (".\StatisticsAttachmentDetails.py", "基础统计"),
                    (".\MultiSubmissionAnalyzer.py", "模式一：一次作业多次提交分析"),
                    (".\MultiAssignmentAnalyzer.py", "模式二：多个作业综合分析")
                ]
                for script, desc in scripts:
                    print(f"\n--- 正在运行: {desc} ---")
                    self.run_script(script)
                print("✅ 所有分析模式运行完成！")
            
        except Exception as e:
            print(f"❌ 分析过程出错: {e}")
        finally:
            self._reset_buttons()

    def run_script(self, script_name):
        """运行指定的Python脚本"""
        try:
            result = subprocess.run([sys.executable, script_name], 
                                  capture_output=False, text=True, 
                                  cwd=os.getcwd())
            if result.returncode == 0:
                print(f"✅ {script_name} 运行完成")
            else:
                print(f"❌ {script_name} 运行失败，返回码: {result.returncode}")
        except Exception as e:
            print(f"❌ 运行 {script_name} 时出错: {e}")

    def preview_results(self):
        """预览分析结果"""
        try:
            # 查找生成的Excel文件
            excel_files = []
            for file in os.listdir('.'):
                if file.endswith('.xlsx') and ('作业' in file or '统计' in file):
                    excel_files.append(file)
            
            if not excel_files:
                messagebox.showinfo("提示", "未找到分析结果文件，请先运行分析。")
                return
            
            # 创建预览窗口
            preview_window = tk.Toplevel(self.root)
            preview_window.title("分析结果预览")
            preview_window.geometry("600x400")
            
            # 文件选择
            file_frame = ttk.Frame(preview_window)
            file_frame.pack(fill="x", padx=10, pady=5)
            
            ttk.Label(file_frame, text="选择文件:").pack(side="left", padx=5)
            
            file_var = tk.StringVar(value=excel_files[0])
            file_combo = ttk.Combobox(file_frame, textvariable=file_var, values=excel_files, state="readonly")
            file_combo.pack(side="left", padx=5)
            
            # 预览区域
            preview_text = scrolledtext.ScrolledText(preview_window, height=15, state='disabled')
            preview_text.pack(fill="both", expand=True, padx=10, pady=5)
            
            def load_preview():
                try:
                    df = pd.read_excel(file_var.get())
                    preview_text.configure(state='normal')
                    preview_text.delete(1.0, tk.END)
                    preview_text.insert(tk.END, f"文件: {file_var.get()}\n")
                    preview_text.insert(tk.END, f"行数: {len(df)}, 列数: {len(df.columns)}\n")
                    preview_text.insert(tk.END, "="*50 + "\n\n")
                    preview_text.insert(tk.END, df.head(10).to_string(index=False))
                    if len(df) > 10:
                        preview_text.insert(tk.END, f"\n\n... 还有 {len(df)-10} 行数据")
                    preview_text.configure(state='disabled')
                except Exception as e:
                    messagebox.showerror("错误", f"读取文件失败: {e}")
            
            ttk.Button(file_frame, text="刷新预览", command=load_preview).pack(side="left", padx=5)
            
            # 初始加载
            load_preview()
            
        except Exception as e:
            messagebox.showerror("错误", f"预览功能出错: {e}")

    def _reset_buttons(self):
        self.root.after(0, lambda: self.btn_download_basic.config(state="normal"))
        self.root.after(0, lambda: self.btn_download_enhanced.config(state="normal"))
        self.root.after(0, lambda: self.btn_analyze.config(state="normal"))

if __name__ == "__main__":
    root = tk.Tk()
    app = EnhancedQQMailApp(root)
    root.mainloop()