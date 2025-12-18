import os
import re
import pandas as pd
from dotenv import load_dotenv
import sys
import io
from datetime import datetime
import glob
from smart_student_info_parser import smart_parse_folder_name, extract_assignment_name, get_folder_modification_time, get_submission_files_info

# ===========================================
# 解决 emoji 报错和中文乱码（仅在需要时重定向）
if hasattr(sys.stdout, 'buffer'):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except:
        pass  # 如果重定向失败，使用默认输出
# ===========================================

# ================= 配置区域 =================
# 加载环境变量
load_dotenv()
SAVE_DIR = os.getenv('SAVE_DIR', 'downloaded_attachments') # 附件保存的根目录
OUTPUT_FILE = '作业提交分析_按作业分组.xlsx'
# ===========================================

def parse_folder_name(folder_name, folder_path=None):
    """
    智能解析文件夹名称，优先使用邮件元数据
    """
    if folder_path:
        return smart_parse_folder_name(folder_path, folder_name)
    else:
        # 如果没有提供路径，使用传统解析方法
        from smart_student_info_parser import traditional_parse_folder_name
        return traditional_parse_folder_name(folder_name)

# extract_assignment_name 函数已从 smart_student_info_parser 导入

# get_folder_modification_time 函数已从 smart_student_info_parser 导入

def classify_submission_status(submissions):
    """
    根据提交时间序列分类提交状态
    """
    if len(submissions) == 1:
        return ["初交"]
    
    statuses = []
    for i, submission in enumerate(submissions):
        if i == 0:
            statuses.append("初交")
        elif i == len(submissions) - 1:
            statuses.append("最终版")
        else:
            # 检查是否包含补交、重交等关键词
            assignment_text = submission.get('assignment', '').lower()
            if any(keyword in assignment_text for keyword in ['补交', '重交', '修订', 'resubmit', 'revise']):
                statuses.append("补交/修订")
            else:
                statuses.append(f"第{i+1}次提交")
    
    return statuses

def analyze_by_assignment():
    """
    按作业分组分析多次提交情况
    """
    if not os.path.exists(SAVE_DIR):
        print(f"❌ 找不到目录: {SAVE_DIR}，请先运行下载程序。")
        return

    print(f"正在扫描目录: {SAVE_DIR} ...")
    
    # 收集所有数据
    all_submissions = []
    
    # 用于检测重复文件夹的字典
    folder_groups = {}
    
    # 遍历根目录下的所有文件夹
    for folder in os.listdir(SAVE_DIR):
        folder_path = os.path.join(SAVE_DIR, folder)
        
        if os.path.isdir(folder_path):
            # 解析文件夹名字
            parsed_info = parse_folder_name(folder, folder_path)
            
            # 统计文件信息
            files = os.listdir(folder_path)
            file_count = len(files)
            file_names = "; ".join(files)
            
            # 获取提交时间
            submit_time = get_folder_modification_time(folder_path)
            
            # 提取标准化作业名称
            assignment_name = extract_assignment_name(parsed_info["assignment"])
            
            # 创建唯一标识符（学号+姓名+作业名）
            unique_key = f"{parsed_info['student_id']}_{parsed_info['name']}_{assignment_name}"
            
            # 检测重复文件夹（处理带(1)、(2)后缀的情况）
            base_folder = re.sub(r'[（(]\d+[）)]$', '', folder)  # 移除末尾的(1)、（2）等
            
            if unique_key not in folder_groups:
                folder_groups[unique_key] = []
            
            folder_groups[unique_key].append({
                "文件夹原名": parsed_info["original_text"],
                "学号": parsed_info["student_id"],
                "姓名": parsed_info["name"],
                "作业名称": assignment_name,
                "作业备注": parsed_info["assignment"],
                "提交时间": submit_time,
                "附件数量": file_count,
                "附件列表": file_names,
                "文件夹路径": folder_path,
                "原始文件夹名": folder
            })
    
    # 处理重复文件夹，只保留最新版本
    for unique_key, submissions in folder_groups.items():
        if len(submissions) == 1:
            # 没有重复，直接添加
            all_submissions.append(submissions[0])
        else:
            # 有重复，选择最新的
            submissions.sort(key=lambda x: x['提交时间'] if x['提交时间'] != datetime.min else datetime.min)
            latest = submissions[-1]
            
            # 标记为重复文件夹
            latest["文件夹原名"] = f"{latest['原始文件夹名']} (合并自{len(submissions)}个重复文件夹)"
            all_submissions.append(latest)
            
            print(f"🔄 合并重复文件夹: {unique_key}")
            for sub in submissions:
                print(f"   - {sub['原始文件夹名']} ({sub['提交时间']})")
            print(f"   ✅ 选择: {latest['原始文件夹名']} ({latest['提交时间']})")
    
    if not all_submissions:
        print("没有找到任何记录。")
        return
    
    print(f"扫描完成，共 {len(all_submissions)} 条提交记录。")
    
    # 按作业分组
    assignment_groups = {}
    for submission in all_submissions:
        assignment_name = submission["作业名称"]
        if assignment_name not in assignment_groups:
            assignment_groups[assignment_name] = []
        assignment_groups[assignment_name].append(submission)
    
    # 创建Excel写入器
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        
        # 创建汇总表 - 只保留每个学生的最新提交
        summary_data = []
        for assignment_name, submissions in assignment_groups.items():
            # 按学号分组，统计每个学生的提交次数
            student_submissions = {}
            for sub in submissions:
                student_key = f"{sub['学号']}_{sub['姓名']}"
                if student_key not in student_submissions:
                    student_submissions[student_key] = []
                student_submissions[student_key].append(sub)
            
            # 按提交时间排序每个学生的提交，只保留最新版本
            for student_key, student_subs in student_submissions.items():
                student_subs.sort(key=lambda x: x['提交时间'] if x['提交时间'] != datetime.min else datetime.min)
                
                # 只保留最新提交
                latest_submission = student_subs[-1]
                total_submissions = len(student_subs)
                
                # 确定提交状态
                if total_submissions == 1:
                    status = "初交"
                else:
                    # 检查最新提交是否包含补交、重交等关键词
                    assignment_text = latest_submission.get('作业备注', '').lower()
                    if any(keyword in assignment_text for keyword in ['补交', '重交', '修订', 'resubmit', 'revise']):
                        status = "补交/修订"
                    else:
                        status = f"第{total_submissions}次提交"
                
                summary_data.append({
                    "作业名称": assignment_name,
                    "学号": latest_submission['学号'],
                    "姓名": latest_submission['姓名'],
                    "提交次数": f"{total_submissions}次",
                    "提交状态": status,
                    "提交时间": latest_submission['提交时间'].strftime('%Y-%m-%d %H:%M:%S'),
                    "附件数量": latest_submission['附件数量'],
                    "文件夹": latest_submission['文件夹原名']
                })
        
        # 写入汇总表
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_df = summary_df.sort_values(['作业名称', '学号', '提交时间'])
            summary_df.to_excel(writer, sheet_name='汇总表', index=False)
        else:
            # 创建空表
            pd.DataFrame(columns=['作业名称', '学号', '姓名', '提交状态', '提交时间', '附件数量', '文件夹']).to_excel(writer, sheet_name='汇总表', index=False)
        

    
    print(f"✅ 分析完成！文件已保存为: {OUTPUT_FILE}")
    print(f"📊 共分析了 {len(assignment_groups)} 个作业")
    print(f"📝 只包含汇总表")
    
    # 打印预览
    if summary_data:
        print("\n--- 汇总预览 ---")
        summary_df = pd.DataFrame(summary_data)
        print(summary_df[['作业名称', '姓名', '学号', '提交状态']].head(10).to_string(index=False))

if __name__ == "__main__":
    analyze_by_assignment()