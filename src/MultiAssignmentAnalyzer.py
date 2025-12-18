import os
import re
import pandas as pd
from dotenv import load_dotenv
import sys
import io
from datetime import datetime
import numpy as np
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
OUTPUT_FILE = '作业完成分析_按学生分组.xlsx'
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

def analyze_by_student():
    """
    按学生分组分析多个作业的完成情况
    """
    if not os.path.exists(SAVE_DIR):
        print(f"❌ 找不到目录: {SAVE_DIR}，请先运行下载程序。")
        return

    print(f"正在扫描目录: {SAVE_DIR} ...")
    
    # 收集所有数据
    all_submissions = []
    
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
            
            submission = {
                "文件夹原名": parsed_info["original_text"],
                "学号": parsed_info["student_id"],
                "姓名": parsed_info["name"],
                "作业名称": assignment_name,
                "作业备注": parsed_info["assignment"],
                "提交时间": submit_time,
                "附件数量": file_count,
                "附件列表": file_names
            }
            
            all_submissions.append(submission)
    
    if not all_submissions:
        print("没有找到任何记录。")
        return
    
    print(f"扫描完成，共 {len(all_submissions)} 条提交记录。")
    
    # 获取所有作业列表
    all_assignments = list(set(sub['作业名称'] for sub in all_submissions))
    all_assignments.sort()
    
    # 按学生分组
    student_groups = {}
    for submission in all_submissions:
        student_id = submission['学号']
        student_name = submission['姓名']
        
        if not student_id:  # 如果没有学号，跳过
            continue
            
        student_key = f"{student_id}_{student_name}"
        if student_key not in student_groups:
            student_groups[student_key] = {
                '学号': student_id,
                '姓名': student_name,
                '作业': {}
            }
        
        student_groups[student_key]['作业'][submission['作业名称']] = submission
    
    print(f"发现 {len(student_groups)} 名学生，{len(all_assignments)} 个作业")
    
    # 创建Excel写入器
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        
        # 1. 创建学生作业完成矩阵
        matrix_data = []
        for student_key, student_data in student_groups.items():
            row = {
                '学号': student_data['学号'],
                '姓名': student_data['姓名']
            }
            
            completed_count = 0
            total_files = 0
            
            for assignment in all_assignments:
                if assignment in student_data['作业']:
                    submission = student_data['作业'][assignment]
                    row[assignment] = f"✓ ({submission['附件数量']}文件)"
                    completed_count += 1
                    total_files += submission['附件数量']
                else:
                    row[assignment] = "✗ 未交"
            
            row['完成作业数'] = completed_count
            row['总作业数'] = len(all_assignments)
            row['完成率'] = f"{completed_count/len(all_assignments)*100:.1f}%"
            row['总文件数'] = total_files
            
            matrix_data.append(row)
        
        # 按学号排序
        matrix_df = pd.DataFrame(matrix_data)
        matrix_df = matrix_df.sort_values('学号')
        matrix_df.to_excel(writer, sheet_name='作业完成矩阵', index=False)
        
        # 2. 创建学生详细报告
        detailed_data = []
        for student_key, student_data in student_groups.items():
            for assignment_name, submission in student_data['作业'].items():
                detailed_data.append({
                    '学号': student_data['学号'],
                    '姓名': student_data['姓名'],
                    '作业名称': assignment_name,
                    '提交时间': submission['提交时间'].strftime('%Y-%m-%d %H:%M:%S'),
                    '附件数量': submission['附件数量'],
                    '附件列表': submission['附件列表'],
                    '文件夹原名': submission['文件夹原名'],
                    '作业备注': submission['作业备注']
                })
        
        detailed_df = pd.DataFrame(detailed_data)
        detailed_df = detailed_df.sort_values(['学号', '作业名称'])
        detailed_df.to_excel(writer, sheet_name='学生详细报告', index=False)
        
        # 3. 创建作业统计报告
        assignment_stats = []
        for assignment in all_assignments:
            submitted_students = 0
            total_files = 0
            submission_times = []
            
            for student_data in student_groups.values():
                if assignment in student_data['作业']:
                    submitted_students += 1
                    submission = student_data['作业'][assignment]
                    total_files += submission['附件数量']
                    submission_times.append(submission['提交时间'])
            
            total_students = len(student_groups)
            completion_rate = submitted_students / total_students * 100 if total_students > 0 else 0
            avg_files = total_files / submitted_students if submitted_students > 0 else 0
            
            # 计算提交时间统计
            if submission_times:
                earliest = min(submission_times)
                latest = max(submission_times)
                avg_time = datetime.fromtimestamp(sum(t.timestamp() for t in submission_times) / len(submission_times))
            else:
                earliest = latest = avg_time = None
            
            assignment_stats.append({
                '作业名称': assignment,
                '应交人数': total_students,
                '实交人数': submitted_students,
                '完成率': f"{completion_rate:.1f}%",
                '缺交人数': total_students - submitted_students,
                '总文件数': total_files,
                '平均文件数': f"{avg_files:.1f}",
                '最早提交': earliest.strftime('%Y-%m-%d %H:%M') if earliest else '-',
                '最晚提交': latest.strftime('%Y-%m-%d %H:%M') if latest else '-',
                '平均提交时间': avg_time.strftime('%Y-%m-%d %H:%M') if avg_time else '-'
            })
        
        stats_df = pd.DataFrame(assignment_stats)
        stats_df = stats_df.sort_values('作业名称')
        stats_df.to_excel(writer, sheet_name='作业统计报告', index=False)
        
        # 4. 创建缺交学生名单
        missing_data = []
        for student_key, student_data in student_groups.items():
            missing_assignments = []
            for assignment in all_assignments:
                if assignment not in student_data['作业']:
                    missing_assignments.append(assignment)
            
            if missing_assignments:  # 只显示有缺交的学生
                missing_data.append({
                    '学号': student_data['学号'],
                    '姓名': student_data['姓名'],
                    '缺交作业数': len(missing_assignments),
                    '缺交作业列表': '; '.join(missing_assignments),
                    '完成率': f"{(len(all_assignments) - len(missing_assignments))/len(all_assignments)*100:.1f}%"
                })
        
        if missing_data:
            missing_df = pd.DataFrame(missing_data)
            missing_df = missing_df.sort_values(['缺交作业数', '学号'], ascending=[False, True])
            missing_df.to_excel(writer, sheet_name='缺交学生名单', index=False)
        
        # 5. 创建班级整体统计
        total_students = len(student_groups)
        total_assignments = len(all_assignments)
        total_possible_submissions = total_students * total_assignments
        total_actual_submissions = sum(len(student_data['作业']) for student_data in student_groups.values())
        
        overall_stats = {
            '统计项': ['学生总数', '作业总数', '应提交总数', '实际提交总数', '整体完成率', '平均每学生完成作业数'],
            '数值': [
                total_students,
                total_assignments,
                total_possible_submissions,
                total_actual_submissions,
                f"{total_actual_submissions/total_possible_submissions*100:.1f}%" if total_possible_submissions > 0 else "0%",
                f"{total_actual_submissions/total_students:.1f}" if total_students > 0 else "0"
            ]
        }
        
        overall_df = pd.DataFrame(overall_stats)
        overall_df.to_excel(writer, sheet_name='班级整体统计', index=False)
    
    print(f"✅ 分析完成！文件已保存为: {OUTPUT_FILE}")
    print(f"📊 共分析了 {len(student_groups)} 名学生，{len(all_assignments)} 个作业")
    print(f"📝 包含工作表：作业完成矩阵、学生详细报告、作业统计报告、缺交学生名单、班级整体统计")
    
    # 打印预览
    print("\n--- 班级整体统计预览 ---")
    print(overall_df.to_string(index=False))
    
    print("\n--- 作业统计预览 ---")
    print(stats_df[['作业名称', '应交人数', '实交人数', '完成率', '缺交人数']].to_string(index=False))

if __name__ == "__main__":
    analyze_by_student()