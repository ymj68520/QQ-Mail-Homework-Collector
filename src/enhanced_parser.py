import os
import re
import json
from datetime import datetime
from typing import Dict, List, Optional, Tuple

def parse_folder_name(folder_name: str) -> Dict[str, str]:
    """
    从文件夹名字中提取 [学号, 姓名, 作业信息]
    """
    info = {
        "original_text": folder_name,
        "student_id": "",
        "name": "",
        "assignment": ""
    }

    # 1. 预处理：把各种分隔符替换成空格
    clean_text = re.sub(r'[+\-_,，\.。=]', ' ', folder_name)
    
    # 2. 提取学号 (通常是连续的数字，假设至少6位)
    id_match = re.search(r'\d{6,}', clean_text)
    if id_match:
        info["student_id"] = id_match.group()
        clean_text = clean_text.replace(info["student_id"], ' ')
    
    # 3. 提取姓名 (假设是2-4个中文字符)
    name_match = re.search(r'[\u4e00-\u9fa5]{2,4}', clean_text)
    if name_match:
        info["name"] = name_match.group()
        clean_text = clean_text.replace(info["name"], ' ')
    
    # 4. 剩下的内容就是"作业名"
    remaining = re.sub(r'\s+', ' ', clean_text).strip()
    info["assignment"] = remaining
    return info

def extract_assignment_name(assignment_text: str) -> str:
    """
    从作业信息中提取标准化的作业名称
    """
    if not assignment_text:
        return "未知作业"
    
    # 常见作业模式匹配
    patterns = [
        r'第[一二三四五六七八九十\d]+次作业',
        r'作业[一二三四五六七八九十\d]+',
        r'实验[一二三四五六七八九十\d]+',
        r'project\d*',
        r'lab\d*',
        r'assignment\d*',
        r'hw\d*',
        r'补交',
        r'重交',
        r'修订'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, assignment_text, re.IGNORECASE)
        if match:
            return match.group()
    
    # 如果没有匹配到标准模式，返回前10个字符作为作业标识
    return assignment_text[:10] if len(assignment_text) > 10 else assignment_text

def get_email_metadata(folder_path: str) -> Optional[Dict]:
    """
    读取邮件元数据文件
    """
    metadata_file = os.path.join(folder_path, 'email_metadata.json')
    if os.path.exists(metadata_file):
        try:
            with open(metadata_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"读取元数据文件失败 {metadata_file}: {e}")
    return None

def get_folder_modification_time(folder_path: str) -> datetime:
    """
    获取文件夹的提交时间，优先使用邮件元数据中的时间
    """
    # 首先尝试从元数据文件获取时间
    metadata = get_email_metadata(folder_path)
    if metadata and "发送时间" in metadata:
        try:
            return datetime.fromisoformat(metadata["发送时间"])
        except:
            pass
    
    # 如果元数据不可用，使用文件系统时间
    if not os.path.exists(folder_path):
        return datetime.min
    
    latest_time = datetime.min
    try:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file == 'email_metadata.json':
                    continue  # 跳过元数据文件
                file_path = os.path.join(root, file)
                file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                if file_time > latest_time:
                    latest_time = file_time
    except:
        # 如果出错，使用文件夹本身的修改时间
        try:
            latest_time = datetime.fromtimestamp(os.path.getmtime(folder_path))
        except:
            latest_time = datetime.min
    
    return latest_time

def get_submission_files_info(folder_path: str) -> Tuple[int, List[str], List[Dict]]:
    """
    获取提交文件的信息，包括文件名、大小等
    """
    if not os.path.exists(folder_path):
        return 0, [], []
    
    files = []
    file_details = []
    
    # 首先尝试从元数据获取文件信息
    metadata = get_email_metadata(folder_path)
    if metadata and "附件列表" in metadata:
        for attachment in metadata["附件列表"]:
            files.append(attachment["文件名"])
            file_details.append({
                "文件名": attachment["文件名"],
                "大小": attachment.get("大小", 0),
                "类型": attachment.get("类型", "unknown")
            })
    else:
        # 从文件系统获取信息
        for item in os.listdir(folder_path):
            if item == 'email_metadata.json':
                continue
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path):
                files.append(item)
                file_details.append({
                    "文件名": item,
                    "大小": os.path.getsize(item_path),
                    "类型": "unknown"
                })
    
    return len(files), files, file_details

def analyze_submission_quality(folder_path: str) -> Dict[str, any]:
    """
    分析提交质量，包括文件类型、大小等
    """
    file_count, file_names, file_details = get_submission_files_info(folder_path)
    
    quality_info = {
        "文件数量": file_count,
        "总大小": sum(detail["大小"] for detail in file_details),
        "文件类型分布": {},
        "可疑文件": [],
        "质量评分": 0
    }
    
    # 分析文件类型
    for detail in file_details:
        ext = os.path.splitext(detail["文件名"])[1].lower()
        if ext:
            quality_info["文件类型分布"][ext] = quality_info["文件类型分布"].get(ext, 0) + 1
    
    # 检查可疑文件（空文件、过小文件等）
    for detail in file_details:
        if detail["大小"] == 0:
            quality_info["可疑文件"].append(f"{detail['文件名']} (空文件)")
        elif detail["大小"] < 1024:  # 小于1KB
            quality_info["可疑文件"].append(f"{detail['文件名']} (文件过小)")
    
    # 简单的质量评分（基于文件数量和大小）
    score = 0
    if file_count > 0:
        score += 20  # 有文件
        if file_count >= 2:
            score += 20  # 多个文件
        if quality_info["总大小"] > 10240:  # 大于10KB
            score += 30  # 文件大小合理
        if not quality_info["可疑文件"]:
            score += 30  # 无可疑文件
    
    quality_info["质量评分"] = min(score, 100)
    
    return quality_info

def classify_submission_type(folder_name: str, assignment_text: str) -> str:
    """
    根据文件夹名称和作业信息分类提交类型
    """
    folder_lower = folder_name.lower()
    assignment_lower = assignment_text.lower()
    
    # 检查补交、重交等关键词
    if any(keyword in folder_lower or keyword in assignment_lower for keyword in ['补交', '重交', '修订', 'resubmit', 'revise', 'makeup']):
        return "补交/修订"
    
    # 检查是否为迟交（通过时间判断，需要结合截止日期信息）
    # 这里可以扩展，需要作业截止日期信息
    
    return "正常提交"

def get_student_identifier(student_id: str, name: str) -> str:
    """
    生成学生唯一标识符
    """
    if student_id:
        return f"{student_id}_{name}" if name else student_id
    elif name:
        return name
    else:
        return "未知学生"

def normalize_assignment_name(assignment_text: str) -> str:
    """
    标准化作业名称，用于分组
    """
    if not assignment_text:
        return "未知作业"
    
    # 移除多余空格和特殊字符
    normalized = re.sub(r'\s+', ' ', assignment_text.strip())
    
    # 统一常见作业名称格式
    replacements = {
        r'第([一二三四五六七八九十\d]+)次作业': r'第\1次作业',
        r'作业([一二三四五六七八九十\d]+)': r'第\1次作业',
        r'实验([一二三四五六七八九十\d]+)': r'实验\1',
        r'project\s*(\d+)': r'Project\1',
        r'lab\s*(\d+)': r'Lab\1',
        r'assignment\s*(\d+)': r'Assignment\1',
        r'hw\s*(\d+)': r'HW\1',
    }
    
    for pattern, replacement in replacements:
        normalized = re.sub(pattern, replacement, normalized, flags=re.IGNORECASE)
    
    return normalized