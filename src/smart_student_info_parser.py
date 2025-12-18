import os
import re
import json
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from email_content_parser import extract_info_from_subject, extract_info_from_body, extract_info_from_filename, extract_info_from_sender, combine_extraction_results

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
            print(f"读取元数据文件失败 {metadata_file}")
    return None

def extract_info_from_attachments(folder_path: str) -> Optional[Dict]:
    """
    从附件文件名中提取学生信息
    """
    if not os.path.exists(folder_path):
        return None
    
    # 获取文件夹中的所有文件（排除元数据文件）
    files = []
    for item in os.listdir(folder_path):
        if item == 'email_metadata.json':
            continue
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            files.append(item)
    
    if not files:
        return None
    
    # 选择最佳的文件进行解析（优先选择PDF、DOC等文档文件）
    priority_files = []
    for filename in files:
        ext = os.path.splitext(filename)[1].lower()
        if ext in ['.pdf', '.doc', '.docx']:
            priority_files.append(filename)
    
    # 如果没有文档文件，使用第一个文件
    target_file = priority_files[0] if priority_files else files[0]
    
    # 使用专门的文件名解析函数
    return extract_info_from_filename_improved(target_file)

def extract_info_from_filename_improved(filename: str) -> Dict[str, any]:
    """
    专门从文件名中提取学生信息，改进版本
    """
    result = {
        "student_id": "",
        "name": "",
        "assignment": "",
        "confidence": 40,  # 文件名解析的默认置信度
        "matches": []
    }
    
    if not filename:
        return result
    
    # 移除文件扩展名
    filename = re.sub(r'\.(pdf|docx?|xlsx?|pptx?)$', '', filename, flags=re.IGNORECASE)
    
    # 1. 提取学号 - 包含容错功能
    # 首先查找所有可能的学号候选
    candidates = []
    
    # 查找13位学号（2025开头）- 包含容错模式
    for match in re.finditer(r'2025-?\d{4}-?\d{5}', filename):
        raw_id = match.group()
        clean_id = raw_id.replace('-', '')
        if len(clean_id) == 13 and clean_id.startswith('2025'):
            candidates.append(clean_id)
    
    # 查找标准13位学号（2025开头）
    for match in re.finditer(r'2025\d{9}', filename):
        candidates.append(match.group())
    
    # 查找8位学号
    for match in re.finditer(r'\d{8}', filename):
        candidate = match.group()
        # 确保不是13位学号的一部分
        if not any(candidate in long_id for long_id in candidates):
            candidates.append(candidate)
    
    # 验证候选学号
    for candidate in candidates:
        if len(candidate) == 13 and candidate.startswith('2025'):
            result["student_id"] = candidate
            result["matches"].append(f"学号匹配: {candidate}")
            break
        elif len(candidate) == 8:
            # 额外检查：确保这个8位数字不在更长的数字串中
            # 查找这个8位数字在文件名中的位置
            pos = filename.find(candidate)
            if pos != -1:
                # 检查前后是否有其他数字
                start_ok = pos == 0 or not filename[pos-1].isdigit()
                end_ok = pos + 8 == len(filename) or not filename[pos+8].isdigit()
                
                if start_ok and end_ok:
                    result["student_id"] = candidate
                    result["matches"].append(f"学号匹配: {candidate}")
                    break
    
    # 2. 提取姓名（2-4个中文字符）
    name_patterns = [
        r'^([\u4e00-\u9fa5]{2,4})\d{6,12}',  # 开头的姓名+学号
        r'^([\u4e00-\u9fa5]{2,4})[-\s]\d{6,12}',  # 开头的姓名-学号
        r'(\d{6,12})[-\s]([\u4e00-\u9fa5]{2,4})',  # 学号-姓名
        r'([\u4e00-\u9fa5]{2,4})(?=\d|$|[^\u4e00-\u9fa5])',  # 姓名后跟非中文字符
    ]
    
    for pattern in name_patterns:
        match = re.search(pattern, filename)
        if match:
            # 确定哪个组是姓名
            if pattern.startswith(r'^([\u4e00-\u9fa5'):
                name = match.group(1)
            elif pattern.startswith(r'(\d{6,12})'):
                name = match.group(2) if len(match.groups()) >= 2 else ""
            else:
                name = match.group(1) if match.groups() else ""
            
            if len(name) >= 2:
                result["name"] = name
                result["matches"].append(f"姓名匹配: {name}")
                break
    
    # 3. 提取作业名称（移除姓名和学号后的部分）
    assignment_text = filename
    
    # 移除已识别的姓名和学号
    if result["name"] and result["student_id"]:
        # 移除姓名学号的各种组合
        assignment_text = re.sub(f'^{result["name"]}{result["student_id"]}', '', assignment_text)
        assignment_text = re.sub(f'^{result["name"]}[-\\s]{result["student_id"]}[-\\s]?', '', assignment_text)
        assignment_text = re.sub(f'^{result["student_id"]}[-\\s]{result["name"]}[-\\s]?', '', assignment_text)
    
    # 移除常见的无意义后缀
    assignment_text = re.sub(r'^[-\s]+', '', assignment_text)  # 开头的横线空格
    assignment_text = re.sub(r'[-\s]+$', '', assignment_text)  # 结尾的横线空格
    
    # 使用改进的作业名提取函数
    result["assignment"] = extract_assignment_name(assignment_text)
    
    # 4. 计算置信度
    confidence = 40
    if result["student_id"]:
        confidence += 30
    if result["name"]:
        confidence += 20
    if result["assignment"] and result["assignment"] != "未知作业":
        confidence += 10
    
    result["confidence"] = min(confidence, 100)
    
    return result

def smart_parse_folder_name(folder_path: str, folder_name: str) -> Dict[str, str]:
    """
    智能解析文件夹名称，优先使用邮件元数据中的解析结果
    
    Args:
        folder_path: 文件夹路径
        folder_name: 文件夹名称
        
    Returns:
        包含解析结果的字典
    """
    import os
    
    # 检查解析模式环境变量
    parse_mode = os.environ.get('PARSE_MODE', 'smart')
    
    if parse_mode == 'traditional':
        # 强制使用传统解析
        result = traditional_parse_folder_name(folder_name)
        result["parsing_method"] = "传统解析（强制）"
        return result
    
    # 智能解析模式
    # 1. 首先尝试从附件文件名解析（最高优先级）
    attachment_info = extract_info_from_attachments(folder_path)
    if attachment_info and attachment_info.get("confidence", 0) > 30:
        return {
            "original_text": folder_name,
            "student_id": attachment_info.get("student_id", ""),
            "name": attachment_info.get("name", ""),
            "assignment": attachment_info.get("assignment", ""),
            "confidence": attachment_info.get("confidence", 0),
            "source": attachment_info.get("source", ""),
            "parsing_method": "附件解析"
        }
    
    # 2. 尝试从邮件元数据获取解析结果
    metadata = get_email_metadata(folder_path)
    if metadata and "解析信息" in metadata:
        parsed_info = metadata["解析信息"]
        if parsed_info["confidence"] > 30:  # 置信度阈值
            return {
                "original_text": folder_name,
                "student_id": parsed_info.get("student_id", ""),
                "name": parsed_info.get("name", ""),
                "assignment": parsed_info.get("assignment", ""),
                "confidence": parsed_info.get("confidence", 0),
                "source": parsed_info.get("source", ""),
                "parsing_method": "智能解析"
            }
    
    # 3. 如果都失败，使用传统方法解析文件夹名
    result = traditional_parse_folder_name(folder_name)
    result["parsing_method"] = "智能解析（回退到传统）"
    return result

def traditional_parse_folder_name(folder_name: str) -> Dict[str, str]:
    """
    传统方法解析文件夹名称
    """
    info = {
        "original_text": folder_name,
        "student_id": "",
        "name": "",
        "assignment": "",
        "confidence": 0,
        "source": "文件夹名",
        "parsing_method": "传统解析"
    }

    # 1. 预处理：把各种分隔符替换成空格
    clean_text = re.sub(r'[+\-_,，\.。=]', ' ', folder_name)
    
    # 2. 提取学号 (通常是连续的数字，假设至少6位)
    id_match = re.search(r'\d{6,}', clean_text)
    if id_match:
        info["student_id"] = id_match.group()
        clean_text = clean_text.replace(info["student_id"], ' ')
        info["confidence"] += 40
    
    # 3. 提取姓名 (假设是2-4个中文字符)
    name = extract_name_from_text(clean_text)
    if name:
        info["name"] = name
        clean_text = clean_text.replace(name, ' ')
        info["confidence"] += 35
    
    # 4. 剩下的内容就是"作业名"
    remaining = re.sub(r'\s+', ' ', clean_text).strip()
    info["assignment"] = remaining
    if remaining:
        info["confidence"] += 25
    
    return info

def extract_name_from_text(text: str) -> str:
    """
    从文本中提取姓名，过滤前缀词
    """
    if not text:
        return ""
    
    # 常见的前缀词，这些不应该被识别为姓名
    prefix_words = [
        '回复', '转发', 'Re', 'FW', 'Fwd', '回复：', '转发：', 
        '老师', '同学', '您好', '你好', '提交', '作业', '报告',
        '智能合约', '平台', '搭建', '实践', '课程', '实训',
        '最终', '期末', '大作业', '项目', '实验', '设计'
    ]
    
    # 首先尝试标准模式：2-4个中文字符
    name_patterns = [
        r'[\u4e00-\u9fa5]{2,4}',
        r'[a-zA-Z]{2,20}',  # 英文姓名
    ]
    
    for pattern in name_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            # 检查是否是前缀词
            if match not in prefix_words:
                # 进一步验证：检查是否看起来像真实姓名
                if is_valid_name(match):
                    return match
    
    # 如果没有找到合适的姓名，尝试更复杂的模式
    # 查找学号附近的姓名
    id_match = re.search(r'(\d{6,})', text)
    if id_match:
        # 在学号前后查找姓名
        id_pos = id_match.start()
        context_before = text[max(0, id_pos-20):id_pos]
        context_after = text[id_pos:id_pos+20]
        
        # 在学号前查找姓名
        for pattern in name_patterns:
            matches = re.findall(pattern, context_before)
            for match in reversed(matches):  # 从后往前找，最接近学号的
                if match not in prefix_words and is_valid_name(match):
                    return match
            
            # 在学号后查找姓名
            matches = re.findall(pattern, context_after)
            for match in matches:
                if match not in prefix_words and is_valid_name(match):
                    return match
    
    return ""

def is_valid_name(name: str) -> bool:
    """
    验证是否是有效的姓名
    """
    if not name or len(name.strip()) == 0:
        return False
    
    name = name.strip()
    
    # 中文名验证
    if re.match(r'^[\u4e00-\u9fa5]+$', name):
        # 常见的姓氏
        common_surnames = ['王', '李', '张', '刘', '陈', '杨', '赵', '黄', '周', '吴', 
                          '徐', '孙', '胡', '朱', '高', '林', '何', '郭', '马', '罗',
                          '梁', '宋', '郑', '谢', '韩', '唐', '冯', '于', '董', '萧',
                          '程', '曹', '袁', '邓', '许', '傅', '沈', '曾', '彭', '吕',
                          '苏', '卢', '蒋', '蔡', '贾', '丁', '魏', '薛', '叶', '阎']
        
        # 检查长度
        if len(name) < 2 or len(name) > 4:
            return False
        
        # 检查是否以常见姓氏开头
        if name[0] in common_surnames:
            return True
        
        # 如果不是常见姓氏，但长度合适，也可能是姓名
        if 2 <= len(name) <= 3:
            return True
        
        return False
    
    # 英文名验证
    if re.match(r'^[a-zA-Z\s]+$', name):
        # 简单的英文名验证
        name_parts = name.split()
        if len(name_parts) >= 2 and len(name_parts) <= 4:
            return True
        if len(name) >= 3 and len(name) <= 20:
            return True
        
        return False
    
    return False

def extract_assignment_name(assignment_text: str) -> str:
    """
    从作业信息中提取标准化的作业名称
    """
    if not assignment_text:
        return "未知作业"
    
    # 移除文件扩展名
    assignment_text = re.sub(r'\.(pdf|docx?|xlsx?|pptx?)$', '', assignment_text, flags=re.IGNORECASE)
    
    # 先移除学生信息（姓名+学号模式）
    assignment_text = re.sub(r'^[\u4e00-\u9fa5]{2,4}\d{6,12}', '', assignment_text)  # 姓名学号直接相连
    assignment_text = re.sub(r'^[\u4e00-\u9fa5]{2,4}[-\s]\d{6,12}[-\s]', '', assignment_text)  # 姓名-学号-
    assignment_text = re.sub(r'^[\u4e00-\u9fa5]{2,4}\s\d{6,12}\s', '', assignment_text)  # 姓名 学号 
    assignment_text = re.sub(r'^[\u4e00-\u9fa5]{2,4}[-\s]\d{6,12}', '', assignment_text)  # 姓名-学号
    
    # 移除开头的纯日期
    assignment_text = re.sub(r'^\d{4}\.\d{1,2}\.\d{1,2}', '', assignment_text)
    assignment_text = re.sub(r'^20\d{6}', '', assignment_text)  # 20251210格式
    
    # 常见作业模式匹配（按优先级排序）
    patterns = [
        # 报告类（最高优先级）
        r'最终报告',
        r'实验报告',
        r'课程报告',
        r'实践报告',
        r'项目报告',
        r'实训报告',
        
        # 作业类
        r'大作业',
        r'课程设计',
        r'智能合约',
        r'区块链',
        r'宠物游戏',
        r'奖学金',
        r'Solidity',
        r'合约设计',
        r'PayRoll',
        r'payroll',
        
        # 标准作业格式
        r'第[一二三四五六七八九十\d]+次作业',
        r'作业[一二三四五六七八九十\d]+',
        r'实验[一二三四五六七八九十\d]+',
        r'project\d*',
        r'lab\d*',
        r'assignment\d*',
        r'hw\d*',
        
        # 提交状态
        r'补交',
        r'重交',
        r'修订'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, assignment_text, re.IGNORECASE)
        if match:
            matched_text = match.group()
            # 如果匹配到的是报告类，尝试获取更完整的名称
            if '报告' in matched_text:
                # 尝试获取报告前的修饰词，但排除学生姓名
                report_pattern = r'([^\s\u4e00-\u9fa5]{2,10}报告|[^\s]{1,5}报告)'
                report_match = re.search(report_pattern, assignment_text)
                if report_match and len(report_match.group(1)) <= 15:
                    candidate = report_match.group(1)
                    # 确保不是学生姓名+报告
                    if not re.match(r'^[\u4e00-\u9fa5]{2,4}报告$', candidate):
                        return candidate
            return matched_text
    
    # 如果没有匹配到标准模式，尝试提取有意义的部分
    # 过滤掉纯日期、纯数字、括号内容等无意义文本
    filtered_text = re.sub(r'\d{4}\.\d{1,2}\.\d{1,2}$', '', assignment_text)  # 移除末尾日期
    filtered_text = re.sub(r'20\d{6}$', '', filtered_text)  # 移除末尾20251210格式
    filtered_text = re.sub(r'\(\d+\)$', '', filtered_text)  # 移除末尾(1)、(2)等
    filtered_text = re.sub(r'^[（(]\d+[）)]\s*', '', filtered_text)  # 移除开头的(1)、（2）等
    
    # 如果过滤后还有内容，返回前15个字符
    if filtered_text.strip():
        result = filtered_text.strip()[:15] if len(filtered_text.strip()) > 15 else filtered_text.strip()
        # 再次检查是否为无意义内容
        if re.match(r'^[\d\s.()（）-]*$', result):  # 如果只包含数字、空格、点、括号、横线或为空
            return "未知作业"
        return result
    
    # 最后的后备方案：检查是否为无意义内容
    if re.match(r'^[\d\s.()（）-]*$', assignment_text[:10]):  # 如果只包含数字、空格、点、括号、横线或为空
        return "未知作业"
    
    return assignment_text[:10] if len(assignment_text) > 10 else assignment_text

def get_folder_modification_time(folder_path: str) -> datetime:
    """
    获取文件夹的提交时间，优先使用邮件元数据中的时间
    """
    # 首先尝试从元数据获取时间
    metadata = get_email_metadata(folder_path)
    if metadata and "发送时间" in metadata:
        try:
            dt = datetime.fromisoformat(metadata["发送时间"])
            # 转换为无时区的datetime以保持一致性
            return dt.replace(tzinfo=None)
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
    
    for pattern, replacement in replacements.items():
        normalized = re.sub(pattern, replacement, normalized, flags=re.IGNORECASE)
    
    return normalized

def generate_parsing_report(directory: str) -> Dict[str, any]:
    """
    生成解析质量报告
    """
    if not os.path.exists(directory):
        return {"error": "目录不存在"}
    
    total_folders = 0
    smart_parsed = 0
    traditional_parsed = 0
    failed_parsed = 0
    confidence_scores = []
    
    for folder in os.listdir(directory):
        folder_path = os.path.join(directory, folder)
        if os.path.isdir(folder_path):
            total_folders += 1
            parsed_info = smart_parse_folder_name(folder_path, folder)
            
            if parsed_info["confidence"] > 70:
                smart_parsed += 1
            elif parsed_info["confidence"] > 30:
                traditional_parsed += 1
            else:
                failed_parsed += 1
            
            confidence_scores.append(parsed_info["confidence"])
    
    average_confidence = sum(confidence_scores) / len(confidence_scores) if confidence_scores else 0
    
    return {
        "总文件夹数": total_folders,
        "智能解析成功": smart_parsed,
        "传统解析成功": traditional_parsed,
        "解析失败": failed_parsed,
        "智能解析成功率": f"{smart_parsed/total_folders*100:.1f}%" if total_folders > 0 else "0%",
        "平均置信度": f"{average_confidence:.1f}",
        "高置信度比例": f"{len([c for c in confidence_scores if c > 70])/len(confidence_scores)*100:.1f}%" if confidence_scores else "0%"
    }