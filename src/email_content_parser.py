import re
import email
from email.header import decode_header
from bs4 import BeautifulSoup
import html
from typing import Dict, List, Optional, Tuple

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

def extract_email_body(msg) -> str:
    """
    提取邮件正文内容，支持HTML和纯文本格式
    """
    body_text = ""
    
    if msg.is_multipart():
        # 多部分邮件，寻找text/plain或text/html部分
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = part.get('Content-Disposition', '')
            
            # 跳过附件部分
            if content_disposition and 'attachment' in content_disposition.lower():
                continue
                
            if content_type == 'text/plain':
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or 'utf-8'
                        body_text = payload.decode(charset, errors='ignore')
                        break  # 优先使用纯文本
                except:
                    continue
                    
            elif content_type == 'text/html' and not body_text:
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or 'utf-8'
                        html_content = payload.decode(charset, errors='ignore')
                        # 使用BeautifulSoup解析HTML并提取文本
                        soup = BeautifulSoup(html_content, 'html.parser')
                        body_text = soup.get_text(separator=' ', strip=True)
                except:
                    continue
    else:
        # 单部分邮件
        content_type = msg.get_content_type()
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or 'utf-8'
                if content_type == 'text/html':
                    soup = BeautifulSoup(payload.decode(charset, errors='ignore'), 'html.parser')
                    body_text = soup.get_text(separator=' ', strip=True)
                else:
                    body_text = payload.decode(charset, errors='ignore')
        except:
            body_text = str(msg.get_payload())
    
    # 清理文本
    body_text = clean_text(body_text)
    
    # 处理回复邮件，提取原始邮件信息
    body_text = extract_reply_info(body_text)
    
    return body_text

def extract_reply_info(text: str) -> str:
    """
    从回复邮件中提取原始邮件信息
    """
    if not text:
        return text
    
    # 常见的回复邮件分隔符模式
    reply_patterns = [
        r'-----Original Message-----',
        r'----- 原始邮件 -----',
        r'From:.*?Sent:.*?To:.*?Subject:',
        r'发件人.*?发送时间.*?收件人.*?主题:',
        r'_{10,}',  # 多个下划线
        r'-{3,}\s*Original Message\s*-{3,}',
        r'On.*wrote:',
        r'在.*写道：',
    ]
    
    # 尝试找到回复分隔符
    for pattern in reply_patterns:
        if re.search(pattern, text, re.IGNORECASE | re.DOTALL):
            # 如果找到分隔符，提取分隔符之前的内容（新写的部分）
            parts = re.split(pattern, text, maxsplit=1, flags=re.IGNORECASE | re.DOTALL)
            if len(parts) > 1:
                # 同时保留新内容和原始邮件内容，但优先处理新内容
                new_content = parts[0].strip()
                original_content = parts[1].strip()
                
                # 尝试从原始邮件中提取主题
                subject_match = re.search(r'主题\s*(.+?)(?:\s+收件人|$)', original_content)
                if subject_match:
                    original_subject = subject_match.group(1).strip()
                    # 总是添加原始主题，因为可能包含重要信息
                    return f"{new_content}\n原始主题: {original_subject}"
                else:
                    # 尝试英文主题
                    subject_match = re.search(r'Subject:\s*(.+?)(?:\n|$)', original_content, re.IGNORECASE)
                    if subject_match:
                        original_subject = subject_match.group(1).strip()
                        return f"{new_content}\n原始主题: {original_subject}"
                
                return new_content
    
    # 如果没有找到分隔符，尝试直接提取主题信息
    subject_patterns = [
        r'(?:主题|Subject):\s*(.+?)(?:\n|$)',
        r'Re:\s*(.+?)(?:\n|$)',
        r'回复\s*[:：]\s*(.+?)(?:\n|$)',
    ]
    
    for pattern in subject_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            subject = match.group(1).strip()
            # 如果主题包含学生信息，添加到正文中
            if re.search(r'\d{6,}', subject) or re.search(r'[\u4e00-\u9fa5]{2,4}', subject):
                return f"{text}\n原始主题: {subject}"
    
    return text

def clean_text(text: str) -> str:
    """
    清理文本，去除多余空白和特殊字符
    """
    if not text:
        return ""
    
    # HTML解码
    text = html.unescape(text)
    
    # 去除多余的空白字符
    text = re.sub(r'\s+', ' ', text)
# 去除特殊字符但保留中文、英文、数字和常用标点
    text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\s\+\-_.,，。、；：""\'（）【】\[\]{}]', '', text)
    return text.strip()

def extract_student_info_from_text(text: str) -> Dict[str, any]:
    """
    从文本中提取学生信息
    
    Args:
        text: 要分析的文本内容
        
    Returns:
        包含提取信息和置信度的字典
    """
    if not text:
        return {
            "student_id": "",
            "name": "",
            "assignment": "",
            "confidence": 0,
            "matches": []
        }
    
    result = {
        "student_id": "",
        "name": "",
        "assignment": "",
        "confidence": 0,
        "matches": []
    }
    
    # 0. 处理原始主题信息
    original_subject_match = re.search(r'原始主题[:：]\s*(.+)', text)
    if original_subject_match:
        original_subject = original_subject_match.group(1).strip()
        # 将原始主题内容添加到文本中进行解析
        text = f"{text}\n{original_subject}"
        
    
    # 预处理文本：去除多余空白符和日期干扰
    text = re.sub(r'\s+', ' ', text)  # 多个空白符合并为一个空格
    text = text.strip()  # 去除首尾空白
    
    # 移除明显的日期模式，避免干扰学号识别（更精确的匹配）
    # 暂时注释掉日期移除，改用更直接的方法
    # text = re.sub(r'(?<!\d)20\d{6}[.\-_]*\d{1,2}[.\-_]*\d{1,2}(?!\d)', '', text)  # 移除独立的日期，如 20251215 或 2025.12.15
    # text = re.sub(r'(?<!\d)\d{4}[.\-_]*\d{1,2}[.\-_]*\d{1,2}(?!\d)', '', text)  # 移除其他独立日期格式
    
    # 1. 提取学号 - 使用更简单直接的方法
    # 首先检测并排除邮箱地址中的数字（只在明确的邮箱上下文中排除）
    email_patterns = [
        r'(2025\d{9})[a-zA-Z0-9._%+-]*@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',  # 2025开头的13位学号
        r'(\d{8})[a-zA-Z0-9._%+-]*@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',           # 8位学号
        r'(2025\d{9})qq\.com',
        r'(\d{8})qq\.com',
        r'(2025\d{9})\s*@\s*[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',
        r'(\d{8})\s*@\s*[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',
    ]
    
    # 收集所有可能是邮箱前缀的数字，这些不应该被当作学号
    excluded_ids = set()
    for pattern in email_patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for match in matches:
            excluded_ids.add(match.group(1))
    
    # 查找所有可能的学号候选（按出现顺序）
    candidates = []
    
    # 查找所有数字序列，然后智能匹配
    all_numbers = []
    for match in re.finditer(r'\d+', text):
        all_numbers.append({
            'start': match.start(),
            'number': match.group(),
            'end': match.end()
        })
    
    # 查找学号（支持12位和13位）
    for i, num_info in enumerate(all_numbers):
        number = num_info['number']
        
        # 直接检查是否是12位或13位学号
        if number.startswith('2025') and len(number) >= 4:
            # 尝试组合后续的数字序列，但更精确地控制长度
            combined = number
            for j in range(i + 1, min(i + 3, len(all_numbers))):  # 最多组合2个后续序列
                next_num = all_numbers[j]['number']
                # 只有当组合后长度不超过13位时才继续
                if len(combined) + len(next_num) <= 13:
                    combined += next_num
                else:
                    break
            
            # 清理组合后的数字
            clean_id = re.sub(r'[^0-9]', '', combined)
            
            # 检查是否是有效的12位或13位学号
            if len(clean_id) in [12, 13] and clean_id.startswith('2025'):
                candidates.append({
                    'start': num_info['start'],
                    'clean_id': clean_id,
                    'raw_id': combined,
                    'type': f'{len(clean_id)}位'
                })
                break  # 找到学号就停止
    
    # 查找标准13位学号（2025开头）
    for num_info in all_numbers:
        number = num_info['number']
        if len(number) == 13 and number.startswith('2025'):
            if number not in [c['clean_id'] for c in candidates]:
                candidates.append({
                    'start': num_info['start'],
                    'clean_id': number,
                    'raw_id': number,
                    'type': '13位'
                })
    
    # 查找8位学号
    for num_info in all_numbers:
        number = num_info['number']
        if len(number) == 8:
            # 确保不是13位学号的一部分
            if not any(number in c['clean_id'] for c in candidates):
                candidates.append({
                    'start': num_info['start'],
                    'clean_id': number,
                    'raw_id': number,
                    'type': '8位'
                })
    
    # 按出现位置排序，优先选择最早出现的
    candidates.sort(key=lambda x: x['start'])
    
    # 优先选择13位学号
    selected_id = None
    for candidate in candidates:
        if candidate['clean_id'] not in excluded_ids:  # 确保不是邮箱前缀
            selected_id = candidate
            break
    
    if selected_id:
        result["student_id"] = selected_id['clean_id']
        result["matches"].append(f"学号匹配: {selected_id['clean_id']} (原始: {selected_id['raw_id']}, 类型: {selected_id['type']})")
    
    # 2. 提取姓名
    name_patterns = [
        r'姓名[：:\s]*([\u4e00-\u9fa5]{2,4})',
        r'姓\s*名[：:\s]*([\u4e00-\u9fa5]{2,4})',
        r'学生[：:\s]*([\u4e00-\u9fa5]{2,4})',
        r'我是([\u4e00-\u9fa5]{2,4})',
        r'提交人[：:\s]*([\u4e00-\u9fa5]{2,4})',
        # 英文姓名模式
        r'name[：:\s]*([a-zA-Z]{2,20})',
        r'Name[：:\s]*([a-zA-Z]{2,20})',
        # 通用模式：学号-姓名（更严格的匹配）
        r'^(\d{6,12})[-\s]([\u4e00-\u9fa5]{2,4})(?=\s|$)',  # 学号-姓名格式，必须在开头
        r'^([\u4e00-\u9fa5]{2,4})(?=\s|$|[^\u4e00-\u9fa5])'  # 姓名后跟非中文字符或结束
    ]
    
    # 常见非姓名词汇，需要排除
    non_name_words = {
        '认真生活', '端正态度', '好好学习', '天天向上', '努力学习', 
        '认真学习', '态度端正', '生活态度', '学习态度',
        '作业完成', '提交作业', '课程作业', '实验报告'
    }
    
    for pattern in name_patterns:
        matches = re.finditer(pattern, text, re.IGNORECASE)
        for match in matches:
            if match.groups():
                name = match.group(1)
            else:
                name = match.group(0)
            if len(name) >= 2 and name not in non_name_words:  # 确保是有效的姓名长度且不是非姓名词汇
                result["name"] = name
                result["matches"].append(f"姓名匹配: {name} (模式: {pattern})")
                break
        if result["name"]:
            break
    
    # 3. 提取作业信息 - 使用更智能的方法
    # 如果已经找到姓名和学号，提取剩余部分作为作业名
    if result["name"] and result["student_id"]:
        # 使用最简单直接的方法：基于原始文本逐步清理
        remaining_text = text
        
        # 移除姓名
        if result["name"]:
            remaining_text = remaining_text.replace(result["name"], "")
        
        # 移除学号（使用所有可能的格式）
        if result["student_id"]:
            # 移除清理后的学号
            remaining_text = remaining_text.replace(result["student_id"], "")
            
            # 对于13位学号，移除各种可能的格式
            if len(result["student_id"]) == 13 and result["student_id"].startswith('2025'):
                suffix = result["student_id"][4:]  # 2025后的9位数字
                
                # 移除 2025-XXXXXXXXX 格式
                remaining_text = remaining_text.replace(f"2025-{suffix}", "")
                
                # 移除 2025-XXXX-XXXXX 格式（分段格式）
                if len(suffix) >= 5:
                    remaining_text = remaining_text.replace(f"2025-{suffix[:4]}-{suffix[4:]}", "")
                
                # 移除 2025XXXXXXXXX 格式
                remaining_text = remaining_text.replace(f"2025{suffix}", "")
                
                # 移除各种可能的分段组合
                for i in range(1, len(suffix)):
                    part1 = suffix[:i]
                    part2 = suffix[i:]
                    remaining_text = remaining_text.replace(f"2025-{part1}-{part2}", "")
                    remaining_text = remaining_text.replace(f"2025{part1}-{part2}", "")
                    remaining_text = remaining_text.replace(f"2025-{part1}{part2}", "")
            
            # 对于8位学号，移除各种格式
            if len(result["student_id"]) == 8:
                remaining_text = remaining_text.replace(result["student_id"], "")
        
        # 清理剩余文本：移除分隔符和空白
        remaining_text = re.sub(r'[-_\s]+', '', remaining_text)
        remaining_text = remaining_text.strip()
        
        # 如果清理后还有内容，作为作业名
        if remaining_text and len(remaining_text) >= 1:
            result["assignment"] = remaining_text
            result["matches"].append(f"作业匹配: {remaining_text} (剩余文本提取)")
        else:
            # 如果没有剩余内容，使用传统模式匹配
            assignment_patterns = [
                r'作业[：:\s]*([^\n\r，。；;]{1,20})',
                r'第[一二三四五六七八九十\d]+次作业',
                r'作业[一二三四五六七八九十\d]+',
                r'实验[一二三四五六七八九十\d]+',
                r'project\s*\d*',
                r'lab\s*\d*',
                r'assignment\s*\d*',
                r'hw\s*\d*',
                r'项目[：:\s]*([^\n\r，。；;]{1,20})',
                r'实验[：:\s]*([^\n\r，。；;]{1,20})',
                r'标题[：:\s]*([^\n\r，。；;]{1,20})',
                r'提交[：:\s]*([^\n\r，。；;]{1,20})',
                r'最终报告[：:\s]*([^\n\r，。；;]{1,20})',
                r'报告[：:\s]*([^\n\r，。；;]{1,20})'
            ]
            
            for pattern in assignment_patterns:
                matches = re.finditer(pattern, text, re.IGNORECASE)
                for match in matches:
                    assignment = match.group(1) if match.groups() else match.group(0)
                    assignment = assignment.strip()
                    if len(assignment) >= 1:
                        result["assignment"] = assignment
                        result["matches"].append(f"作业匹配: {assignment} (模式: {pattern})")
                        break
                if result["assignment"]:
                    break
    else:
        # 如果没有找到姓名或学号，使用传统模式匹配
        assignment_patterns = [
            r'作业[：:\s]*([^\n\r，。；;]{1,20})',
            r'第[一二三四五六七八九十\d]+次作业',
            r'作业[一二三四五六七八九十\d]+',
            r'实验[一二三四五六七八九十\d]+',
            r'project\s*\d*',
            r'lab\s*\d*',
            r'assignment\s*\d*',
            r'hw\s*\d*',
            r'项目[：:\s]*([^\n\r，。；;]{1,20})',
            r'实验[：:\s]*([^\n\r，。；;]{1,20})',
            r'标题[：:\s]*([^\n\r，。；;]{1,20})',
            r'提交[：:\s]*([^\n\r，。；;]{1,20})',
            r'最终报告[：:\s]*([^\n\r，。；;]{1,20})',
            r'报告[：:\s]*([^\n\r，。；;]{1,20})'
        ]
        
        for pattern in assignment_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                assignment = match.group(1) if match.groups() else match.group(0)
                assignment = assignment.strip()
                if len(assignment) >= 1:
                    result["assignment"] = assignment
                    result["matches"].append(f"作业匹配: {assignment} (模式: {pattern})")
                    break
            if result["assignment"]:
                break
    
    # 4. 计算置信度
    confidence_score = 0
    if result["student_id"]:
        confidence_score += 40
    if result["name"]:
        confidence_score += 35
    if result["assignment"]:
        confidence_score += 25
    
    # 额外加分项
    if result["student_id"] and result["name"]:
        confidence_score += 10  # 学号和姓名都找到
    if len(result["matches"]) >= 3:
        confidence_score += 5   # 匹配模式多
    
    result["confidence"] = min(confidence_score, 100)
    
    return result

def extract_info_from_subject(subject: str) -> Dict[str, any]:
    """
    从邮件标题中提取学生信息
    """
    return extract_student_info_from_text(subject)

def extract_info_from_body(body: str) -> Dict[str, any]:
    """
    从邮件正文中提取学生信息
    """
    # 先处理回复邮件格式
    processed_body = extract_reply_info(body)
    return extract_student_info_from_text(processed_body)

def extract_info_from_filename(filename: str) -> Dict[str, any]:
    """
    从文件名中提取学生信息
    """
    return extract_student_info_from_text(filename)

def extract_info_from_sender(sender: str) -> Dict[str, any]:
    """
    从发件人信息中提取学生信息（置信度较低）
    """
    result = {
        "student_id": "",
        "name": "",
        "assignment": "",
        "confidence": 0,
        "matches": []
    }
    
    if not sender:
        return result
    
    # 尝试从邮箱地址中提取姓名
    email_pattern = r'(.+?)@'
    match = re.search(email_pattern, sender)
    if match:
        potential_name = match.group(1)
        # 清理邮箱前缀
        potential_name = re.sub(r'[._-]', ' ', potential_name).strip()
        
        # 检查是否包含中文
        if re.search(r'[\u4e00-\u9fa5]', potential_name):
            result["name"] = potential_name
            result["confidence"] = 20
            result["matches"].append(f"发件人姓名: {potential_name}")
    
    return result

def combine_extraction_results(subject_result: Dict, body_result: Dict, 
                         filename_result: Dict, sender_result: Dict) -> Dict:
    """
    合并多个来源的提取结果，选择最佳匹配
    """
    combined = {
        "student_id": "",
        "name": "",
        "assignment": "",
        "confidence": 0,
        "source": "",
        "all_matches": []
    }
    
    # 按优先级排序的结果
    candidates = [
        (subject_result, "标题"),
        (body_result, "正文"),
        (filename_result, "文件名"),
        (sender_result, "发件人")
    ]
    
    # 选择最佳学号
    best_id = ""
    best_id_confidence = 0
    best_id_source = ""
    
    for result, source in candidates:
        if result.get("student_id") and result.get("confidence", 0) > best_id_confidence:
            best_id = result["student_id"]
            best_id_confidence = result.get("confidence", 0)
            best_id_source = source
    
    combined["student_id"] = best_id
    
    # 选择最佳姓名
    best_name = ""
    best_name_confidence = 0
    best_name_source = ""
    
    for result, source in candidates:
        if result.get("name") and result.get("confidence", 0) > best_name_confidence:
            best_name = result["name"]
            best_name_confidence = result.get("confidence", 0)
            best_name_source = source
    
    combined["name"] = best_name
    
    # 选择最佳作业信息
    best_assignment = ""
    best_assignment_confidence = 0
    best_assignment_source = ""
    
    for result, source in candidates:
        if result.get("assignment") and result.get("confidence", 0) > best_assignment_confidence:
            best_assignment = result["assignment"]
            best_assignment_confidence = result.get("confidence", 0)
            best_assignment_source = source
    
    combined["assignment"] = best_assignment
    
    # 计算综合置信度
    total_confidence = best_id_confidence + best_name_confidence + best_assignment_confidence
    combined["confidence"] = min(total_confidence / 3, 100) if total_confidence > 0 else 0
    
    # 记录信息来源
    sources = []
    if best_id_source:
        sources.append(f"学号来自{best_id_source}")
    if best_name_source:
        sources.append(f"姓名来自{best_name_source}")
    if best_assignment_source:
        sources.append(f"作业来自{best_assignment_source}")
    
    combined["source"] = ", ".join(sources)
    
    # 收集所有匹配信息
    for result, source in candidates:
        for match in result.get("matches", []):
            combined["all_matches"].append(f"{source}: {match}")
    
    return combined