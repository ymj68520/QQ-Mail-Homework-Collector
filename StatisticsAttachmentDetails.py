import os
import re
import pandas as pd
from dotenv import load_dotenv
import sys
import io

# ===========================================
# 强制将标准输出设置为 utf-8，解决 emoji 报错和中文乱码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# ===========================================

# ================= 配置区域 =================
# 加载环境变量
load_dotenv()
SAVE_DIR = os.getenv('SAVE_DIR', 'downloaded_attachments') # 附件保存的根目录
OUTPUT_FILE = '作业统计表.xlsx'
# ===========================================

def parse_folder_name(folder_name):
    """
    核心逻辑：从混乱的文件夹名字中提取 [学号, 姓名, 作业信息]
    """
    info = {
        "original_text": folder_name,
        "student_id": "",
        "name": "",
        "assignment": ""
    }

    # 1. 预处理：把各种分隔符 (+, -, _, 甚至中文逗号) 全部替换成 空格
    # 这样 "学号+姓名" 和 "学号 姓名" 就变成一样了
    clean_text = re.sub(r'[+\-_,，\.。=]', ' ', folder_name)
    
    # 2. 提取学号 (通常是连续的数字，假设至少6位)
    # 逻辑：找到长数字，认为是学号，提取后从字符串中移除，避免干扰后续步骤
    id_match = re.search(r'\d{6,}', clean_text)
    if id_match:
        info["student_id"] = id_match.group()
        # 将找到的学号替换为空格，以免干扰后续分析
        clean_text = clean_text.replace(info["student_id"], ' ')
    
    # 3. 提取姓名 (假设是2-4个中文字符)
    # 逻辑：在剩下的文本中找 2-4 个连续的汉字
    name_match = re.search(r'[\u4e00-\u9fa5]{2,4}', clean_text)
    if name_match:
        info["name"] = name_match.group()
        clean_text = clean_text.replace(info["name"], ' ')
    
    # 4. 剩下的内容就是“作业名”了
    # 去除首尾空格，去除多余的空格
    remaining = re.sub(r'\s+', ' ', clean_text).strip()
    
    # 如果没找到中文名，尝试在剩下的部分里找英文名（纯字母），但这容易和作业名混淆
    # 这里做一个简单的兜底：如果还没找到名字，且剩下的部分包含短单词，可能就是名字
    if not info["name"] and not info["assignment"]:
         parts = remaining.split()
         # 简单的启发式：如果剩下两部分，且第一部分很短，可能是英文名
         if len(parts) >= 2 and len(parts[0]) < 10:
             # 这里只是猜测，英文名情况比较复杂，视情况启用
             pass

    info["assignment"] = remaining
    return info

def format_size(size_bytes):
    """将字节转换为易读的格式 (KB, MB, GB)"""
    if size_bytes == 0:
        return "0B"
    size_name = ("B", "KB", "MB", "GB")
    i = 0
    p = size_bytes
    while p >= 1024 and i < len(size_name) - 1:
        p /= 1024.0
        i += 1
    return f"{p:.2f} {size_name[i]}"

def generate_report():
    if not os.path.exists(SAVE_DIR):
        print(f"❌ 找不到目录: {SAVE_DIR}，请先运行下载程序。")
        return

    print(f"正在扫描目录: {SAVE_DIR} ...")
    
    # 使用字典按学号聚合数据
    student_map = {}
    # 存储无法识别学号的记录
    no_id_list = []
    
    # 遍历根目录下的所有文件夹
    for folder in os.listdir(SAVE_DIR):
        folder_path = os.path.join(SAVE_DIR, folder)
        
        # 确保是文件夹
        if os.path.isdir(folder_path):
            # 1. 解析文件夹名字
            parsed_info = parse_folder_name(folder)
            
            # 2. 统计里面的文件
            files = os.listdir(folder_path)
            file_count = len(files)
            
            # 计算文件夹内文件总大小
            folder_size = 0
            for f in files:
                fp = os.path.join(folder_path, f)
                if os.path.isfile(fp):
                    folder_size += os.path.getsize(fp)
            
            sid = parsed_info["student_id"]
            
            if sid:
                # 如果有学号，进行聚合
                if sid not in student_map:
                    student_map[sid] = {
                        "学号": sid,
                        "姓名": parsed_info["name"],
                        "作业备注": set(),  # 使用集合去重
                        "来源文件夹": [],
                        "附件数量": 0,
                        "附件总大小Bytes": 0,
                        "附件列表": []
                    }
                
                # 更新信息
                entry = student_map[sid]
                # 如果当前解析的名字比已有的长（更完整），则更新名字
                if len(parsed_info["name"]) > len(entry["姓名"]):
                    entry["姓名"] = parsed_info["name"]
                
                if parsed_info["assignment"]:
                    entry["作业备注"].add(parsed_info["assignment"])
                
                entry["来源文件夹"].append(parsed_info["original_text"])
                entry["附件数量"] += file_count
                entry["附件总大小Bytes"] += folder_size
                entry["附件列表"].extend(files)
                
            else:
                # 没有学号，作为单独条目
                no_id_list.append({
                    "学号": "",
                    "姓名": parsed_info["name"],
                    "作业备注": parsed_info["assignment"],
                    "来源文件夹": parsed_info["original_text"],
                    "附件数量": file_count,
                    "附件总大小": format_size(folder_size),
                    "附件列表": "; ".join(files),
                    "状态": "无法识别学号" if file_count > 0 else "无附件(无学号)"
                })

    # 将聚合后的 student_map 转换为列表
    final_data = []
    
    # 处理聚合数据
    for sid, data in student_map.items():
        # 格式化
        note_str = "; ".join(sorted(list(data["作业备注"])))
        folder_str = "; ".join(data["来源文件夹"])
        file_str = "; ".join(data["附件列表"])
        size_str = format_size(data["附件总大小Bytes"])
        status = "正常" if data["附件数量"] > 0 else "无附件"
        
        final_data.append({
            "学号": sid,
            "姓名": data["姓名"],
            "作业备注": note_str,
            "来源文件夹": folder_str,
            "附件数量": data["附件数量"],
            "附件总大小": size_str,
            "附件列表": file_str,
            "状态": status
        })
    
    # 合并无学号数据
    final_data.extend(no_id_list)

    if not final_data:
        print("没有找到任何记录。")
        return

    # 使用 Pandas 生成表格
    df = pd.DataFrame(final_data)
    
    # 简单的排序：按学号排序
    df = df.sort_values(by="学号", ascending=True)

    print(f"扫描完成，共 {len(df)} 条学生/记录。正在写入 Excel...")
    
    try:
        # 使用 ExcelWriter 启用筛选功能
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            # 开启自动筛选
            worksheet.auto_filter.ref = worksheet.dimensions

        print(f"✅ 统计完成！文件已保存为: {OUTPUT_FILE}")
        
        # 打印预览
        print("\n--- 预览前 5 行 ---")
        # 确保预览列存在
        preview_cols = ["学号", "姓名", "作业备注", "附件数量", "附件总大小", "状态"]
        # 防止某些列不存在（比如全是无学号情况可能导致结构差异，虽然上面逻辑保证了键一致，但为了安全）
        actual_cols = [c for c in preview_cols if c in df.columns]
        print(df[actual_cols].head().to_string(index=False))
        
    except Exception as e:
        print(f"❌ 保存 Excel 失败: {e}")
        print("请确保你没有在其他软件中打开该 Excel 文件。")

if __name__ == "__main__":
    generate_report()