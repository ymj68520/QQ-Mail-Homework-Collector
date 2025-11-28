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

def generate_report():
    if not os.path.exists(SAVE_DIR):
        print(f"❌ 找不到目录: {SAVE_DIR}，请先运行下载程序。")
        return

    print(f"正在扫描目录: {SAVE_DIR} ...")
    
    data_list = []
    
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
            file_names = "; ".join(files) # 把所有文件名拼在一起
            
            # 3. 汇总数据
            entry = {
                "文件夹原名": parsed_info["original_text"],
                "学号": parsed_info["student_id"],
                "姓名": parsed_info["name"],
                "作业备注/其他信息": parsed_info["assignment"],
                "附件数量": file_count,
                "附件列表": file_names,
                "状态": "正常" if file_count > 0 else "无附件"
            }
            data_list.append(entry)

    if not data_list:
        print("没有找到任何记录。")
        return

    # 使用 Pandas 生成表格
    df = pd.DataFrame(data_list)
    
    # 简单的排序：按学号排序（如果学号为空，放到最后）
    df = df.sort_values(by="学号", ascending=True)

    print(f"扫描完成，共 {len(df)} 条记录。正在写入 Excel...")
    
    try:
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ 统计完成！文件已保存为: {OUTPUT_FILE}")
        
        # 打印预览
        print("\n--- 预览前 5 行 ---")
        print(df[["学号", "姓名", "附件数量", "状态"]].head().to_string(index=False))
        
    except Exception as e:
        print(f"❌ 保存 Excel 失败: {e}")
        print("请确保你没有在其他软件中打开该 Excel 文件。")

if __name__ == "__main__":
    generate_report()