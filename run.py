#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QQ邮箱自动下载工具 - 快速启动脚本
"""

import os
import sys
import subprocess

def main():
    """主菜单"""
    print("=" * 60)
    print("🚀 QQ邮箱自动下载工具 - 增强版 v2.0")
    print("=" * 60)
    print()
    print("请选择功能：")
    print("1. 📥 增强版下载器（推荐）")
    print("2. 📊 按作业分组分析")
    print("3. 👥 按学生分组分析")
    print("4. 🖥️  GUI界面（增强版）")
    print("5. 🧪 运行测试")
    print("6. 📋 查看项目结构")
    print("0. 🚪 退出")
    print()
    
    while True:
        try:
            choice = input("请输入选项 (0-6): ").strip()
            
            if choice == "0":
                print("👋 再见！")
                break
            elif choice == "1":
                print("📥 启动增强版下载器...")
                subprocess.run([sys.executable, "src/EnhancedDownloadQQAttachments.py"])
                break
            elif choice == "2":
                print("📊 启动按作业分组分析...")
                subprocess.run([sys.executable, "src/MultiSubmissionAnalyzer.py"])
                break
            elif choice == "3":
                print("👥 启动按学生分组分析...")
                subprocess.run([sys.executable, "src/MultiAssignmentAnalyzer.py"])
                break
            elif choice == "4":
                print("🖥️  启动GUI界面...")
                subprocess.run([sys.executable, "src/enhanced_app_gui.py"])
                break
            elif choice == "5":
                print("🧪 运行测试...")
                subprocess.run([sys.executable, "tests/test_real_data.py"])
                break
            elif choice == "6":
                print("📋 项目结构：")
                print("""
📁 目录结构：
├── src/          # 核心功能模块
├── tests/        # 测试文件
├── docs/         # 文档
├── reports/      # 生成的报告
├── utils/        # 工具和配置
├── 25XC/         # 数据目录
└── 25TA/         # 数据目录

📖 详细说明请查看：项目结构说明.md
                """)
                continue
            else:
                print("❌ 无效选项，请重新输入 (0-6)")
                continue
                
        except KeyboardInterrupt:
            print("\n👋 用户取消，再见！")
            break
        except Exception as e:
            print(f"❌ 错误: {e}")
            break

if __name__ == "__main__":
    main()