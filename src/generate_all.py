# -*- coding: utf-8 -*-
"""
主控脚本 - 依次执行所有模块脚本，生成全部42个Excel文件
最终输出目录：D:\Trade-ERP-Learning\output
"""

import os
import sys
import subprocess

# 输出根目录
OUTPUT_ROOT = r"D:\Trade-ERP-Learning\output"

# 模块文件夹名称及顺序（与设计方案一致）
MODULE_FOLDERS = [
    "1.基础数据",
    "2.技术管理",
    "3.销售管理",
    "4.采购管理",
    "5.生产管理",
    "6.质量管理",
    "7.库存管理",
    "8.外贸物流",
    "9.财务管理"
]

def main():
    print("="*50)
    print("开始生成全部表单数据...")
    print(f"输出根目录：{OUTPUT_ROOT}")
    print("="*50)

    # 1. 创建模块文件夹
    for folder in MODULE_FOLDERS:
        folder_path = os.path.join(OUTPUT_ROOT, folder)
        os.makedirs(folder_path, exist_ok=True)
        print(f"已创建/确认文件夹：{folder}")

    # 2. 按顺序调用模块脚本
    module_scripts = [
        "1_gen_base.py",
        "2_gen_plm.py",
        "3_gen_sales.py",
        "4_gen_purchase.py",
        "5_gen_production.py",
        "6_gen_quality.py",
        "7_gen_inventory.py",
        "8_gen_logistics.py",
        "9_gen_finance.py"
    ]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for script in module_scripts:
        script_path = os.path.join(script_dir, script)
        if not os.path.exists(script_path):
            print(f"警告：脚本 {script} 不存在，跳过")
            continue
        print(f"\n正在执行：{script}")
        try:
            result = subprocess.run([sys.executable, script_path], capture_output=True, text=True, check=True)
            print(result.stdout)
            if result.stderr:
                print("错误输出：", result.stderr)
        except subprocess.CalledProcessError as e:
            print(f"执行 {script} 时出错：{e}")
            print(e.stdout)
            print(e.stderr)
            sys.exit(1)

    print("\n" + "="*50)
    print("所有数据生成完毕！")
    print(f"输出目录：{OUTPUT_ROOT}")

if __name__ == "__main__":
    main()
