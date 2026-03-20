# -*- coding: utf-8 -*-
"""
初始化脚本 - 生成主控和9个模块的空白脚本框架
运行后将在当前目录下生成 generate_all.py 和 1_gen_base.py ~ 9_gen_finance.py
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MAIN_TEMPLATE = '''# -*- coding: utf-8 -*-
"""
主控脚本 - 依次执行所有模块脚本，生成全部42个Excel文件
最终输出目录：D:\\Trade-ERP-Learning\\output
"""

import os
import sys
import subprocess

OUTPUT_ROOT = r"D:\\Trade-ERP-Learning\\output"

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

    for folder in MODULE_FOLDERS:
        folder_path = os.path.join(OUTPUT_ROOT, folder)
        os.makedirs(folder_path, exist_ok=True)
        print(f"已创建/确认文件夹：{folder}")

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
        print(f"\\n正在执行：{script}")
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

    print("\\n" + "="*50)
    print("所有数据生成完毕！")
    print(f"输出目录：{OUTPUT_ROOT}")

if __name__ == "__main__":
    main()
'''

MODULE_TEMPLATE = '''# -*- coding: utf-8 -*-
"""
{module_name}模块 - 生成表单
输出路径：D:\\Trade-ERP-Learning\\output\\{folder_name}
"""

import pandas as pd
import os
import random
from datetime import datetime

def generate():
    # TODO: 实现数据生成逻辑
    pass

if __name__ == "__main__":
    generate()
'''

MODULES = {
    "1_gen_base.py": ("基础数据", "1.基础数据"),
    "2_gen_plm.py": ("技术管理", "2.技术管理"),
    "3_gen_sales.py": ("销售管理", "3.销售管理"),
    "4_gen_purchase.py": ("采购管理", "4.采购管理"),
    "5_gen_production.py": ("生产管理", "5.生产管理"),
    "6_gen_quality.py": ("质量管理", "6.质量管理"),
    "7_gen_inventory.py": ("库存管理", "7.库存管理"),
    "8_gen_logistics.py": ("外贸物流", "8.外贸物流"),
    "9_gen_finance.py": ("财务管理", "9.财务管理"),
}

def create_scripts():
    main_path = os.path.join(BASE_DIR, "generate_all.py")
    if not os.path.exists(main_path):
        with open(main_path, "w", encoding="utf-8") as f:
            f.write(MAIN_TEMPLATE)
        print(f"已创建：generate_all.py")
    else:
        print(f"文件已存在，跳过：generate_all.py")

    for filename, (module_name, folder_name) in MODULES.items():
        filepath = os.path.join(BASE_DIR, filename)
        if os.path.exists(filepath):
            print(f"文件已存在，跳过：{filename}")
            continue
        content = MODULE_TEMPLATE.format(module_name=module_name, folder_name=folder_name)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"已创建：{filename}")

    print("\n初始化完成！请根据设计方案填充各模块脚本的具体数据生成逻辑。")

if __name__ == "__main__":
    create_scripts()