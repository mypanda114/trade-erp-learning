# -*- coding: utf-8 -*-
"""
库存管理模块 - 生成4个核心表单，直接添加数字前缀
包含：库存台账表、出入库流水表、库存盘点单、库存调拨单
输出路径：D:\Trade-ERP-Learning\output\7.库存管理\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

# 固定随机种子
random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\7.库存管理"
os.makedirs(output_dir, exist_ok=True)

# ================== 辅助函数 ==================
def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))

def find_base_file(base_dir, target_name):
    """在基础数据文件夹中查找目标文件（忽略数字前缀）"""
    for f in os.listdir(base_dir):
        if f.endswith('.xlsx'):
            name_without_prefix = re.sub(r'^\d+\.\s*', '', f)
            if name_without_prefix == target_name:
                return os.path.join(base_dir, f)
    fallback = os.path.join(base_dir, target_name)
    if os.path.exists(fallback):
        return fallback
    raise FileNotFoundError(f"未找到文件 {target_name} 在 {base_dir}")

def save_with_prefix(df, filename, prefix_num):
    final_filename = f"{prefix_num}.{filename}"
    filepath = os.path.join(output_dir, final_filename)
    df.to_excel(filepath, index=False)
    print(f"已生成：{final_filename}")

# ================== 该模块文件顺序 ==================
FILE_ORDER = [
    "库存台账表",
    "出入库流水表",
    "库存盘点单",
    "库存调拨单"
]

# ================== 读取基础数据 ==================
base_dir = r"D:\Trade-ERP-Learning\output\1.基础数据"

# 产品
products_file = find_base_file(base_dir, "产品主数据表.xlsx")
products_df = pd.read_excel(products_file)
product_codes = products_df["产品编码"].tolist()
print(f"读取到 {len(product_codes)} 个产品")

# 员工表
employee_file = find_base_file(base_dir, "员工表.xlsx")
employees_df = pd.read_excel(employee_file)
print(f"读取到 {len(employees_df)} 名员工")

employee_info = {}
for _, row in employees_df.iterrows():
    employee_info[row["工号"]] = {
        "部门": row["所属部门"],
        "职位": row["职位"],
        "状态": row["状态"]
    }

# 仓储物流部在职员工
warehouse_staff = [eid for eid, info in employee_info.items()
                   if info["部门"] == "仓储物流部" and info["状态"] == "在职"]
print(f"仓储物流部在职员工总数: {len(warehouse_staff)}")
if not warehouse_staff:
    print("警告: 仓储物流部无在职员工！")

# ================== 1. 库存台账表（期初库存） ==================
inventory_ledger_data = [
    ("P001", "WH02", "", 5, 2, "台", "2024-06-01"),
    ("P002", "WH03", "", 30, 15, "个", "2024-06-01"),
    ("P003", "WH03", "", 50, 30, "个", "2024-06-01"),
    ("P004", "WH01", "B20240501", 800, 300, "个", "2024-06-01"),
    ("P005", "WH01", "BATCH20240520", 15, 5, "吨", "2024-06-01"),
    ("P006", "WH01", "", 1500, 800, "公斤", "2024-06-01"),
    ("P007", "WH02", "", 10, 5, "台", "2024-06-01"),
    ("P008", "WH03", "", 20, 10, "台", "2024-06-01"),
    ("P009", "WH01", "", 5000, 2000, "米", "2024-06-01"),
    ("P010", "WH03", "", 60, 30, "个", "2024-06-01"),
]
inventory_ledger = pd.DataFrame(inventory_ledger_data,
                                columns=["物料编码", "仓库", "批次号", "当前库存", "安全库存", "单位", "最后更新时间"])
save_with_prefix(inventory_ledger, FILE_ORDER[0] + ".xlsx", 1)

# ================== 2. 出入库流水表 ==================
transactions = []
trans_idx = 1

# 采购入库（模拟10条）
for i in range(1, 11):
    if not warehouse_staff:
        break
    trans_no = f"TR-{trans_idx:04d}"
    product = random.choice(product_codes)
    qty = random.randint(10, 200)
    trans_date = random_date(datetime(2024,1,1), datetime(2024,6,30)).date()
    operator = random.choice(warehouse_staff)
    operator_dept = employee_info[operator]["部门"]
    transactions.append((trans_no, "采购入库", f"PREC-{i:04d}", product, qty, 0,
                         trans_date.strftime("%Y-%m-%d %H:%M:%S"), operator, operator_dept))
    trans_idx += 1

# 生产领料（模拟10条）
for i in range(1, 11):
    if not warehouse_staff:
        break
    trans_no = f"TR-{trans_idx:04d}"
    product = random.choice(product_codes)
    qty = -random.randint(1, 50)
    trans_date = random_date(datetime(2024,1,1), datetime(2024,6,30)).date()
    operator = random.choice(warehouse_staff)
    operator_dept = employee_info[operator]["部门"]
    transactions.append((trans_no, "生产领料", f"PICK-{i:04d}", product, qty, 0,
                         trans_date.strftime("%Y-%m-%d %H:%M:%S"), operator, operator_dept))
    trans_idx += 1

# 生产入库（模拟10条）
for i in range(1, 11):
    if not warehouse_staff:
        break
    trans_no = f"TR-{trans_idx:04d}"
    product = random.choice(product_codes)
    qty = random.randint(1, 20)
    trans_date = random_date(datetime(2024,1,1), datetime(2024,6,30)).date()
    operator = random.choice(warehouse_staff)
    operator_dept = employee_info[operator]["部门"]
    transactions.append((trans_no, "生产入库", f"PENT-{i:04d}", product, qty, 0,
                         trans_date.strftime("%Y-%m-%d %H:%M:%S"), operator, operator_dept))
    trans_idx += 1

# 销售出库（模拟10条）
for i in range(1, 11):
    if not warehouse_staff:
        break
    trans_no = f"TR-{trans_idx:04d}"
    product = random.choice(product_codes)
    qty = -random.randint(1, 10)
    trans_date = random_date(datetime(2024,1,1), datetime(2024,6,30)).date()
    operator = random.choice(warehouse_staff)
    operator_dept = employee_info[operator]["部门"]
    transactions.append((trans_no, "销售出库", f"SHP-{i:04d}", product, qty, 0,
                         trans_date.strftime("%Y-%m-%d %H:%M:%S"), operator, operator_dept))
    trans_idx += 1

transactions_df = pd.DataFrame(transactions,
                               columns=["流水号", "单据类型", "关联单号", "物料编码", "数量", "操作后库存",
                                        "操作时间", "操作人", "操作人部门"])
save_with_prefix(transactions_df, FILE_ORDER[1] + ".xlsx", 2)

# ================== 3. 库存盘点单 ==================
count_data = [
    ("CNT-001", "WH01", "P005", 15, 14.8, -0.2, "2024-06-30", "已完成"),
    ("CNT-002", "WH02", "P001", 5, 5, 0, "2024-06-30", "已完成"),
    ("CNT-003", "WH03", "P002", 30, 29, -1, "2024-06-30", "审核中"),
    ("CNT-004", "WH01", "P004", 800, 798, -2, "2024-06-30", "已完成"),
    ("CNT-005", "WH03", "P008", 20, 20, 0, "2024-06-30", "已完成"),
]
count = pd.DataFrame(count_data,
                     columns=["盘点单号", "仓库", "物料编码", "账面数量", "实盘数量", "盈亏数量", "盘点日期", "状态"])
save_with_prefix(count, FILE_ORDER[2] + ".xlsx", 3)

# ================== 4. 库存调拨单 ==================
transfer_data = [
    ("TF-001", "WH01", "WH03", "P004", 100, "2024-06-15", "已完成"),
    ("TF-002", "WH01", "WH03", "P009", 200, "2024-06-20", "已完成"),
    ("TF-003", "WH02", "WH03", "P008", 5, "2024-06-25", "待审核"),
]
transfer = pd.DataFrame(transfer_data,
                        columns=["调拨单号", "调出仓库", "调入仓库", "物料编码", "调拨数量", "调拨日期", "状态"])
save_with_prefix(transfer, FILE_ORDER[3] + ".xlsx", 4)

print("\n库存管理模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)