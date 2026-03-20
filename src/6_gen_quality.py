# -*- coding: utf-8 -*-
"""
质量管理模块 - 生成1个核心表单，直接添加数字前缀
包含：质量检验表
输出路径：D:\Trade-ERP-Learning\output\6.质量管理\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

# 固定随机种子
random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\6.质量管理"
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
FILE_ORDER = ["质量检验表"]

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

# 质检部在职员工
quality_inspectors = [eid for eid, info in employee_info.items()
                      if info["部门"] == "质检部" and info["状态"] == "在职"]
print(f"质检部在职员工总数: {len(quality_inspectors)}")
if not quality_inspectors:
    print("警告: 质检部无在职员工！")

# 获取关联单号（若前序模块未运行则用模拟单号）
purchase_receipts = []
pr_path = os.path.join(r"D:\Trade-ERP-Learning\output\4.采购管理", "4.采购入库单.xlsx")
if os.path.exists(pr_path):
    pr_df = pd.read_excel(pr_path)
    purchase_receipts = pr_df["入库单号"].tolist()
else:
    purchase_receipts = [f"PREC-{i:04d}" for i in range(1, 21)]

production_entries = []
pe_path = os.path.join(r"D:\Trade-ERP-Learning\output\5.生产管理", "5.生产入库单.xlsx")
if os.path.exists(pe_path):
    pe_df = pd.read_excel(pe_path)
    production_entries = pe_df["入库单号"].tolist()
else:
    production_entries = [f"PENT-{i:04d}" for i in range(1, 26)]

op_reports = []
op_path = os.path.join(r"D:\Trade-ERP-Learning\output\5.生产管理", "2.工序报工表.xlsx")
if os.path.exists(op_path):
    op_df = pd.read_excel(op_path)
    op_reports = op_df["报工单号"].tolist()
else:
    op_reports = [f"RPT-{i:04d}" for i in range(1, 50)]

# ================== 1. 质量检验表 ==================
qc_records = []
qc_sources = ["来料", "工序", "成品"]
qc_statuses = ["待复核", "已关闭"]

# 采购入库质检（来料）
for i, rec_no in enumerate(purchase_receipts[:10]):
    if random.random() < 0.7 and quality_inspectors:
        qc_no = f"QC-{i+100:03d}"
        source_type = "来料"
        product = random.choice(product_codes)
        insp_qty = random.randint(10, 100)
        pass_qty = insp_qty if random.random() < 0.8 else random.randint(1, insp_qty-1)
        fail_qty = insp_qty - pass_qty
        reason = "" if fail_qty == 0 else random.choice(["尺寸超差", "外观瑕疵", "材质不符"])
        disposition = "合格" if fail_qty == 0 else "退货"
        inspector = random.choice(quality_inspectors)
        inspector_dept = employee_info[inspector]["部门"]
        inspect_time = datetime.now() - timedelta(days=random.randint(1, 60))
        status = random.choice(qc_statuses)
        qc_records.append((qc_no, source_type, rec_no, product, insp_qty, pass_qty, fail_qty,
                           reason, disposition, inspector, inspector_dept,
                           inspect_time.strftime("%Y-%m-%d %H:%M:%S"), "", status))

# 工序报工质检
for i, rep_no in enumerate(op_reports[:10]):
    if random.random() < 0.5 and quality_inspectors:
        qc_no = f"QC-{i+200:03d}"
        source_type = "工序"
        product = random.choice(product_codes)
        insp_qty = random.randint(1, 20)
        pass_qty = insp_qty if random.random() < 0.9 else random.randint(1, insp_qty-1)
        fail_qty = insp_qty - pass_qty
        reason = "" if fail_qty == 0 else "加工不良"
        disposition = "返工" if fail_qty > 0 else "合格"
        inspector = random.choice(quality_inspectors)
        inspector_dept = employee_info[inspector]["部门"]
        inspect_time = datetime.now() - timedelta(days=random.randint(1, 30))
        status = random.choice(qc_statuses)
        qc_records.append((qc_no, source_type, rep_no, product, insp_qty, pass_qty, fail_qty,
                           reason, disposition, inspector, inspector_dept,
                           inspect_time.strftime("%Y-%m-%d %H:%M:%S"), "", status))

# 生产入库质检（成品）
for i, ent_no in enumerate(production_entries[:10]):
    if quality_inspectors:
        qc_no = f"QC-{i+300:03d}"
        source_type = "成品"
        product = random.choice(product_codes)
        insp_qty = random.randint(1, 10)
        pass_qty = insp_qty if random.random() < 0.95 else random.randint(1, insp_qty-1)
        fail_qty = insp_qty - pass_qty
        reason = "" if fail_qty == 0 else "外观瑕疵"
        disposition = "合格" if fail_qty == 0 else "返工"
        inspector = random.choice(quality_inspectors)
        inspector_dept = employee_info[inspector]["部门"]
        inspect_time = datetime.now() - timedelta(days=random.randint(1, 15))
        status = random.choice(qc_statuses)
        qc_records.append((qc_no, source_type, ent_no, product, insp_qty, pass_qty, fail_qty,
                           reason, disposition, inspector, inspector_dept,
                           inspect_time.strftime("%Y-%m-%d %H:%M:%S"), "", status))

# 补充一些随机记录确保总数
while len(qc_records) < 20 and quality_inspectors:
    qc_no = f"QC-{len(qc_records)+400:03d}"
    source_type = random.choice(qc_sources)
    product = random.choice(product_codes)
    insp_qty = random.randint(1, 50)
    pass_qty = random.randint(1, insp_qty)
    fail_qty = insp_qty - pass_qty
    reason = "" if fail_qty == 0 else random.choice(["尺寸超差", "外观瑕疵", "加工不良"])
    disposition = "合格" if fail_qty == 0 else random.choice(["退货", "返工"])
    inspector = random.choice(quality_inspectors)
    inspector_dept = employee_info[inspector]["部门"]
    inspect_time = datetime.now() - timedelta(days=random.randint(1, 30))
    status = random.choice(qc_statuses)
    qc_records.append((qc_no, source_type, f"关联单号{len(qc_records)}", product, insp_qty, pass_qty, fail_qty,
                       reason, disposition, inspector, inspector_dept,
                       inspect_time.strftime("%Y-%m-%d %H:%M:%S"), "", status))

qc_df = pd.DataFrame(qc_records,
                     columns=["检验单号", "来源类型", "关联单号", "检验产品", "检验数量", "合格数量", "不合格数量",
                              "不合格原因", "处理方式", "检验员", "检验员部门", "检验时间", "复核意见", "状态"])
save_with_prefix(qc_df, FILE_ORDER[0] + ".xlsx", 1)

print("\n质量管理模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)