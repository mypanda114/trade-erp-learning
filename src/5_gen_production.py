# -*- coding: utf-8 -*-
"""
生产管理模块 - 生成5个核心表单，直接添加数字前缀
包含：生产工单表、工序报工表、生产领料单、生产退料单、生产入库单
输出路径：D:\Trade-ERP-Learning\output\5.生产管理\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

# 固定随机种子
random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\5.生产管理"
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
    "生产工单表",
    "工序报工表",
    "生产领料单",
    "生产退料单",
    "生产入库单"
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

# 构建员工信息字典：工号 -> {部门名称, 职位, 状态, 入职日期}
employee_info = {}
for _, row in employees_df.iterrows():
    employee_info[row["工号"]] = {
        "部门": row["所属部门"],
        "职位": row["职位"],
        "状态": row["状态"],
        "入职日期": row["入职日期"]
    }

# 生产部在职员工（操作工、班组长）
production_operators = [eid for eid, info in employee_info.items()
                        if info["部门"] == "生产部"
                        and info["状态"] == "在职"
                        and info["职位"] in ["操作工", "班组长"]]
print(f"生产部操作工/班组长在职人数: {len(production_operators)}")
if not production_operators:
    print("警告: 生产部无符合条件的在职操作工/班组长！")

# 仓储物流部在职员工（用于领料、入库操作人）
warehouse_staff = [eid for eid, info in employee_info.items()
                   if info["部门"] == "仓储物流部" and info["状态"] == "在职"]
print(f"仓储物流部在职员工总数: {len(warehouse_staff)}")
if not warehouse_staff:
    print("警告: 仓储物流部无在职员工！")

# 设备表
equipment_file = find_base_file(base_dir, "设备管理表.xlsx")
equipment_df = pd.read_excel(equipment_file)
equipment_codes = equipment_df["设备编码"].tolist()
print(f"读取到 {len(equipment_codes)} 台设备")

# 销售订单（用于关联）
sales_order_path = os.path.join(r"D:\Trade-ERP-Learning\output\3.销售管理", "2.销售订单表.xlsx")
if os.path.exists(sales_order_path):
    sales_orders_df = pd.read_excel(sales_order_path)
    sales_orders = sales_orders_df["订单号"].tolist()
    print(f"读取到 {len(sales_orders)} 个销售订单")
else:
    sales_orders = []
    print("销售订单表不存在，将使用模拟订单号")

# BOM简化（此处硬编码）
bom_dict = {
    "P001": [("P002", 2), ("P003", 1), ("P004", 4)],
    "P002": [("P005", 0.05), ("P006", 2)],
    "P007": [("P004", 2), ("P008", 1)],
    "P008": [("P009", 1.5), ("P010", 1)],
}
# 工艺路线简化
routing_dict = {
    "P001": [("机加工", 120), ("装配", 60)],
    "P002": [("绕线", 30), ("组装", 45)],
    "P007": [("箱体加工", 90), ("装配", 50)],
    "P008": [("插件", 20), ("测试", 15)],
}

# ================== 1. 生产工单表 ==================
work_orders = []
wo_statuses = ["计划", "执行中", "已完成", "取消"]
for i in range(1, 26):
    wo_no = f"WO-2024{str(i).zfill(4)}"
    source_so = random.choice(sales_orders) if sales_orders and random.random() < 0.6 else ""
    product = random.choice(["P001", "P002", "P007", "P008"])
    planned_qty = random.randint(1, 20)
    planned_start = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
    planned_end = planned_start + timedelta(days=random.randint(5, 30))
    actual_start = planned_start if random.random() < 0.7 else planned_start + timedelta(days=random.randint(1,5))
    workshop = random.choice(["金工车间", "装配车间", "电机车间", "电子车间"])
    equipment = random.choice(equipment_codes) if random.random() < 0.6 else ""
    status = random.choices(wo_statuses, weights=[20,50,25,5])[0]

    if status == "已完成":
        completed_qty = planned_qty
    elif status == "执行中":
        if planned_qty > 1:
            completed_qty = random.randint(1, planned_qty-1)
        else:
            completed_qty = 0
    else:
        completed_qty = 0

    good_qty = int(completed_qty * random.uniform(0.9, 1.0))
    defect_qty = completed_qty - good_qty

    work_orders.append((wo_no, source_so, product, planned_qty, completed_qty, good_qty, defect_qty,
                        planned_start.strftime("%Y-%m-%d"), planned_end.strftime("%Y-%m-%d"),
                        actual_start.strftime("%Y-%m-%d") if actual_start else "",
                        (actual_start+timedelta(days=random.randint(5,20))).strftime("%Y-%m-%d") if status=="已完成" else "",
                        workshop, equipment, status))

wo_df = pd.DataFrame(work_orders,
                     columns=["工单号", "来源销售订单", "生产产品", "计划数量", "已完工数量", "良品数量", "不良品数量",
                              "计划开工日", "计划完工日", "实际开工日", "实际完工日", "生产车间", "关联设备", "工单状态"])
save_with_prefix(wo_df, FILE_ORDER[0] + ".xlsx", 1)

# ================== 2. 工序报工表 ==================
op_reports = []
for wo in work_orders:
    wo_no, source_so, product, planned_qty, completed_qty, good_qty, defect_qty, p_start, p_end, a_start, a_end, workshop, equipment, status = wo
    if status in ["执行中", "已完成"] and completed_qty > 0 and product in routing_dict:
        if not production_operators:
            break
        processes = routing_dict[product]
        for seq, (process_name, std_min) in enumerate(processes, start=10):
            report_no = f"RPT-{wo_no[-4:]}{seq}"
            team = random.choice(["金工一班", "装配一班", "电机班", "电子班"])
            operator = random.choice(production_operators)
            operator_dept = employee_info[operator]["部门"]
            reported_qty = completed_qty
            good = reported_qty - random.randint(0, min(2, reported_qty))
            defect = reported_qty - good
            hours = round(reported_qty * std_min / 60, 1)
            report_time = datetime.strptime(a_start, "%Y-%m-%d") if a_start else datetime.strptime(p_start, "%Y-%m-%d")
            report_time = report_time + timedelta(days=random.randint(1,10), hours=random.randint(8,17))
            eq = equipment if random.random()<0.5 else ""
            op_reports.append((report_no, wo_no, process_name, team, operator, operator_dept,
                               reported_qty, good, defect, hours,
                               report_time.strftime("%Y-%m-%d %H:%M:%S"), eq))

op_reports_df = pd.DataFrame(op_reports,
                             columns=["报工单号", "关联工单", "工序名称", "操作班组", "操作人工号", "操作人部门",
                                      "报工数量", "良品数量", "不良品数量", "工时(小时)", "报工时间", "设备编号"])
save_with_prefix(op_reports_df, FILE_ORDER[1] + ".xlsx", 2)

# ================== 3. 生产领料单 ==================
picks = []
for wo in work_orders:
    wo_no, source_so, product, planned_qty, completed_qty, good_qty, defect_qty, p_start, p_end, a_start, a_end, workshop, equipment, status = wo
    if random.random() < 0.8 and product in bom_dict:
        for child, usage in bom_dict[product]:
            if not warehouse_staff:
                break
            pick_no = f"PICK-{wo_no[-4:]}{random.randint(10,99)}"
            need_qty = planned_qty * usage
            pick_qty = need_qty if random.random()<0.9 else need_qty * random.uniform(0.8,1.0)
            warehouse = random.choice(["WH01","WH03"])
            pick_date = datetime.strptime(a_start, "%Y-%m-%d") if a_start else datetime.strptime(p_start, "%Y-%m-%d")
            pick_date = pick_date - timedelta(days=random.randint(0,5))
            operator = random.choice(warehouse_staff)
            operator_dept = employee_info[operator]["部门"]
            picks.append((pick_no, wo_no, child, round(pick_qty,2), warehouse, operator, operator_dept,
                          pick_date.strftime("%Y-%m-%d")))

picks_df = pd.DataFrame(picks,
                        columns=["领料单号", "关联工单", "物料编码", "领料数量", "仓库", "操作人", "操作人部门", "领料日期"])
save_with_prefix(picks_df, FILE_ORDER[2] + ".xlsx", 3)

# ================== 4. 生产退料单 ==================
returns = []
for wo in work_orders:
    wo_no = wo[0]
    if random.random() < 0.1:
        pick_records = [p for p in picks if p[1] == wo_no]
        if pick_records and warehouse_staff:
            pick = random.choice(pick_records)
            return_no = f"PMR-{wo_no[-4:]}"
            return_qty = round(pick[3] * random.uniform(0.05, 0.2), 2)
            return_date = datetime.strptime(pick[7], "%Y-%m-%d") + timedelta(days=2)
            operator = random.choice(warehouse_staff)
            operator_dept = employee_info[operator]["部门"]
            returns.append((return_no, wo_no, pick[2], return_qty, pick[4], operator, operator_dept,
                            return_date.strftime("%Y-%m-%d")))

returns_df = pd.DataFrame(returns,
                          columns=["退料单号", "关联工单", "物料编码", "退料数量", "仓库", "操作人", "操作人部门", "退料日期"])
save_with_prefix(returns_df, FILE_ORDER[3] + ".xlsx", 4)

# ================== 5. 生产入库单 ==================
entries = []
for wo in work_orders:
    wo_no, source_so, product, planned_qty, completed_qty, good_qty, defect_qty, p_start, p_end, a_start, a_end, workshop, equipment, status = wo
    if status == "已完成" and good_qty > 0 and warehouse_staff:
        entry_no = f"PENT-{wo_no[-4:]}"
        entry_qty = good_qty
        entry_date = datetime.strptime(a_end, "%Y-%m-%d") if a_end else datetime.strptime(p_end, "%Y-%m-%d")
        operator = random.choice(warehouse_staff)
        operator_dept = employee_info[operator]["部门"]
        entries.append((entry_no, wo_no, product, entry_qty, "WH02", operator, operator_dept,
                        entry_date.strftime("%Y-%m-%d")))

entries_df = pd.DataFrame(entries,
                          columns=["入库单号", "关联工单", "产品编码", "入库数量", "仓库", "操作人", "操作人部门", "入库日期"])
save_with_prefix(entries_df, FILE_ORDER[4] + ".xlsx", 5)

print("\n生产管理模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)