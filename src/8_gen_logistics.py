# -*- coding: utf-8 -*-
"""
外贸物流模块 - 生成3个核心表单，直接添加数字前缀
包含：物流委托单、物流跟踪表、报关资料表
输出路径：D:\Trade-ERP-Learning\output\8.外贸物流\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

# 固定随机种子
random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\8.外贸物流"
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
    "物流委托单",
    "物流跟踪表",
    "报关资料表"
]

# ================== 读取基础数据 ==================
base_dir = r"D:\Trade-ERP-Learning\output\1.基础数据"

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

# 仓储物流部在职员工（物流专员、单证员）
logistics_staff = [eid for eid, info in employee_info.items()
                   if info["部门"] == "仓储物流部"
                   and info["状态"] == "在职"
                   and info["职位"] in ["物流专员", "单证员"]]
print(f"仓储物流部物流专员/单证员在职人数: {len(logistics_staff)}")
if not logistics_staff:
    print("警告: 仓储物流部无符合条件的物流专员/单证员！")

# 从销售模块获取外贸销售订单
sales_order_path = os.path.join(r"D:\Trade-ERP-Learning\output\3.销售管理", "2.销售订单表.xlsx")
if os.path.exists(sales_order_path):
    sales_orders_df = pd.read_excel(sales_order_path)
    if "订单类型" in sales_orders_df.columns:
        foreign_orders = sales_orders_df[sales_orders_df["订单类型"] == "外贸"]["订单号"].tolist()
        print(f"读取到 {len(foreign_orders)} 个外贸销售订单")
    else:
        foreign_orders = []
else:
    foreign_orders = []
    print("销售订单表不存在，将使用模拟外贸订单号")

if not foreign_orders:
    foreign_orders = [f"SO-2024{str(i).zfill(4)}" for i in range(1, 31) if i % 3 == 0]
    print(f"使用 {len(foreign_orders)} 个模拟外贸订单号")

# ================== 1. 物流委托单 ==================
logistics_orders = []
statuses = ["待审核", "已订舱", "已装箱", "已离港", "已完成"]
forwarders = ["全球物流有限公司", "中外运", "德迅"]
shipping_co = ["马士基", "中远", "地中海航运"]
container_types = ["20GP", "40HQ", "40GP"]

for i, so_no in enumerate(foreign_orders, start=1):
    log_no = f"LOG-2024{str(i).zfill(4)}"
    forwarder = random.choice(forwarders)
    shipping = random.choice(shipping_co)
    booking = f"BOOK{random.randint(1000,9999)}"
    vessel = f"{random.choice(['MAERSK','COSCO'])} {random.randint(2401,2412)}"
    bl_no = f"B/L{random.randint(10000,99999)}" if random.random()<0.8 else ""
    container_type = random.choice(container_types)
    container_qty = random.randint(1, 3)
    etd = random_date(datetime(2024,1,1), datetime(2024,12,31))
    eta = etd + timedelta(days=random.randint(15, 30))
    dest = random.choice(["洛杉矶", "汉堡", "釜山", "新加坡", "悉尼"])
    status = random.choice(statuses)
    logistics_orders.append((log_no, so_no, forwarder, shipping, booking, vessel, bl_no,
                             container_type, container_qty, etd.strftime("%Y-%m-%d"), eta.strftime("%Y-%m-%d"),
                             dest, status))

logistics_df = pd.DataFrame(logistics_orders,
                            columns=["委托单号", "关联销售订单", "货代公司", "船公司", "订舱号", "船名航次", "提单号(B/L)",
                                     "集装箱类型", "集装箱数量", "预计离港日(ETD)", "预计到港日(ETA)", "目的港", "状态"])
save_with_prefix(logistics_df, FILE_ORDER[0] + ".xlsx", 1)

# ================== 2. 物流跟踪表 ==================
trackings = []
nodes = ["已订舱", "已装箱", "已报关", "已离港", "已到港", "已提货"]
for i, log in enumerate(logistics_orders, start=1):
    if not logistics_staff:
        break
    log_no = log[0]
    etd = datetime.strptime(log[9], "%Y-%m-%d")
    node_times = [etd + timedelta(days=offset) for offset in [-5, -2, 0, 2, 20, 22]]
    num_nodes = random.randint(3, 5)
    for j in range(num_nodes):
        track_no = f"TRK-{i}{j}"
        node = nodes[j]
        track_time = node_times[j].strftime("%Y-%m-%d %H:%M:%S")
        location = random.choice(["上海", "宁波", "深圳", "香港", log[11]]) if "离港" not in node else "在途"
        operator = random.choice(logistics_staff)
        operator_dept = employee_info[operator]["部门"]
        trackings.append((track_no, log_no, track_time, node, location, "", operator, operator_dept))

trackings_df = pd.DataFrame(trackings,
                            columns=["跟踪记录号", "关联物流委托单", "状态时间", "状态节点", "当前位置", "备注",
                                     "操作人", "操作人部门"])
save_with_prefix(trackings_df, FILE_ORDER[1] + ".xlsx", 2)

# ================== 3. 报关资料表 ==================
customs = []
hs_codes = ["84571010", "85011099", "84834090"]
ports = ["上海海关", "深圳海关", "宁波海关"]
for i, log in enumerate(logistics_orders, start=1):
    if random.random() < 0.7:
        doc_no = f"CUS-{i:04d}"
        so_no = log[1]
        log_no = log[0]
        hs = random.choice(hs_codes)
        port = random.choice(ports)
        customs.append((doc_no, so_no, log_no, hs, "申报要素示例", port,
                        "invoice.pdf", "packing.pdf", "contract.pdf", "declaration.pdf", "已提交"))

customs_df = pd.DataFrame(customs,
                          columns=["资料单号", "关联销售订单", "关联物流委托单", "HS编码", "申报要素", "出口口岸",
                                   "商业发票(附件)", "装箱单(附件)", "合同(附件)", "报关单(附件)", "状态"])
save_with_prefix(customs_df, FILE_ORDER[2] + ".xlsx", 3)

print("\n外贸物流模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)