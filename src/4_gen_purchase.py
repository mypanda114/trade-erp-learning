# -*- coding: utf-8 -*-
"""
采购管理模块 - 生成5个核心表单，直接添加数字前缀
包含：采购申请单、采购订单表、采购订单明细表、采购入库单、采购退货单
输出路径：D:\Trade-ERP-Learning\output\4.采购管理\
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\4.采购管理"
os.makedirs(output_dir, exist_ok=True)

def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))

def find_base_file(base_dir, target_name):
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

FILE_ORDER = ["采购申请单", "采购订单表", "采购订单明细表", "采购入库单", "采购退货单"]

base_dir = r"D:\Trade-ERP-Learning\output\1.基础数据"

# 供应商
suppliers_file = find_base_file(base_dir, "供应商信息表.xlsx")
suppliers_df = pd.read_excel(suppliers_file)
supplier_codes = suppliers_df["供应商编码"].tolist()
print(f"读取到 {len(supplier_codes)} 个供应商")

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
        "状态": row["状态"],
        "入职日期": row["入职日期"]
    }

# 采购部在职员工
purchaser_ids = [eid for eid, info in employee_info.items()
                 if info["部门"] == "采购部" and info["状态"] == "在职"]
print(f"采购部在职员工总数: {len(purchaser_ids)}")
if not purchaser_ids:
    print("警告: 采购部无在职员工！")

# 仓储物流部在职员工
warehouse_ids = [eid for eid, info in employee_info.items()
                 if info["部门"] == "仓储物流部" and info["状态"] == "在职"]
print(f"仓储物流部在职员工总数: {len(warehouse_ids)}")
if not warehouse_ids:
    print("警告: 仓储物流部无在职员工！")

# 销售订单（用于关联）
sales_order_path = os.path.join(r"D:\Trade-ERP-Learning\output\3.销售管理", "2.销售订单表.xlsx")
if os.path.exists(sales_order_path):
    sales_orders_df = pd.read_excel(sales_order_path)
    sales_orders = sales_orders_df["订单号"].tolist()
    print(f"读取到 {len(sales_orders)} 个销售订单")
else:
    sales_orders = []
    print("销售订单表不存在，将使用模拟订单号")

# ================== 1. 采购申请单 ==================
pr_data = []
for i in range(1, 16):
    if not purchaser_ids:
        break
    pr_no = f"PR-{i:04d}"
    req_by = random.choice(purchaser_ids)
    req_by_dept = employee_info[req_by]["部门"]
    prod = random.choice(product_codes)
    qty = random.randint(10, 1000)
    req_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
    pr_data.append((pr_no, req_by, req_by_dept, prod, qty, req_date.strftime("%Y-%m-%d"), "生产备料"))

pr_df = pd.DataFrame(pr_data, columns=["申请单号", "申请人", "申请人部门", "物料编码", "数量", "需求日期", "用途"])
save_with_prefix(pr_df, FILE_ORDER[0] + ".xlsx", 1)

# ================== 2. 采购订单表 ==================
purchase_orders = []
purchase_order_details = []
po_statuses = ["待审核", "已审核", "执行中", "已完成"]

for i in range(1, 21):
    po_no = f"PO-2024{str(i).zfill(4)}"
    supplier = random.choice(supplier_codes)
    # 生成下单日期
    order_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
    req_delivery = order_date + timedelta(days=random.randint(7, 30))  # 要求到货日在下单后7-30天
    related_so = random.choice(sales_orders) if sales_orders and random.random() < 0.3 else ""
    total = 0
    num_items = random.randint(1, 4)
    items = []
    for j in range(num_items):
        prod = random.choice(product_codes)
        qty = random.randint(5, 200)
        price = random.randint(10, 2000)
        amt = qty * price
        total += amt
        items.append((po_no, prod, qty, price, amt))
    status = random.choices(po_statuses, weights=[10,30,40,20])[0]
    purchase_orders.append((po_no, supplier, order_date.strftime("%Y-%m-%d"), req_delivery.strftime("%Y-%m-%d"),
                            related_so, "采购原因", total, status))
    purchase_order_details.extend(items)

po_df = pd.DataFrame(purchase_orders,
                     columns=["采购单号", "供应商", "下单日期", "要求到货日", "关联销售订单", "采购原因", "总金额", "状态"])
save_with_prefix(po_df, FILE_ORDER[1] + ".xlsx", 2)

po_detail_df = pd.DataFrame(purchase_order_details,
                            columns=["采购单号", "物料编码", "数量", "单价", "金额"])
save_with_prefix(po_detail_df, FILE_ORDER[2] + ".xlsx", 3)

# ================== 3. 采购入库单 ==================
receipts = []
for po_no, supplier, order_date, req_date, related_so, reason, total, status in purchase_orders:
    if random.random() < 0.8:
        items = [item for item in purchase_order_details if item[0] == po_no]
        for po_no, prod, qty, price, amt in items:
            if not warehouse_ids:
                break
            rec_no = f"PREC-{po_no[-4:]}{random.randint(10,99)}"
            rec_qty = qty if random.random() < 0.9 else random.randint(1, qty)
            # 入库日期应 >= 下单日期，可以在要求到货日前前后
            rec_date = datetime.strptime(req_date, "%Y-%m-%d") + timedelta(days=random.randint(-3, 10))
            # 确保不小于下单日期
            order_date_dt = datetime.strptime(order_date, "%Y-%m-%d")
            if rec_date < order_date_dt:
                rec_date = order_date_dt + timedelta(days=random.randint(0, 3))
            qc_result = random.choices(["合格", "不合格", "待检"], weights=[80,5,15])[0]
            operator = random.choice(warehouse_ids)
            operator_dept = employee_info[operator]["部门"]
            receipts.append((rec_no, po_no, prod, rec_qty, random.choice(["WH01","WH02"]),
                             f"BATCH{rec_date.strftime('%Y%m%d')}", qc_result, rec_date.strftime("%Y-%m-%d"),
                             operator, operator_dept))

receipts_df = pd.DataFrame(receipts,
                           columns=["入库单号", "关联采购订单", "物料编码", "入库数量", "仓库", "批次号",
                                    "质检结果", "入库日期", "操作人", "操作人部门"])
save_with_prefix(receipts_df, FILE_ORDER[3] + ".xlsx", 4)

# ================== 4. 采购退货单 ==================
returns = []
for i, (po_no, supplier, order_date, req_date, related_so, reason, total, status) in enumerate(purchase_orders, start=1):
    if random.random() < 0.05:
        items = [item for item in purchase_order_details if item[0] == po_no]
        if items:
            prod, qty, price, amt = random.choice([(it[1], it[2], it[3], it[4]) for it in items])
            return_no = f"PRET-{po_no[-4:]}"
            return_qty = random.randint(1, min(5, qty))
            reason = "质量问题"
            # 退货日期应在入库日期之后，简单取下单日期后一段时间
            return_date = datetime.strptime(req_date, "%Y-%m-%d") + timedelta(days=random.randint(10, 20))
            returns.append((return_no, po_no, prod, return_qty, reason, return_date.strftime("%Y-%m-%d")))

returns_df = pd.DataFrame(returns, columns=["退货单号", "关联采购订单", "物料编码", "数量", "原因", "退货日期"])
save_with_prefix(returns_df, FILE_ORDER[4] + ".xlsx", 5)

print("\n采购管理模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)