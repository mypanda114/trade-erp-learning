# -*- coding: utf-8 -*-
"""
财务管理模块 - 生成4个核心表单，直接添加数字前缀
包含：收款记录表、付款记录表、销项发票表、进项发票表
输出路径：D:\Trade-ERP-Learning\output\9.财务管理\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

# 固定随机种子
random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\9.财务管理"
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

# ================== 该模块文件顺序 ==================
FILE_ORDER = [
    "收款记录表",
    "付款记录表",
    "销项发票表",
    "进项发票表"
]

# ================== 读取基础数据 ==================
base_dir = r"D:\Trade-ERP-Learning\output\1.基础数据"

# 销售订单（从销售模块获取）
sales_order_path = os.path.join(r"D:\Trade-ERP-Learning\output\3.销售管理", "2.销售订单表.xlsx")
if os.path.exists(sales_order_path):
    sales_orders_df = pd.read_excel(sales_order_path)
    sales_orders = sales_orders_df["订单号"].tolist()
    # 将下单日期字符串转换为 date 对象
    sales_order_dates = {}
    for _, row in sales_orders_df.iterrows():
        try:
            # 尝试转换日期字符串
            order_date = datetime.strptime(row["下单日期"], "%Y-%m-%d").date()
        except:
            order_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
        sales_order_dates[row["订单号"]] = order_date
else:
    sales_orders = [f"SO-2024{str(i).zfill(4)}" for i in range(1, 31)]
    sales_order_dates = {so: random_date(datetime(2024,1,1), datetime(2024,12,31)).date() for so in sales_orders}
    print("销售订单表不存在，将使用模拟销售订单")

# 采购订单（从采购模块获取）
purchase_order_path = os.path.join(r"D:\Trade-ERP-Learning\output\4.采购管理", "2.采购订单表.xlsx")
if os.path.exists(purchase_order_path):
    purchase_orders_df = pd.read_excel(purchase_order_path)
    purchase_orders = purchase_orders_df["采购单号"].tolist()
    purchase_order_dates = {}
    for _, row in purchase_orders_df.iterrows():
        try:
            order_date = datetime.strptime(row["下单日期"], "%Y-%m-%d").date()
        except:
            order_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
        purchase_order_dates[row["采购单号"]] = order_date
else:
    purchase_orders = [f"PO-2024{str(i).zfill(4)}" for i in range(1, 21)]
    purchase_order_dates = {po: random_date(datetime(2024,1,1), datetime(2024,12,31)).date() for po in purchase_orders}
    print("采购订单表不存在，将使用模拟采购订单")

# ================== 1. 收款记录表 ==================
receipts = []
currencies = ["CNY", "USD"]
methods = ["T/T", "L/C", "现金"]
for so in sales_orders:
    if random.random() < 0.6:  # 60%的订单有收款
        rec_no = f"REC-{so[-4:]}"
        # 获取下单日期，若不存在则随机生成
        order_date = sales_order_dates.get(so)
        if order_date is None:
            order_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
        # 收款日期在下单后1-30天
        rec_date = order_date + timedelta(days=random.randint(1, 30))
        amount = random.randint(5000, 100000)
        currency = random.choice(currencies)
        method = random.choice(methods)
        receipts.append((rec_no, so, rec_date.strftime("%Y-%m-%d"), amount, currency, method, ""))

receipts_df = pd.DataFrame(receipts,
                           columns=["收款单号", "关联销售订单", "收款日期", "金额", "币种", "收款方式", "备注"])
save_with_prefix(receipts_df, FILE_ORDER[0] + ".xlsx", 1)

# ================== 2. 付款记录表 ==================
payments = []
for po in purchase_orders:
    if random.random() < 0.5:  # 50%的订单有付款
        pay_no = f"PAY-{po[-4:]}"
        order_date = purchase_order_dates.get(po)
        if order_date is None:
            order_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
        pay_date = order_date + timedelta(days=random.randint(1, 30))
        amount = random.randint(1000, 50000)
        payments.append((pay_no, po, pay_date.strftime("%Y-%m-%d"), amount, "T/T", ""))

payments_df = pd.DataFrame(payments,
                           columns=["付款单号", "关联采购订单", "付款日期", "金额", "付款方式", "备注"])
save_with_prefix(payments_df, FILE_ORDER[1] + ".xlsx", 2)

# ================== 3. 销项发票表 ==================
sales_invoices = []
for so in sales_orders:
    if random.random() < 0.5:  # 50%的订单有发票
        inv_no = f"INV-{so[-4:]}"
        order_date = sales_order_dates.get(so)
        if order_date is None:
            order_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
        inv_date = order_date + timedelta(days=random.randint(5, 40))
        amount = random.randint(5000, 100000)
        tax_no = f"91310115MA1H2{random.randint(1000,9999)}"
        sales_invoices.append((inv_no, so, inv_date.strftime("%Y-%m-%d"), amount, tax_no, ""))

sales_inv_df = pd.DataFrame(sales_invoices,
                            columns=["发票号", "关联销售订单", "开票日期", "金额", "税号", "备注"])
save_with_prefix(sales_inv_df, FILE_ORDER[2] + ".xlsx", 3)

# ================== 4. 进项发票表 ==================
purchase_invoices = []
for po in purchase_orders:
    if random.random() < 0.5:
        inv_no = f"PINV-{po[-4:]}"
        order_date = purchase_order_dates.get(po)
        if order_date is None:
            order_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
        inv_date = order_date + timedelta(days=random.randint(5, 40))
        amount = random.randint(1000, 50000)
        tax_no = f"91330105MA27{random.randint(1000,9999)}"
        purchase_invoices.append((inv_no, po, inv_date.strftime("%Y-%m-%d"), amount, tax_no, ""))

purchase_inv_df = pd.DataFrame(purchase_invoices,
                               columns=["发票号", "关联采购订单", "收票日期", "金额", "税号", "备注"])
save_with_prefix(purchase_inv_df, FILE_ORDER[3] + ".xlsx", 4)

print("\n财务管理模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)