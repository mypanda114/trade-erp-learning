# -*- coding: utf-8 -*-
"""
销售管理模块 - 生成8个核心表单，直接添加数字前缀
包含：销售报价单、销售订单表、销售订单明细表、发货单、发货明细表、销售退货单、客户跟进记录表、样品管理表
输出路径：D:\Trade-ERP-Learning\output\3.销售管理\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

# 固定随机种子
random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\3.销售管理"
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
    "销售报价单",
    "销售订单表",
    "销售订单明细表",
    "发货单",
    "发货明细表",
    "销售退货单",
    "客户跟进记录表",
    "样品管理表"
]

# ================== 读取基础数据 ==================
base_dir = r"D:\Trade-ERP-Learning\output\1.基础数据"

# 客户
customers_file = find_base_file(base_dir, "客户信息表.xlsx")
customers_df = pd.read_excel(customers_file)
customer_codes = customers_df["客户编码"].tolist()
print(f"读取到 {len(customer_codes)} 个客户")

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

# 销售部在职员工（根据部门名称）
sales_dept_name = "销售部"
sales_staff_ids = [eid for eid, info in employee_info.items() 
                   if info["部门"] == sales_dept_name and info["状态"] == "在职"]
print(f"销售部在职员工总数: {len(sales_staff_ids)}")
if not sales_staff_ids:
    print("警告: 销售部无在职员工！")

# 销售专员（包括销售专员、销售经理）
sales_specialist_ids = [eid for eid, info in employee_info.items()
                        if info["部门"] == sales_dept_name and info["状态"] == "在职"
                        and info["职位"] in ["销售专员", "销售经理"]]
print(f"符合条件的销售专员/经理人数: {len(sales_specialist_ids)}")

# 如果销售专员列表为空，则降级使用所有销售部在职员工
if not sales_specialist_ids:
    print("警告：销售部无符合条件的专员/经理，将使用所有销售部在职员工作为报价人")
    sales_specialist_ids = sales_staff_ids

# 跟单员（职位包含“跟单员”）
inside_sales_ids = [eid for eid, info in employee_info.items()
                    if info["部门"] == sales_dept_name and info["状态"] == "在职"
                    and "跟单员" in info["职位"]]
print(f"符合条件的跟单员人数: {len(inside_sales_ids)}")
if not inside_sales_ids:
    print("警告：销售部无符合条件的跟单员，将使用所有销售部在职员工作为跟单员")
    inside_sales_ids = sales_staff_ids

# ================== 1. 销售报价单 ==================
quotation_data = []
for i in range(1, 16):
    if not sales_specialist_ids:
        break
    quot_no = f"Q-2024{str(i).zfill(4)}"
    cust = random.choice(customer_codes)
    prod = random.choice(product_codes)
    qty = random.randint(1, 10)
    price = random.randint(500, 50000)
    valid_until = (datetime(2024,12,31)).date()
    quot_by = random.choice(sales_specialist_ids)
    quot_by_dept = employee_info[quot_by]["部门"]
    quotation_data.append((quot_no, cust, prod, qty, price, valid_until.strftime("%Y-%m-%d"), 
                           quot_by, quot_by_dept))

quotations = pd.DataFrame(quotation_data, 
                          columns=["报价单号", "客户", "产品", "数量", "单价", "有效期", "报价人", "报价人部门"])
save_with_prefix(quotations, FILE_ORDER[0] + ".xlsx", 1)

# ================== 2. 销售订单表 ==================
sales_orders = []
sales_order_details = []
order_statuses = ["待审核", "已审核", "执行中", "已完成", "已取消"]
for i in range(1, 31):
    if not sales_specialist_ids or not inside_sales_ids:
        break
    order_no = f"SO-2024{str(i).zfill(4)}"
    customer = random.choice(customer_codes)
    salesperson = random.choice(sales_specialist_ids)
    salesperson_dept = employee_info[salesperson]["部门"]
    insider = random.choice(inside_sales_ids)
    insider_dept = employee_info[insider]["部门"]
    order_type = "外贸" if random.random() < 0.4 else "内销"
    
    # 生成下单日期（订单生成日期，早于要求交期）
    order_date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31)).date()
    req_date = order_date + timedelta(days=random.randint(7, 30))  # 要求交期在下单后7-30天
    prom_date = req_date - timedelta(days=random.randint(2, 10))   # 承诺交期早于要求交期
    
    if order_type == "外贸":
        trade_term = random.choice(["FOB", "CIF", "EXW"])
        dest_port = random.choice(["洛杉矶", "汉堡", "釜山", "新加坡", "悉尼"])
        freight_req = random.choice(["订舱需20尺柜", "订舱需40尺柜", "拼箱", ""])
        mark = random.choice(["ABC123", "XYZ789", "CUST01", ""])
        special_packing = random.choice(["木箱包装", "熏蒸木箱", "防潮包装", ""])
        lc = f"LC{random.randint(100000, 999999)}" if random.random() < 0.3 else ""
    else:
        trade_term = dest_port = freight_req = mark = special_packing = lc = ""

    num_products = random.randint(1, 3)
    selected_products = random.sample(product_codes, num_products)
    total_amount = 0
    details = []
    for prod in selected_products:
        qty = random.randint(1, 20)
        unit_price = random.randint(500, 50000)
        amount = qty * unit_price
        total_amount += amount
        details.append((order_no, prod, qty, unit_price, amount))
    status = random.choices(order_statuses, weights=[10,30,40,15,5])[0]
    sales_orders.append((order_no, customer, salesperson, salesperson_dept, insider, insider_dept, 
                         order_type, order_date.strftime("%Y-%m-%d"), req_date.strftime("%Y-%m-%d"), 
                         prom_date.strftime("%Y-%m-%d"), random.choice(["月结30天", "款到发货", "T/T"]),
                         trade_term, dest_port, freight_req, mark, special_packing, lc, total_amount, status))
    sales_order_details.extend(details)

so_df = pd.DataFrame(sales_orders, 
                     columns=["订单号", "客户", "销售专员", "销售专员部门", "跟单员", "跟单员部门", "订单类型", 
                              "下单日期", "要求交期", "承诺交期", "付款方式", "贸易术语", "目的港", "货代要求", 
                              "唛头", "特殊包装要求", "信用证号", "总金额", "订单状态"])
save_with_prefix(so_df, FILE_ORDER[1] + ".xlsx", 2)

# ================== 3. 销售订单明细表 ==================
so_detail_df = pd.DataFrame(sales_order_details, columns=["订单号", "产品编码", "数量", "单价", "金额"])
save_with_prefix(so_detail_df, FILE_ORDER[2] + ".xlsx", 3)

# ================== 4. 发货单和明细 ==================
shipments = []
shipment_details = []
shipment_statuses = ["运输中", "已签收"]
for i, (order_no, *_) in enumerate(sales_orders, start=1):
    if random.random() < 0.6:
        ship_no = f"SHP-2024{str(i).zfill(4)}"
        order_details = [d for d in sales_order_details if d[0] == order_no]
        ship_date = datetime.strptime(sales_orders[i-1][9], "%Y-%m-%d") + timedelta(days=random.randint(1,5))
        status = random.choice(shipment_statuses)
        sign_date = ship_date + timedelta(days=random.randint(3,7)) if status == "已签收" else ""
        shipments.append((ship_no, order_no, ship_date.strftime("%Y-%m-%d"), random.choice(["顺丰速运", "德邦物流", "中通"]),
                          f"SF{random.randint(100000000, 999999999)}", status, sign_date if sign_date else ""))
        for _, prod, qty, _, _ in order_details:
            ship_qty = random.randint(1, qty) if random.random() < 0.3 else qty
            shipment_details.append((ship_no, prod, ship_qty))

shipments_df = pd.DataFrame(shipments, columns=["发货单号", "关联销售订单", "发货日期", "物流公司", "运单号", "发货状态", "签收日期"])
save_with_prefix(shipments_df, FILE_ORDER[3] + ".xlsx", 4)

shipment_details_df = pd.DataFrame(shipment_details, columns=["发货单号", "产品编码", "数量"])
save_with_prefix(shipment_details_df, FILE_ORDER[4] + ".xlsx", 5)

# ================== 5. 销售退货单 ==================
returns = []
for i, (order_no, *_) in enumerate(sales_orders, start=1):
    if random.random() < 0.05:
        return_no = f"SR-2024{str(i).zfill(4)}"
        order_details = [d for d in sales_order_details if d[0] == order_no]
        if order_details:
            prod = random.choice([d[1] for d in order_details])
            return_qty = random.randint(1, 2)
            reason = "质量瑕疵"
            return_date = datetime.strptime(sales_orders[i-1][8], "%Y-%m-%d") + timedelta(days=30)
            returns.append((return_no, order_no, prod, return_qty, reason, return_date.strftime("%Y-%m-%d")))

returns_df = pd.DataFrame(returns, columns=["退货单号", "关联销售订单", "产品编码", "数量", "原因", "退货日期"])
save_with_prefix(returns_df, FILE_ORDER[5] + ".xlsx", 6)

# ================== 6. 客户跟进记录表 ==================
followups = []
followup_methods = ["电话", "邮件", "视频会议"]
for i, (order_no, customer, *_) in enumerate(sales_orders, start=1):
    for j in range(random.randint(1, 3)):
        if not sales_staff_ids:
            break
        follow_no = f"F-{i}{j}"
        follow_date = random_date(datetime(2024, 1, 1), datetime(2024, 12, 31)).date()
        follower = random.choice(sales_staff_ids)
        follower_dept = employee_info[follower]["部门"]
        method = random.choice(followup_methods)
        content = "跟进内容示例"
        next_plan = "下一步计划"
        next_follow = follow_date + timedelta(days=7)
        followups.append((follow_no, customer, follower, follower_dept, follow_date.strftime("%Y-%m-%d"), method,
                          content, next_plan, next_follow.strftime("%Y-%m-%d"), order_no))

followups_df = pd.DataFrame(followups, columns=["记录编号", "客户", "跟进人", "跟进人部门", "跟进日期", "跟进方式", 
                                                 "跟进内容", "下一步计划", "下次跟进日期", "关联销售订单"])
save_with_prefix(followups_df, FILE_ORDER[6] + ".xlsx", 7)

# ================== 7. 样品管理表 ==================
samples = []
sample_statuses = ["准备中", "已寄出", "已反馈"]
for i in range(1, 9):
    if not sales_specialist_ids:
        break
    samp_no = f"SMP-2024{str(i).zfill(4)}"
    cust = random.choice(customer_codes)
    prod = random.choice(product_codes)
    qty = 1
    apply_date = random_date(datetime(2024,1,1), datetime(2024,12,31)).date()
    send_date = apply_date + timedelta(days=random.randint(2, 10)) if random.random() < 0.7 else None
    status = "已寄出" if send_date else "准备中"
    express_no = f"SF{random.randint(100000,999999)}" if send_date else ""
    feedback = "客户反馈良好" if send_date and random.random()<0.5 else ""
    sales_rep = random.choice(sales_specialist_ids)
    sales_rep_dept = employee_info[sales_rep]["部门"]
    samples.append((samp_no, cust, sales_rep, sales_rep_dept, prod, qty, apply_date.strftime("%Y-%m-%d"),
                    send_date.strftime("%Y-%m-%d") if send_date else "", express_no, feedback, status))

samples_df = pd.DataFrame(samples, columns=["样品单号", "客户", "销售专员", "销售专员部门", "产品编码", "样品数量", 
                                            "申请日期", "寄出日期", "快递单号", "客户反馈", "状态"])
save_with_prefix(samples_df, FILE_ORDER[7] + ".xlsx", 8)

print("\n销售管理模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)