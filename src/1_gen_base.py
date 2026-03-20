# -*- coding: utf-8 -*-
"""
基础数据模块 - 生成9个核心表单，直接添加数字前缀
包含：部门表、职位表、班组表、员工表、客户信息表、供应商信息表、产品主数据表、仓库仓位表、设备管理表
输出路径：D:\Trade-ERP-Learning\output\1.基础数据\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
from datetime import datetime, timedelta

# 固定随机种子，保证结果可复现（可注释掉以生成不同数据）
random.seed(42)

# 输出目录
output_dir = r"D:\Trade-ERP-Learning\output\1.基础数据"
os.makedirs(output_dir, exist_ok=True)

# ================== 辅助函数 ==================
def random_date(start, end):
    """生成随机日期"""
    return start + timedelta(days=random.randint(0, (end - start).days))

def generate_chinese_name():
    """生成中文名（姓氏+两位小写字母）"""
    first_names = ['张', '王', '李', '刘', '陈', '杨', '赵', '黄', '周', '吴',
                   '徐', '孙', '胡', '朱', '高', '林', '何', '郭', '马', '罗']
    first = random.choice(first_names)
    suffix = ''.join(random.choices('abcdefghijklmnopqrstuvwxyz', k=2))
    return first + suffix

def generate_employee_id(index):
    """生成员工工号，格式 EMP001"""
    return f"EMP{index:03d}"

def save_with_prefix(df, filename, prefix_num):
    """保存 DataFrame 到 Excel，文件名带数字前缀"""
    final_filename = f"{prefix_num}.{filename}"
    filepath = os.path.join(output_dir, final_filename)
    df.to_excel(filepath, index=False)
    print(f"已生成：{final_filename}")

# ================== 部门配置（用于生成辅助表和员工表） ==================
dept_config = {
    "总经办": {"positions": ["总经理"], "teams": []},
    "生产部": {"positions": ["生产总监", "生产经理", "生产主管", "班组长", "操作工"],
              "teams": ["金工班", "装配班", "电机班", "注塑班"]},
    "销售部": {"positions": ["销售总监", "销售经理", "销售专员", "跟单员（内贸）", "跟单员（外贸）"],
              "teams": []},
    "采购部": {"positions": ["采购经理", "采购主管", "采购专员"], "teams": []},
    "质检部": {"positions": ["质检经理", "质检主管", "质检员"], "teams": []},
    "仓储物流部": {"positions": ["仓储经理", "仓库管理员", "物流专员", "单证员"], "teams": []},
    "财务部": {"positions": ["财务经理", "会计", "出纳"], "teams": []},
    "行政部": {"positions": ["行政经理", "行政专员"], "teams": []},
    "人事部": {"positions": ["人事经理", "人事专员"], "teams": []},
    "研发部": {"positions": ["研发总监", "研发工程师", "技术员"], "teams": []}
}

# ================== 1. 部门表 ==================
dept_data = []
for idx, (dept_name, _) in enumerate(dept_config.items(), start=1):
    dept_id = f"DEP{idx:03d}"
    dept_data.append([dept_id, dept_name])
dept_df = pd.DataFrame(dept_data, columns=["部门ID", "部门名称"])
save_with_prefix(dept_df, "部门表.xlsx", 1)

# ================== 2. 职位表 ==================
pos_data = []
pos_idx = 1
for dept_name, config in dept_config.items():
    for pos_name in config["positions"]:
        pos_id = f"POS{pos_idx:03d}"
        pos_data.append([pos_id, pos_name, dept_name])
        pos_idx += 1
pos_df = pd.DataFrame(pos_data, columns=["职位ID", "职位名称", "所属部门"])
save_with_prefix(pos_df, "职位表.xlsx", 2)

# ================== 3. 班组表（仅生产部有班组） ==================
team_data = []
team_idx = 1
for dept_name, config in dept_config.items():
    if config["teams"]:
        for team_name in config["teams"]:
            team_id = f"TEAM{team_idx:03d}"
            team_data.append([team_id, team_name, dept_name])
            team_idx += 1
team_df = pd.DataFrame(team_data, columns=["班组ID", "班组名称", "所属部门"])
save_with_prefix(team_df, "班组表.xlsx", 3)

# ================== 4. 员工表（300人） ==================
# 各部门人数分配（总和300）
dept_counts = {
    "总经办": 1,
    "生产部": 134,
    "销售部": 36,
    "采购部": 15,
    "质检部": 24,
    "仓储物流部": 30,
    "财务部": 15,
    "行政部": 9,
    "人事部": 9,
    "研发部": 27
}

employee_records = []
for dept_name, count in dept_counts.items():
    config = dept_config[dept_name]
    positions = config["positions"]
    teams = config.get("teams", [])
    
    # 平均分配岗位人数
    pos_counts = [count // len(positions)] * len(positions)
    remainder = count - sum(pos_counts)
    for i in range(remainder):
        pos_counts[i] += 1
    
    for pos, pos_count in zip(positions, pos_counts):
        for _ in range(pos_count):
            name = generate_chinese_name()
            
            # 根据部门和岗位设置离职概率
            if dept_name == "总经办":
                # 总经理始终在职
                status = "在职"
            elif dept_name == "生产部":
                if pos == "操作工":
                    status = "离职" if random.random() < 0.20 else "在职"
                elif pos == "班组长":
                    status = "离职" if random.random() < 0.15 else "在职"
                else:  # 生产总监、生产经理、生产主管
                    status = "离职" if random.random() < 0.05 else "在职"
            else:  # 其他部门所有岗位统一5%离职率
                status = "离职" if random.random() < 0.05 else "在职"
            
            hire_date = random_date(datetime(2018, 1, 1), datetime(2024, 6, 30)).date()
            phone = f"1{random.randint(3,9)}{random.randint(0,9):08d}"
            email = f"{name.lower()}@company.com"
            jiandaoyun_account = f"{name.lower()}_jd"
            
            # 所属班组：仅生产部的班组长/操作工有班组
            team = ""
            if dept_name == "生产部" and pos in ["班组长", "操作工"] and teams:
                team = random.choice(teams)
            
            employee_records.append({
                "工号": "",
                "姓名": name,
                "所属部门": dept_name,
                "所属班组": team,
                "职位": pos,
                "入职日期": hire_date.strftime("%Y-%m-%d"),
                "状态": status,
                "联系电话": phone,
                "邮箱": email,
                "简道云账号": jiandaoyun_account
            })

# 打乱顺序并分配工号
random.shuffle(employee_records)
for i, emp in enumerate(employee_records, start=1):
    emp["工号"] = generate_employee_id(i)

employees_df = pd.DataFrame(employee_records)
cols = ["工号", "姓名", "所属部门", "所属班组", "职位", "入职日期", "状态",
        "联系电话", "邮箱", "简道云账号"]
employees_df = employees_df[cols]
save_with_prefix(employees_df, "员工表.xlsx", 4)

# ================== 5. 客户信息表（10家） ==================
customers_data = [
    ("C001", "上海机械有限公司", "国内", "中国", "张三", "021-12345678", "zhang@shanghai.com", "月结30天", 500000, "合作中"),
    ("C002", "广东电子厂", "国内", "中国", "李四", "0755-87654321", "li@guangdong.com", "款到发货", 200000, "合作中"),
    ("C003", "ABC Trading Co., Ltd", "国外", "美国", "John Smith", "+1-555-1234", "john@abc.com", "T/T", 100000, "合作中"),
    ("C004", "Euro Tech GmbH", "国外", "德国", "Hans Mueller", "+49-30-123456", "hans@eurotech.de", "L/C", 150000, "潜在客户"),
    ("C005", "韩国重工", "国外", "韩国", "Kim", "+82-2-7890", "kim@korea.co.kr", "T/T", 300000, "合作中"),
    ("C006", "越南制造", "国外", "越南", "Nguyen", "+84-28-123456", "nguyen@vietnam.vn", "T/T", 80000, "合作中"),
    ("C007", "北京科技公司", "国内", "中国", "王伟", "010-87654321", "wang@beijing.com", "月结30天", 150000, "合作中"),
    ("C008", "重庆机械", "国内", "中国", "刘芳", "023-12345678", "liu@chongqing.com", "月结60天", 200000, "合作中"),
    ("C009", "Singapore Electronics", "国外", "新加坡", "Tan", "+65-6789-1234", "tan@sg.com", "T/T", 120000, "合作中"),
    ("C010", "Brazil Trading", "国外", "巴西", "Silva", "+55-11-12345678", "silva@brazil.com", "L/C", 90000, "潜在客户"),
]
customers = pd.DataFrame(customers_data,
                         columns=["客户编码", "名称", "类型", "国家", "联系人", "联系电话", "邮箱",
                                  "付款方式", "信用额度", "合作状态"])
save_with_prefix(customers, "客户信息表.xlsx", 5)

# ================== 6. 供应商信息表（8家） ==================
suppliers_data = [
    ("S001", "浙江钢铁集团", "原材料", "钢材", "王五", "0571-1111111", "wang@zjsteel.com", "月结60天", "合作中"),
    ("S002", "江苏塑料厂", "原材料", "塑料粒子", "赵六", "0512-2222222", "zhao@jsplastic.com", "预付30%", "合作中"),
    ("S003", "全球物流有限公司", "物流服务", "", "陈七", "010-3333333", "chen@global-log.com", "月结", "合作中"),
    ("S004", "华南电机", "外购件", "电机", "周八", "020-4444444", "zhou@motor.com", "货到付款", "合作中"),
    ("S005", "山东轴承厂", "外购件", "轴承", "孙九", "0531-5555555", "sun@bearing.com", "月结30天", "合作中"),
    ("S006", "天津电子", "外购件", "电子元件", "李华", "022-6666666", "li@tjelec.com", "月结30天", "合作中"),
    ("S007", "河北包装", "辅料", "包装材料", "赵岩", "0311-7777777", "zhao@hebei.com", "月结", "合作中"),
    ("S008", "上海工具", "外购件", "刀具", "钱进", "021-8888888", "qian@shanghai.com", "预付", "合作中"),
]
suppliers = pd.DataFrame(suppliers_data,
                         columns=["供应商编码", "名称", "类型", "主要物料", "联系人", "联系电话", "邮箱",
                                  "付款条件", "合作状态"])
save_with_prefix(suppliers, "供应商信息表.xlsx", 6)

# ================== 7. 产品主数据表（10种） ==================
products_data = [
    ("P001", "数控机床", "CNC-2000", "整机", "高精度数控机床", "钢铁/电子", 2000, "台", "V1.0", "2024-01-01", "启用"),
    ("P002", "伺服电机", "SM-100", "半成品", "200W伺服电机", "铜/磁钢", 5, "个", "V2.1", "2024-02-01", "启用"),
    ("P003", "PLC控制器", "PLC-X3", "配件", "可编程控制器", "电子元件", 0.5, "个", "V1.5", "2024-01-15", "启用"),
    ("P004", "轴承", "6204", "外购件", "深沟球轴承", "轴承钢", 0.2, "个", "V1.0", "2024-01-01", "启用"),
    ("P005", "钢板", "Q235 10mm", "原材料", "热轧钢板", "碳钢", 100, "吨", "V1.0", "2024-01-01", "启用"),
    ("P006", "ABS塑料颗粒", "ABS-757", "原材料", "注塑级", "ABS", 25, "公斤", "V1.2", "2024-03-01", "启用"),
    ("P007", "减速机", "RV40", "整机", "蜗轮蜗杆减速机", "铸铁", 15, "台", "V1.0", "2024-02-10", "启用"),
    ("P008", "变频器", "VFD-5.5", "外购件", "5.5kW变频器", "电子元件", 2, "台", "V1.1", "2024-02-15", "启用"),
    ("P009", "电线", "RVVP 2*1.0", "原材料", "屏蔽电缆", "铜/塑料", 0.1, "米", "V1.0", "2024-01-20", "启用"),
    ("P010", "开关电源", "S-120-24", "外购件", "24V 5A开关电源", "电子元件", 1.2, "个", "V1.0", "2024-03-05", "启用"),
]
products = pd.DataFrame(products_data,
                        columns=["产品编码", "名称", "型号", "类型", "规格参数", "材质", "重量", "单位",
                                 "版本", "生效日期", "状态"])
save_with_prefix(products, "产品主数据表.xlsx", 7)

# ================== 8. 仓库仓位表 ==================
warehouse_bins_data = [
    ("WH01", "原料仓", "原材料", "A-01"),
    ("WH01", "原料仓", "原材料", "A-02"),
    ("WH01", "原料仓", "原材料", "A-03"),
    ("WH02", "成品仓", "成品", "B-01"),
    ("WH02", "成品仓", "成品", "B-02"),
    ("WH03", "半成品仓", "半成品", "C-01"),
    ("WH03", "半成品仓", "半成品", "C-02"),
]
warehouse_bins = pd.DataFrame(warehouse_bins_data, columns=["仓库编码", "名称", "类型", "仓位"])
save_with_prefix(warehouse_bins, "仓库仓位表.xlsx", 8)

# ================== 9. 设备管理表（5台） ==================
equipments_data = [
    ("E001", "五轴加工中心", "DMG-50", "金工车间", "运行中", "2023-05-10", "2024-05-10", "2024-11-10"),
    ("E002", "注塑机", "Haitian-280", "注塑车间", "运行中", "2023-08-20", "2024-08-20", "2025-02-20"),
    ("E003", "数控车床", "CK6140", "金工车间", "维护中", "2022-11-15", "2024-02-15", "2024-08-15"),
    ("E004", "铣床", "X5032", "金工车间", "运行中", "2023-01-10", "2024-01-10", "2024-07-10"),
    ("E005", "磨床", "M1432", "金工车间", "运行中", "2023-06-20", "2024-06-20", "2024-12-20"),
]
equipments = pd.DataFrame(equipments_data,
                          columns=["设备编码", "名称", "型号", "所在车间", "状态",
                                   "购入日期", "上次保养日", "下次保养日"])
save_with_prefix(equipments, "设备管理表.xlsx", 9)

print("\n基础数据模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)