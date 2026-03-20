# -*- coding: utf-8 -*-
"""
技术管理模块 - 生成3个核心表单，直接添加数字前缀
包含：BOM物料清单表、工艺路线表、工程变更记录表
输出路径：D:\Trade-ERP-Learning\output\2.技术管理\
每次运行将覆盖原有数据，但保持文件格式和命名规范。
"""

import pandas as pd
import os
import random
import re
from datetime import datetime, timedelta

# 固定随机种子
random.seed(42)

output_dir = r"D:\Trade-ERP-Learning\output\2.技术管理"
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
    "BOM物料清单表",
    "工艺路线表",
    "工程变更记录表"
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

# 研发部在职员工（提交人）
rd_staff = [eid for eid, info in employee_info.items()
            if info["部门"] == "研发部" and info["状态"] == "在职"]
print(f"研发部在职员工总数: {len(rd_staff)}")
if not rd_staff:
    print("警告: 研发部无在职员工！")

# 生产部及技术部经理级员工（审批人）
approver_candidates = [eid for eid, info in employee_info.items()
                       if info["状态"] == "在职"
                       and info["部门"] in ["生产部", "研发部"]
                       and info["职位"] in ["生产经理", "研发总监"]]
if not approver_candidates:
    approver_candidates = [eid for eid, info in employee_info.items()
                           if info["部门"] == "生产部"
                           and info["状态"] == "在职"
                           and info["职位"] == "生产经理"]

# ================== 1. BOM物料清单表 ==================
bom_data = [
    ("BOM-P001-V1", "P001", "P002", 2, "个", 10, "2024-01-01", "初始版本"),
    ("BOM-P001-V1", "P001", "P003", 1, "个", 20, "2024-01-01", "初始版本"),
    ("BOM-P001-V1", "P001", "P004", 4, "个", 30, "2024-01-01", "初始版本"),
    ("BOM-P002-V2", "P002", "P005", 0.05, "吨", 10, "2024-02-01", "电机外壳改用钢板"),
    ("BOM-P002-V2", "P002", "P006", 2, "公斤", 20, "2024-02-01", "塑料部件"),
    ("BOM-P007-V1", "P007", "P004", 2, "个", 10, "2024-02-10", "初始版本"),
    ("BOM-P008-V1", "P008", "P009", 1.5, "米", 10, "2024-02-15", "连接线"),
    ("BOM-P008-V1", "P008", "P010", 1, "个", 20, "2024-02-15", "电源模块"),
    ("BOM-P007-V2", "P007", "P008", 1, "台", 30, "2024-03-01", "升级版含变频器"),
]
bom = pd.DataFrame(bom_data,
                   columns=["BOM编号", "父项产品编码", "子项物料编码", "用量", "单位", "工序序号", "生效日期", "版本说明"])
save_with_prefix(bom, FILE_ORDER[0] + ".xlsx", 1)

# ================== 2. 工艺路线表 ==================
routing_data = [
    ("P001", 10, "机加工", "金工车间", 120, "五轴加工中心", "尺寸公差±0.01mm"),
    ("P001", 20, "装配", "装配车间", 60, "手动", "紧固扭矩符合规范"),
    ("P002", 10, "绕线", "电机车间", 30, "绕线机", "电阻值测试"),
    ("P002", 20, "组装", "电机车间", 45, "手动", "绝缘测试"),
    ("P007", 10, "箱体加工", "金工车间", 90, "数控车床", "同轴度0.02mm"),
    ("P007", 20, "装配", "装配车间", 50, "手动", "间隙调整"),
    ("P008", 10, "插件", "电子车间", 20, "自动插件机", "元件位置正确"),
    ("P008", 20, "测试", "电子车间", 15, "测试台", "功能测试"),
]
routing = pd.DataFrame(routing_data,
                       columns=["产品编码", "工序序号", "工序名称", "工作中心", "标准工时(分钟)", "设备要求", "质检要求"])
save_with_prefix(routing, FILE_ORDER[1] + ".xlsx", 2)

# ================== 3. 工程变更记录表 ==================
eco_data = [
    ("ECO-001", "P002", "材料变更", "将轴承型号由6204改为6205", "供应商升级", "V2.1", "2024-03-01", "eco001.pdf", "已批准",
     random.choice(rd_staff) if rd_staff else None,
     random.choice(approver_candidates) if approver_candidates else None,
     "2024-02-25"),
    ("ECO-002", "P001", "工艺优化", "增加热处理工序", "提高强度", "V1.1", "2024-04-01", "eco002.pdf", "审核中",
     random.choice(rd_staff) if rd_staff else None, "", ""),
]

# 添加部门列
eco_records = []
for row in eco_data:
    record = list(row)
    submitter = row[9] if len(row) > 9 else None
    submitter_dept = employee_info[submitter]["部门"] if submitter and submitter in employee_info else None
    approver = row[10] if len(row) > 10 else None
    approver_dept = employee_info[approver]["部门"] if approver and approver in employee_info else None
    new_record = record[:9] + [submitter_dept] + [record[9] if len(record)>9 else None] + [approver_dept] + record[10:]
    eco_records.append(new_record)

eco = pd.DataFrame(eco_records,
                   columns=["变更单号", "变更产品编码", "变更类型", "变更内容", "变更原因", "涉及BOM版本", "生效日期", "附件", "状态",
                            "提交人部门", "提交人", "审批人部门", "审批人", "审批日期"])
# 调整列顺序
eco = eco[["变更单号", "变更产品编码", "变更类型", "变更内容", "变更原因", "涉及BOM版本", "生效日期", "附件", "状态",
           "提交人", "提交人部门", "审批人", "审批人部门", "审批日期"]]
save_with_prefix(eco, FILE_ORDER[2] + ".xlsx", 3)

print("\n技术管理模块全部生成完毕！")
print("文件已按数字前缀顺序保存在：", output_dir)