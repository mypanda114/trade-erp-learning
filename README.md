# 内外贸一体化管理系统（学习项目）

[![Status](https://img.shields.io/badge/status-archived-lightgrey.svg)]()

> ⚠️ **项目状态**：此项目已完成，不再维护，已归档。内容仅供学习参考，请勿用于生产环境。

[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![Code style](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![Docs](https://img.shields.io/badge/docs-passing-brightgreen)](docs/)

> 基于简道云平台的中小型制造业内外贸一体化管理解决方案  
> 涵盖销售、采购、生产、库存、外贸物流等 9 大模块，包含 42 个表单详细设计及 Python 模拟数据生成脚本。

---

## 📖 目录

- [项目简介](#项目简介)
- [项目特点](#项目特点)
- [快速开始](#快速开始)
  - [环境准备](#环境准备)
  - [生成模拟数据](#生成模拟数据)
- [项目结构](#项目结构)
- [文档列表](#文档列表)
- [技术栈](#技术栈)
- [许可证](#许可证)
- [免责声明](#免责声明)
- [返回顶部](#内外贸一体化管理系统学习项目)

---

## 项目简介

本项目为个人学习探索项目，旨在研究中小型制造业企业内外贸一体化管理系统的设计与实现。所有内容**仅用于学习研究目的**，不包含任何真实企业信息，不涉及商业用途。

- 设计了一套覆盖销售、采购、生产、库存、外贸物流、财务等 9 大模块的通用型管理系统方案。
- 包含 42 个核心表单的详细字段定义，可直接用于简道云平台搭建。
- 提供完整的 Python 脚本，可自动生成模拟数据（300 名员工、30 条销售订单等），数据严格遵循业务逻辑（部门归属、在职状态、日期时序）。

---

## 项目特点

- **内外贸一体化**：统一平台支持国内销售和出口贸易，差异化处理外贸物流、关务、收汇流程。
- **跟单效率提升**：为跟单员提供实时订单进度视图，关联生产、库存、发货信息。
- **数据关联完整**：所有业务单据通过唯一编码建立逻辑链接，实现全链路追溯。
- **模块化设计**：划分为主数据、技术管理、销售、采购、生产、质量、库存、外贸物流、财务 9 个模块，便于扩展。
- **可复现数据**：脚本固定随机种子，保证数据一致性，便于学习和测试。

---

## 快速开始

### 环境准备

1. **安装 Python 3.6+**（推荐 3.8 以上）
2. **克隆本项目**（或下载源码）
   ```bash
   git clone https://github.com/mypanda114/trade-erp-learning
   cd trade-erp-learning
   ```
3. **创建虚拟环境并安装依赖**
   ```bash
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1          # Windows PowerShell
   # source .venv/bin/activate            # macOS / Linux
   pip install -r requirements.txt
   ```

### 生成模拟数据

运行主控脚本，自动生成所有 42 个 Excel 文件（按模块存放在 `output/` 目录下）：

```bash
cd src
python generate_all.py
```

> **注意**：脚本中默认使用绝对路径 `D:\Trade-ERP-Learning\output`，如需在其他位置运行，请修改各模块脚本中的 `output_dir` 变量为实际路径，或改为相对路径。

生成的 Excel 文件可直接导入简道云平台（导入方法见 [简道云搭建指南](docs/04_简道云搭建指南.md)）。

---

## 项目结构

```
trade-erp-learning/
├── docs/                       # 核心文档
│   ├── 01_方案设计.md
│   ├── 02_表单详细设计.md
│   ├── 03_脚本规划.md
│   ├── 04_简道云搭建指南.md
│   └── 05_脚本生成提示词.md
├── src/                        # Python 脚本
│   ├── generate_all.py         # 主控脚本
│   ├── 1_gen_base.py           # 基础数据模块
│   ├── 2_gen_plm.py            # 技术管理模块
│   ├── 3_gen_sales.py          # 销售管理模块
│   ├── 4_gen_purchase.py       # 采购管理模块
│   ├── 5_gen_production.py     # 生产管理模块
│   ├── 6_gen_quality.py        # 质量管理模块
│   ├── 7_gen_inventory.py      # 库存管理模块
│   ├── 8_gen_logistics.py      # 外贸物流模块
│   └── 9_gen_finance.py        # 财务管理模块
├── output/                     # 生成的 Excel 文件（已加入 .gitignore）
├── examples/                   # 可选示例数据
├── .gitignore
├── LICENSE                     # MIT 许可证
├── README.md                   # 本文件
└── requirements.txt            # Python 依赖
```

---

## 文档列表

| 文档 | 说明 |
|------|------|
| [01_方案设计](docs/01_方案设计.md) | 系统整体设计目标、原则、业务流程、审批流程等 |
| [02_表单详细设计](docs/02_表单详细设计.md) | 42 个核心表单的字段级定义 |
| [03_脚本规划](docs/03_脚本规划.md) | 从零开始搭建 Python 环境及运行脚本的完整指南 |
| [04_简道云搭建指南](docs/04_简道云搭建指南.md) | 在简道云平台导入数据、配置关联、仪表盘的步骤 |
| [05_脚本生成提示词](docs/05_脚本生成提示词.md) | 向大模型请求生成全套脚本的提示词模板 |

---

## 技术栈

- **平台**：简道云（低代码应用构建）
- **脚本语言**：Python 3.6+
- **依赖库**：pandas、openpyxl
- **版本控制**：Git

---

## 许可证

本项目采用 **MIT 许可证**。详情请参阅 [LICENSE](LICENSE) 文件。

---

## 免责声明

- 本项目为独立完成的学习成果，非商业产品，不提供任何商业保证。
- 项目中的设计方案、脚本代码仅供参考，使用者需自行承担风险。
- 项目不包含任何商业平台的专有代码或资源。
- 如涉及相关平台商标，仅用于说明兼容性，所有权归各自所有者。
- 任何人不得将本项目用于商业用途。

---

[返回顶部](#内外贸一体化管理系统学习项目)
```
