# CDM 工具包 — 三层架构重新设计规格书

**日期**: 2026-05-30
**状态**: 已批准
**替代**: `2026-05-30-cdm-vba-python-toolkit-design.md`（旧版VBA全量设计）

---

## 1. 概述

基于对 DM（数据管理）用户日常工作流的深入分析，将 CDM 工具包从"VBA + Python 两套平行实现"重新设计为**三层架构**：表示层（VBA + Web UI）、逻辑层（Python 统一引擎）、数据层（Excel/SAS/CSV）。

### 核心洞察

DM 数据管理员日常接触 Excel 和 SAS 数据集。两类操作需要不同的工具形态：

- **即时操作**（选中 → 执行 → 看到结果）：必须在 Excel 内完成，VBA 是唯一选择
- **复杂处理**（多文件批处理、可视化报告、差异对比）：Web UI + Python 后端更优

### 设计目标

- VBA 和 Web **互补共存**，功能不重叠
- Python 引擎作为**唯一逻辑源**，VBA 和 Web 都调用/翻译同一套逻辑
- Web 应用作为本地 localhost 服务，类似 Jupyter Notebook 体验
- 支持企业内网桌面和个人灵活环境
- 新增 SAS 数据集读写和批量操作管道

---

## 2. 架构总览

### 三层模型

```
┌─────────────────────────────────────────────────┐
│              表示层 (Presentation)                │
│  ┌──────────────────┐  ┌──────────────────────┐ │
│  │  VBA 宏 (Excel内) │  │  Web UI (浏览器)      │ │
│  │  Alt+F8 即时操作   │  │  localhost:8520       │ │
│  │  单元格级交互      │  │  拖拽上传/批处理       │ │
│  │  批注/形状/超链接  │  │  可视化报告/对比       │ │
│  └──────────────────┘  └──────────────────────┘ │
├─────────────────────────────────────────────────┤
│              逻辑层 (Logic)                       │
│  ┌──────────────────────────────────────────────┐│
│  │          cdm_engine (Python 包)               ││
│  │  · Excel 读写 (openpyxl)                      ││
│  │  · SAS 读写 (pandas/pyreadstat)               ││
│  │  · 所有业务逻辑（12个功能模块）                 ││
│  │  · FastAPI REST 端点                          ││
│  └──────────────────────────────────────────────┘│
├─────────────────────────────────────────────────┤
│              数据层 (Data)                        │
│  Excel (.xlsx/.xls) · SAS (.sas7bdat) · CSV     │
│  本地文件系统                                    │
└─────────────────────────────────────────────────┘
```

### 关键技术决策

| 决策点 | 选择 | 理由 |
|--------|------|------|
| Web 部署模式 | 本地 localhost 服务 | 数据不出本地，零配置启动，不需要 IT 审批服务器 |
| 前端技术栈 | 原生 HTML/CSS/JS (ES Modules + Web Components) | 零构建、无 Node.js 依赖、`pip install` 即用 |
| 后端框架 | FastAPI | 异步支持、自动 OpenAPI 文档、文件上传处理优秀 |
| Python 引擎 | 统一 `cdm_engine/` 包 | VBA 翻译调用，Web 直接调用，单一逻辑源 |
| VBA 角色 | Excel 内即时操作 | 深度 Excel 集成（Comment、Shape、Hyperlink） |

---

## 3. 功能分配矩阵

### 3.1 VBA 专属（6个模块）— Excel 内即时操作

这些功能的核心价值在于"选中单元格 → 按键执行 → 立即看到结果"。离开 Excel 环境就失去意义。

| # | 模块 | 功能 | 保留理由 |
|---|------|------|---------|
| 1 | modTextStandardise | 文本术语标准化 | 单元格即时替换 + 视觉确认，Web 上传→处理→下载流程太长 |
| 2 | modHighlightDup | 重复值高亮 | 条件格式的即时可视化，颜色即信息 |
| 3 | modBlankRows | 空白行管理 | 行插入/删除是 Excel 原生操作，VBA 直接操作 Range |
| 4 | modTOC | TOC 目录生成 | Shapes、Hyperlinks、跨 sheet 导航是 Excel 专有功能 |
| 5 | modSheetManager | 工作表管理 | 直接操作 Workbook 对象，Web 需要上传整个工作簿 |
| 6 | modTrackChanges | 修改痕迹跟踪 | Excel Comment 对象只有 VBA/COM 能完整操作 |

### 3.2 Web 专属（4个模块）— 复杂可视化 & 批处理

这些功能在 VBA 中体验差（单线程卡顿、UI 简陋），Web 端有天然优势。

| # | 模块 | 功能 | Web 优势 |
|---|------|------|---------|
| 7 | 工作簿合并工作台 | 多文件合并为一个 | 拖拽上传 + 预览 + 进度条 + 合并策略配置 |
| 8 | 工作表拆分工作台 | 按列值拆分为多 sheet | 异步处理 + 进度反馈 + 拆分前预览，大数据不卡顿 |
| 9 | 特殊字符扫描报告 | EDC 合规检查 | 交互式结果表格 + 单元格定位 + Rave/Clinflash 报告一键导出 |
| 10 | 文件夹浏览器 | 文件目录生成 | 树形组件天然适合文件夹浏览，远胜 Excel 行内缩进 |

### 3.3 双端共享（2个模块）— 核心计算逻辑

纯计算功能，Python 写核心逻辑，VBA 和 Web 各自调用。

| # | 模块 | VBA 场景 | Web 场景 |
|---|------|---------|---------|
| 11 | 日期计算器 | 选中 → +30天（轻量即时） | 批量列运算（大批量） |
| 12 | 实验室单位转换 | 选中 → 即时转换 | 批量转换 + 管理转换表 |

### 汇总

| 分类 | 数量 | 模块 |
|------|------|------|
| 🟢 VBA 专属 | 6 | 文本标准化、重复高亮、空白行、TOC、工作表管理、修改痕迹 |
| 🔵 Web 专属 | 4 | 工作簿合并、工作表拆分、特殊字符、文件目录 |
| 🟡 双端共享 | 2 | 日期计算、单位转换 |
| **总计** | **12** | |

---

## 4. Web 应用设计

### 4.1 启动方式

```bash
# 方式一：命令行
pip install cdm-tools
cdm-tools serve          # → 自动打开 http://localhost:8520

# 方式二：从 Excel VBA
DM_LaunchWeb             # → Shell 启动 server.py → 自动打开浏览器
```

### 4.2 页面结构

| 页面 | 路由 | API 端点 | 功能描述 |
|------|------|---------|---------|
| 仪表盘 | `/` | — | 功能卡片导航 + 最近处理记录 + VBA/引擎状态 |
| 工作簿合并 | `/merge` | `POST /api/merge` | 拖拽上传多文件 → 预览 → 配置合并策略 → 下载 |
| 工作表拆分 | `/split` | `POST /api/split` | 上传文件 → 选拆分列 → 预览分组 → 下载 |
| 特殊字符扫描 | `/scan` | `POST /api/scan-chars` | 上传文件 → 扫描 → 交互式结果表格 → EDC报告导出 |
| 文件目录浏览 | `/files` | `POST /api/file-listing` | 选择本地文件夹 → 树形浏览 → 导出Excel |
| 日期计算 | `/dates` | `POST /api/date-calc` | 上传文件 → 选列 → 批量加/减天数 → 下载 |
| 单位转换 | `/units` | `POST /api/unit-convert` | 上传文件 → 选列 → 批量转换 → 下载 |

### 4.3 前端技术方案

- **路由**: 自定义 SPA Router（基于 `history.pushState` + Web Components）
- **组件**: 4 个可复用 Web Components
  - `<file-upload>` — 拖拽上传 + 文件列表
  - `<data-table>` — 可排序/可筛选的交互式表格
  - `<folder-tree>` — 递归文件夹树形组件
  - `<progress-bar>` — 异步任务进度条
- **样式**: CSS Custom Properties 实现主题系统
- **HTTP**: 原生 `fetch` API，FormData 上传文件
- **依赖**: 零外部 npm 包，浏览器原生 ES Modules

### 4.4 后端 API 设计

```python
# server.py — FastAPI 入口
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from cdm_engine import engine

app = FastAPI(title="CDM Toolkit API")

# ---- 工作簿合并 ----
@app.post("/api/merge")
async def api_merge(
    files: list[UploadFile] = File(...),
    create_toc: bool = Form(True),
    add_prefix: bool = Form(True),
) -> FileResponse:
    """多文件上传合并，返回合并后的 .xlsx"""
    result = engine.merge_workbooks_from_uploads(files, create_toc, add_prefix)
    return FileResponse(result, filename="merged.xlsx")

# ---- 工作表拆分 ----
@app.post("/api/split")
async def api_split(
    file: UploadFile = File(...),
    split_col: str = Form("A"),
    header_rows: int = Form(1),
) -> FileResponse:
    """按列值拆分工作表"""
    result = engine.split_sheet(file, split_col, header_rows)
    return FileResponse(result, filename="split.xlsx")

# ---- 特殊字符扫描 ----
@app.post("/api/scan-chars")
async def api_scan_chars(
    file: UploadFile = File(...),
    system: str = Form("Rave"),  # "Rave" or "Clinflash"
) -> dict:
    """扫描特殊字符，返回JSON结果 + 可选生成EDC报告"""
    findings = engine.scan_special_chars(file)
    return {"findings": findings, "total": len(findings)}

# ---- 文件目录浏览 ----
@app.post("/api/file-listing")
async def api_file_listing(
    path: str = Form(...),
    include_subfolders: bool = Form(True),
    file_pattern: str = Form("*.*"),
) -> dict:
    """浏览文件夹内容"""
    return engine.list_directory(path, include_subfolders, file_pattern)

# ---- 日期计算 ----
@app.post("/api/date-calc")
async def api_date_calc(
    file: UploadFile = File(...),
    col_letter: str = Form(...),
    days: int = Form(...),
    date_format: str = Form("YYYY-MM-DD"),
) -> FileResponse:
    """批量日期加/减天数"""
    result = engine.add_days_to_column_file(file, col_letter, days, date_format)
    return FileResponse(result, filename="dates_result.xlsx")

# ---- 单位转换 ----
@app.post("/api/unit-convert")
async def api_unit_convert(
    file: UploadFile = File(...),
    test_col: str = Form(...),
    result_col: str = Form(...),
    unit_col: str = Form(...),
    standard_unit: str = Form(...),
    test_filter: str = Form(""),
) -> FileResponse:
    """批量实验室单位转换"""
    result = engine.convert_lab_units_file(file, test_col, result_col, unit_col, standard_unit, test_filter)
    return FileResponse(result, filename="units_result.xlsx")
```

---

## 5. VBA 精简设计

### 5.1 模块列表（8个，精简33%）

| 模块 | 功能 | 状态 |
|------|------|------|
| modTextStandardise | 文本术语标准化 | 保持 |
| modHighlightDup | 重复值高亮 | 保持 |
| modBlankRows | 空白行管理（插空行/删空行） | 保持 |
| modTOC | TOC 目录生成 + SAS Derive 风格 | 保持 |
| modSheetManager | 工作表批量重命名/删除/列出 | 保持 |
| modTrackChanges | 修改痕迹跟踪（批注+颜色+报告） | 保持 |
| modDateCalc | 日期计算（加/减天数、日期差） | 🟡 双端共享 |
| modUnitConvert | 实验室单位转换 | 🟡 双端共享 |

### 5.2 移除的模块（4个，迁移到 Web）

| 原模块 | 迁移至 | 理由 |
|--------|--------|------|
| modWorkbookMerge | Web `/merge` | 拖拽上传+预览+进度条体验远优于VBA对话框 |
| modSheetSplit | Web `/split` | 大数据拆分VBA会卡死，Web异步处理+进度反馈 |
| modSpecialChars | Web `/scan` | 交互式结果表格+EDC报告是VBA做不到的 |
| modFileListing | Web `/files` | 树形组件天然优于Excel行内缩进 |

### 5.3 新增功能

- **DM_LaunchWeb**: 一键启动本地 Web 服务器并打开浏览器
  ```vb
  Public Sub DM_LaunchWeb()
      ' 启动 Python Web 服务并打开浏览器
      Shell "python -m cdm_tools serve", vbNormalFocus
  End Sub
  ```

### 5.4 VBA ↔ Web 互操作

| 方向 | 方式 | 实现 |
|------|------|------|
| VBA → Web | Alt+F8 → DM_LaunchWeb | Shell 启动 server.py → 自动打开浏览器 |
| 命令行 → Web | `cdm-tools serve` | FastAPI 启动 → 自动打开浏览器 |
| Web → Excel | 处理后下载 .xlsx | openpyxl 生成 → 浏览器下载 |

---

## 6. Python 引擎重构

### 6.1 从分散文件到统一包

**旧结构:**
```
cdm_toolkit.py       (14个函数)
cdm_toolkit_plus.py  (4个函数)
```

**新结构:**
```
cdm_engine/
├── __init__.py       # 公共 API 导出
├── engine.py         # 12个业务函数（合并重构）
├── conversions.py    # 常量数据（TERM_MAP、UNIT_CONVERSIONS）
└── sas_utils.py      # [新增] SAS 数据集读写
```

### 6.2 公共 API

```python
# cdm_engine/__init__.py
from cdm_engine.engine import (
    # 文本处理
    standardise_terms,
    # 重复高亮
    highlight_duplicates,
    # 空白行
    insert_blank_every_n,
    insert_blank_between_groups,
    delete_blank_rows,
    # 目录
    generate_toc,
    generate_sas_derive_toc,
    # 工作表管理
    batch_rename_sheets,
    list_sheet_names,
    delete_hidden_sheets,
    # 修改痕迹
    track_changes,
    changes_report,
    # 特殊字符
    scan_special_chars,
    # 文件目录
    generate_file_listing,
    generate_folder_tree,
    # 日期
    add_days_to_column,
    calc_date_diff,
    # 单位
    convert_lab_units,
    # 工作表拆分
    split_sheet_by_column,
    # 工作簿合并
    merge_workbooks,
    # 工具
    apply_header_style,
    auto_width,
)
```

### 6.3 新增 SAS 工具

```python
# cdm_engine/sas_utils.py
def read_sas(sas_path: str) -> pd.DataFrame:
    """读取 .sas7bdat 文件为 DataFrame"""

def sas_to_excel(sas_path: str, output_path: str) -> Workbook:
    """SAS 数据集 → Excel"""

def excel_to_sas(xlsx_path: str, output_path: str) -> None:
    """Excel → SAS 数据集（生成CSV + SAS导入脚本）"""
```

---

## 7. 项目目录结构

```
Tools_DM/
├── cdm_engine/                    # Python 核心引擎（统一逻辑层）
│   ├── __init__.py
│   ├── engine.py                  # 所有业务函数合并重构
│   ├── conversions.py             # 术语映射表 + 单位转换表
│   └── sas_utils.py               # [新增] SAS数据集读写
│
├── web/                           # Web 前端（纯静态文件）
│   ├── index.html                 # 仪表盘主页
│   ├── css/
│   │   └── app.css                # 全局样式 (CSS Custom Properties)
│   ├── js/
│   │   ├── app.js                 # SPA 路由 + 状态管理
│   │   ├── components/            # Web Components
│   │   │   ├── file-upload.js     # 拖拽上传组件
│   │   │   ├── data-table.js      # 交互式表格组件
│   │   │   ├── folder-tree.js     # 文件夹树形组件
│   │   │   └── progress-bar.js    # 进度条组件
│   │   └── pages/
│   │       ├── merge.js           # 工作簿合并页
│   │       ├── split.js           # 工作表拆分页
│   │       ├── scan.js            # 特殊字符扫描页
│   │       ├── files.js           # 文件目录浏览页
│   │       ├── dates.js           # 日期计算页
│   │       └── units.js           # 单位转换页
│   └── favicon.ico
│
├── server.py                      # FastAPI 入口: cdm-tools serve
│
├── vba/                           # VBA 工具箱（精简到8个模块）
│   ├── CDM_Toolkit.bas            # 主汇总模块
│   ├── CDM_Install.bas            # 一键安装 + DM_LaunchWeb
│   ├── CDM_Toolkit.xlsm           # 含宏工作簿
│   ├── README.md
│   └── modules/
│       ├── modTextStandardise.bas
│       ├── modHighlightDup.bas
│       ├── modBlankRows.bas
│       ├── modTOC.bas
│       ├── modSheetManager.bas
│       ├── modTrackChanges.bas
│       ├── modDateCalc.bas
│       └── modUnitConvert.bas
│
├── sas_macros/                    # [已有] 21个SAS宏（不变）
│   └── ...
│
├── tests/
│   ├── test_engine.py             # Python 引擎测试
│   └── test_vba_workbook.bas      # VBA 测试代码
│
├── reference/                     # [已有] 参考材料（不变）
├── DM_Toolbox.sas                 # [已有] SAS宏工具包主文件
├── pyproject.toml                 # pip install cdm-tools
└── README.md                      # 项目总览
```

---

## 8. 与旧版设计的对比

| 对比维度 | 旧设计（VBA全量平行） | 新设计（三层分离） |
|----------|----------------------|-------------------|
| VBA 模块数 | 12 | **8**（精简33%） |
| Python 代码 | 2个独立文件 | **1个统一包** `cdm_engine/` |
| Web 前端 | 无 | **6页SPA** + 4个Web Components |
| 部署方式 | VBA手动导入.bas | VBA导入 + `pip install cdm-tools` 一键启动Web |
| SAS 支持 | 仅SAS宏 | **新增** Python读写.sas7bdat |
| 批量管道 | 无 | **新增** 合并→清理→导出流水线 |
| 用户体验 | Excel内全操作 | Excel即时操作 + Web复杂处理 |
| 代码复用 | VBA和Python重复实现 | Python引擎统一，VBA翻译调用 |

---

## 9. 实现顺序（建议）

1. **Python 引擎重构**：合并 `cdm_toolkit.py` + `cdm_toolkit_plus.py` → `cdm_engine/`
2. **FastAPI 骨架**：`server.py` + 6个API端点
3. **Web 前端基础**：SPA路由 + 全局样式 + 仪表盘主页
4. **核心 Web 模块**：工作簿合并 + 工作表拆分（最高频需求）
5. **高级 Web 模块**：特殊字符扫描 + 文件目录浏览
6. **双端模块**：日期计算 + 单位转换（Web + VBA同步）
7. **VBA 精简**：移除4个模块 + 新增 DM_LaunchWeb
8. **SAS 工具**：`sas_utils.py` (read_sas, sas_to_excel, excel_to_sas)
9. **集成测试**：Python引擎 ↔ VBA ↔ Web 往返验证
10. **打包 & 文档**：`pyproject.toml` + `README.md` 更新

---

## 10. 风险与缓解

| 风险 | 缓解措施 |
|------|----------|
| 用户机器上没有 Python | 提供 `cdm-tools-standalone.exe` (PyInstaller打包)，内置Python运行时 |
| 企业环境限制 localhost 端口 | 默认8520端口；支持 `cdm-tools serve --port XXXX` 自定义 |
| Web Components 浏览器兼容性 | 目标 Chrome/Edge (Chromium)；企业环境几乎统一使用 |
| VBA 安全警告阻止宏运行 | 数字签名 + README 中详细说明启用步骤 |
| 中文编码问题 | 全部源文件 UTF-8；`.bas` 文件使用 UTF-8 BOM |
