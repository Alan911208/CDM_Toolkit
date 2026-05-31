# CDM Toolkit v3.0

临床数据管理 (Clinical Data Management) 工具包 — Python + SAS + Web 架构。

## 架构

| 层 | 技术 | 职责 |
|---|------|------|
| 前端 | JavaScript SPA + Web Components | 浏览器内交互、文件上传、结果预览 |
| 后端 | Python FastAPI + `cdm_engine` | API 服务、17 个业务模块、SAS 执行 |
| 数据 | 本地文件系统 | Excel (.xlsx/.xls)、SAS (.sas7bdat)、CSV |

## 快速开始

### 启动 Web 工具

```bash
pip install -e .
cdm-tools serve          # → 浏览器自动打开 http://localhost:8520
```

或直接运行：

```bash
python server.py
```

### Python 引擎

```python
from cdm_engine import standardise_terms, generate_toc, calc_cockcroft_gault

wb = load_workbook("your_file.xlsx")
standardise_terms(wb["AE"])
generate_toc(wb)
wb.save("output.xlsx")

# 医学计算
crcl = calc_cockcroft_gault(age=65, weight_kg=70, scr_mg_dl=1.2, is_female=False)
```

## 功能清单

| 功能 | Web | Python Engine |
|------|:---:|:-------------:|
| 工作簿合并 | ✅ | ✅ |
| 工作表拆分 | ✅ | ✅ |
| 特殊字符扫描 | ✅ | ✅ |
| 文件目录浏览 | ✅ | ✅ |
| 日期计算 | ✅ | ✅ |
| 单位转换 | ✅ | ✅ |
| 数据清理 (文本/重复/空白/删除线) | ✅ | ✅ |
| 工作表管理 (创建/删除/重命名/列出) | ✅ | ✅ |
| TOC 目录生成 (标准/SAS Derive) | ✅ | ✅ |
| 修改痕迹跟踪 | ✅ | ✅ |
| 医学计算器 (CrCl/eGFR) | ✅ | ✅ |
| SAS 程序运行器 | ✅ | — |
| SAS 数据读写 (.sas7bdat) | — | ✅ |

## 项目结构

```
Tools_DM/
├── cdm_engine/          # Python 核心引擎
│   ├── __init__.py      #   公共 API
│   ├── engine.py        #   30+ 业务函数
│   ├── conversions.py   #   常量（术语表、单位转换表、样式）
│   └── sas_utils.py     #   SAS 数据集读写
├── web/                 # Web 前端
│   ├── index.html       #   仪表盘 (SPA)
│   ├── css/app.css      #   全局样式
│   └── js/              #   路由、组件、11个功能页面
├── server.py            #   FastAPI 入口 (15个 API 端点)
├── cdm_tools/cli.py     #   命令行入口 (cdm-tools)
├── pyproject.toml       #   打包配置
├── sas_macros/          #   SAS 宏 (26个，含测试)
├── tests/               #   测试套件
└── reference/           #   SAS 参考材料
```

## 系统要求

- Python 3.10+
- Chrome/Edge/Firefox (Web UI)
- Optional: SAS 9.4 (SAS 程序运行器)
- Optional: pyreadstat (SAS .sas7bdat 支持)
