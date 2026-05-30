# CDM Toolkit v3.0

临床数据管理 (Clinical Data Management) 工具包 — 三层架构。

## 架构

| 层 | 技术 | 职责 |
|---|------|------|
| 表示层 | VBA 宏 + Web UI | Excel内即时操作 + 浏览器复杂处理 |
| 逻辑层 | Python `cdm_engine` | 统一业务逻辑（12个功能模块） |
| 数据层 | 本地文件系统 | Excel (.xlsx/.xls)、SAS (.sas7bdat)、CSV |

## 快速开始

### Web 工具

```bash
pip install -e .
cdm-tools serve          # → 浏览器自动打开 http://localhost:8520
```

### VBA 工具

1. 打开 Excel，按 Alt+F11
2. File → Import File → 选择 `vba/modules/` 下所有 `.bas` 文件
3. 再导入 `vba/CDM_Toolkit.bas` 和 `vba/CDM_Install.bas`
4. 从 VBA 启动 Web: Alt+F8 → DM_LaunchWeb

### Python 引擎

```python
from cdm_engine import standardise_terms, generate_toc, merge_workbooks

wb = load_workbook("your_file.xlsx")
standardise_terms(wb["AE"])
generate_toc(wb)
wb.save("output.xlsx")
```

## 功能清单

| 功能 | VBA | Web | Python |
|------|:---:|:---:|:------:|
| 文本术语标准化 | ✅ | — | ✅ |
| 重复值高亮 | ✅ | — | ✅ |
| 空白行管理 | ✅ | — | ✅ |
| TOC 目录生成 | ✅ | — | ✅ |
| 工作表管理 | ✅ | — | ✅ |
| 修改痕迹跟踪 | ✅ | — | ✅ |
| 工作簿合并 | — | ✅ | ✅ |
| 工作表拆分 | — | ✅ | ✅ |
| 特殊字符扫描 | — | ✅ | ✅ |
| 文件目录浏览 | — | ✅ | ✅ |
| 日期计算 | ✅ | ✅ | ✅ |
| 单位转换 | ✅ | ✅ | ✅ |
| SAS 读写 | — | — | ✅ |

## 项目结构

```
Tools_DM/
├── cdm_engine/          # Python 核心引擎
│   ├── __init__.py      #   公共 API
│   ├── engine.py        #   27个业务函数
│   ├── conversions.py   #   常量（术语表、单位转换表、样式）
│   └── sas_utils.py     #   SAS 数据集读写
├── web/                 # Web 前端（纯静态文件）
│   ├── index.html       #   仪表盘
│   ├── css/app.css      #   全局样式
│   └── js/              #   SPA路由 + 组件 + 页面
├── server.py            #   FastAPI 入口
├── cdm_tools/cli.py     #   命令行入口 (cdm-tools)
├── pyproject.toml       #   打包配置
├── vba/                 #   VBA 模块 (8个)
├── sas_macros/          #   SAS 宏 (26个)
├── tests/               #   测试套件
└── reference/           #   参考材料
```

## 系统要求

- Python 3.10+
- Excel 2010+ (for VBA)
- Chrome/Edge (for Web UI)
- Optional: pyreadstat (for SAS .sas7bdat support)
