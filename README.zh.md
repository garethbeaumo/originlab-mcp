<p align="center">
  <h1 align="center">🔬 OriginLab MCP Server</h1>
  <p align="center">
    <strong>让 AI 成为你的 OriginLab 助手</strong>
  </p>
  <p align="center">
    通过 <a href="https://modelcontextprotocol.io">MCP 协议</a> 将 OriginLab 的数据分析与可视化能力无缝接入 Antigravity、Claude、Cursor 等 AI 客户端
  </p>
  <p align="center">
    <a href="https://github.com/garethbeaumo/originlab-mcp/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License"></a>
    <img src="https://img.shields.io/badge/python-3.10%2B-blue.svg" alt="Python">
    <img src="https://img.shields.io/badge/version-0.2.1-green.svg" alt="Version">
    <img src="https://img.shields.io/badge/platform-Windows-lightgrey.svg" alt="Platform">
    <img src="https://img.shields.io/badge/tools-66-orange.svg" alt="Tools">
  </p>
  <p align="center">
    <a href="#-快速开始">快速开始</a> · <a href="#-功能一览">功能一览</a> · <a href="#-使用示例">使用示例</a> · <a href="#-客户端配置">客户端配置</a>
  </p>
  <p align="center">
    简体中文 · <a href="README.md">English</a>
  </p>
</p>

> [!WARNING]
> **v0.2 早期版本** — 本项目仍处于早期开发阶段，功能和 API 可能随时变更。欢迎试用和反馈，但请勿用于生产环境。

---

## ✨ 什么是 OriginLab MCP Server？

OriginLab MCP Server 是一个连接 AI 与 OriginLab 的桥梁。它让你可以用**自然语言**完成数据导入、图表绘制、样式定制、数据分析和结果导出——无需手动操作 Origin 界面。

```
你：「把桌面上的 experiment.csv 导入 Origin，用第一列做 X、第二列做 Y，画散点图，
     做个高斯拟合，然后导出为 PNG。」

AI：好的，我来帮你完成。
    ✅ 已导入 experiment.csv → Sheet1（200 行 × 5 列）
    ✅ 已创建散点图 → Graph1
    ✅ 高斯拟合完成 → xc=2.35, w=0.82, A=156.3, R²=0.9987
    ✅ 已导出 → C:\Users\Desktop\Graph1.png
```

### 工作原理

```
AI 客户端 (Antigravity / Claude / Cursor)
       ↓ MCP 协议 (stdio)
OriginLab MCP Server (Python)
       ↓ originpro + COM
OriginLab
```

## 📋 前置条件

| 条件 | 要求 |
| :--- | :--- |
| 操作系统 | Windows |
| OriginLab | 2021 或更新版本，持有有效许可证 |
| Python | 3.10+ |

## 🚀 快速开始

给 AI 助手或本机用户的一键安装入口，在项目根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-and-open.ps1
```

这个脚本会在需要时安装 `uv`、执行 `uv sync`、启动本地状态面板，并自动打开 `http://127.0.0.1:8765/`。打开后可以测试 Origin 连接，也可以写入 MCP 客户端配置。

提供两种安装方式，任选其一：

<details open>
<summary><b>方式 A：使用 uv（推荐，更快）</b></summary>

**1. 安装 uv**

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

**2. 安装依赖**

```powershell
cd C:\path\to\originlab-mcp
uv sync
```

**3. 启动**

```powershell
uv run originlab-mcp
```

</details>

<details>
<summary><b>方式 B：使用 pip</b></summary>

**1. 创建虚拟环境（可选但推荐）**

```powershell
cd C:\path\to\originlab-mcp
python -m venv .venv
.venv\Scripts\activate
```

**2. 安装项目**

```powershell
pip install -e .
```

**3. 启动**

```powershell
originlab-mcp
```

</details>

Server 启动后通过 stdio 等待客户端连接，首次调用 tool 时自动连接本机 OriginLab。

### 可选：本地状态面板

如果只想检查 MCP Server 是否能启动、Origin 是否能连接、常见客户端配置是否存在，可以打开一个本地 UI：

```powershell
uv run originlab-mcp-ui
```

然后访问 `http://127.0.0.1:8765/`。这个页面可以启动/停止一个调试用 MCP Server 子进程、测试 Origin 连接、**阅读当前 Origin 会话**（工作表和图表），并为 Antigravity / Gemini、Cursor、Codex、Trae、Claude Desktop 一键写入 `originlab` MCP 配置；正常使用时仍建议让 AI 客户端按配置自动拉起 Server。

写入已有配置时会先创建 `.bak-时间戳` 备份。如果不希望启动时自动打开浏览器：

```powershell
$env:ORIGINLAB_MCP_UI_NO_BROWSER = "1"
uv run originlab-mcp-ui
```

## 🧰 功能一览

共提供 **66 个工具**，覆盖 OriginLab 的数据全流程：

### 📊 数据管理（14 个工具）

| 分类 | 工具 |
| :--- | :--- |
| 导入 | `import_csv` · `import_excel` · `import_data_from_text` |
| 查看 | `list_worksheets` · `get_worksheet_info` · `get_worksheet_data` · `get_cell_value` |
| 编辑 | `set_column_designations` · `set_column_labels` · `set_column_formula` |
| 管理 | `add_worksheet` · `sort_worksheet` · `clear_worksheet` · `delete_columns` |

### 📈 绘图（11 个工具）

| 分类 | 工具 |
| :--- | :--- |
| 创建 | `create_plot` · `create_double_y_plot` |
| 修改 | `add_plot_to_graph` · `remove_plot_from_graph` · `change_plot_data` · `change_plot_type` |
| 图层 | `add_graph_layer` · `group_plots` |
| 查看 | `list_graphs` · `list_graph_templates` · `get_graph_info` |

### 🎨 图表定制（25 个工具）

| 分类 | 工具 |
| :--- | :--- |
| 坐标轴 | `set_axis_range` · `set_axis_scale` · `set_axis_step` · `set_axis_title` |
| 线条 | `set_plot_line_style` · `set_plot_line_width` |
| 字体与刻度 | `set_graph_font` · `set_tick_style` |
| 颜色 | `set_plot_color` · `set_plot_colormap` · `set_plot_transparency` |
| 符号 | `set_plot_symbols` · `set_symbol_size` · `set_symbol_interior` |
| 分组递增 | `set_color_increment` · `set_symbol_increment` |
| 误差棒 | `set_error_bar_style` |
| 填充 | `set_fill_area` |
| 图例 | `set_legend` |
| 一键风格 | `apply_publication_style` |
| 标注 | `set_graph_title` · `add_text_label` · `add_line_to_graph` · `remove_graph_label` |

### 📐 数据分析（3 个工具）

`linear_fit` · `nonlinear_fit` · `list_fit_functions`

> 支持 Gauss、Lorentz、ExpDec1、Boltzmann 等常用拟合函数，可固定参数、设置初始值、带误差棒拟合。

### 💾 导出与项目（6 个工具）

`export_graph` · `export_all_graphs` · `export_worksheet_to_csv` · `save_project` · `open_project` · `new_project`

### 🔧 系统管理（5 个工具）

`get_origin_info` · `read_origin_session` · `release_origin` · `reconnect_origin` · `close_origin`

> `read_origin_session` 只读当前项目快照：工作簿/工作表、图表、矩阵、Notes、活动对象和项目路径。需要单元格预览时传入 `include_preview=true`。

### 📖 MCP Resources（系统阅读）

除 tools 外，客户端还可以通过 MCP `resources/read` 阅读 Origin 会话，不会修改项目：

| URI | 内容 |
| :--- | :--- |
| `originlab://session` | 当前项目全貌快照 |
| `originlab://worksheets` | 工作表列表 |
| `originlab://graphs` | 图表列表 |
| `originlab://worksheet/{book}/{sheet}` | 单个工作表结构与数据预览 |
| `originlab://graph/{name}` | 单个图表的图层与曲线 |

### ⚡ 高级（2 个工具）

`execute_labtalk` · `get_labtalk_variable`

其中 `execute_labtalk` 用于最后手段的脚本逃生舱，`get_labtalk_variable` 用于安全读取 LabTalk 变量值。

## 💬 使用示例

配置好客户端后，在 AI 对话中直接用自然语言操作：

| 你说的话 | AI 调用的 tool |
| :--- | :--- |
| 「当前 Origin 项目里有什么？」 | `read_origin_session` |
| 「把 data.csv 导入 Origin」 | `import_csv` |
| 「这个表有几列？列头是什么？」 | `get_worksheet_info` |
| 「第一列设 X，第二三列设 Y」 | `set_column_designations` |
| 「画个散点图」 | `create_plot` |
| 「再加一条第三列的曲线」 | `add_plot_to_graph` |
| 「X 轴标题改成 Time (s)，曲线改红色」 | `set_axis_title` + `set_plot_color` |
| 「一键套成论文图风格」 | `apply_publication_style` |
| 「把第二图层套成论文风格，线宽 3、符号 12」 | `apply_publication_style` |
| 「做个高斯拟合」 | `nonlinear_fit` |
| 「把 Y 轴改成对数刻度」 | `set_axis_scale` |
| 「导出 PNG 到桌面」 | `export_graph` |
| 「把项目里所有图都导出成 SVG」 | `export_all_graphs` |
| 「操作完了，释放 Origin 给我手动用」 | `release_origin` |

`apply_publication_style` 支持指定 `layer_index`，以及轴标题字号、刻度字号、图例字号、主刻度长度、次刻度数量、线宽、符号大小等常用论文图参数。

典型的完整工作流：

```
导入数据 → 查看结构 → 设置列角色 → 创建图表 → 定制外观 → 数据分析 → 导出结果
```

## 🔌 客户端配置

> [!NOTE]
> **你不需要手动启动 MCP Server。** 配置好后，AI 客户端会在需要时自动拉起 Server 子进程，通过 stdin/stdout 通信，整个过程对用户无感知。

路径请替换为你的实际项目路径。

也可以运行 `uv run originlab-mcp-ui`，在本地状态面板中选择客户端并自动写入配置。下面是手动配置示例。

### Antigravity (Gemini) — 推荐

在项目根目录创建 `.gemini/settings.json`：

**使用 uv：**

```json
{
  "mcpServers": {
    "originlab": {
      "command": "uv",
      "args": ["--directory", "C:\\path\\to\\originlab-mcp", "run", "originlab-mcp"]
    }
  }
}
```

**使用 pip：**

```json
{
  "mcpServers": {
    "originlab": {
      "command": "C:\\path\\to\\originlab-mcp\\.venv\\Scripts\\originlab-mcp.exe"
    }
  }
}
```

<details>
<summary><b>Claude Desktop</b></summary>

编辑 `%APPDATA%\Claude\claude_desktop_config.json`，内容格式同上（选择对应安装方式的配置）。

</details>

<details>
<summary><b>Cursor</b></summary>

在项目根目录创建 `.cursor/mcp.json`，内容格式同上。

</details>

<details>
<summary><b>Trae</b></summary>

在项目根目录创建 `.trae/mcp.json`，内容格式同上。

</details>

<details>
<summary><b>Codex (OpenAI)</b></summary>

在项目根目录创建 `.codex/config.json`，内容格式同上。

</details>

## 🧪 测试

```powershell
# uv
uv run python -m pytest tests/ -v

# pip（激活虚拟环境后）
pytest tests/ -v
```

基础测试不依赖 Origin 安装，可在任何环境运行。

## 📁 项目结构

```
originlab-mcp/
├── pyproject.toml                # 项目配置与依赖
├── CHANGELOG.md                  # 版本变更记录
├── scripts/
│   └── install-and-open.ps1      # 一键安装并启动 UI
├── src/originlab_mcp/
│   ├── server.py                 # MCP Server 入口 & 依赖注入
│   ├── ui.py                     # 本地状态面板
│   ├── origin_manager.py         # Origin COM 连接管理（线程安全）
│   ├── session.py                # 只读 Origin 会话快照
│   ├── resources.py              # MCP Resources（系统阅读）
│   ├── exceptions.py             # 自定义异常类
│   ├── types.py                  # Protocol 类型定义
│   ├── tools/
│   │   ├── data.py               # 📊 数据导入与工作表管理 (14)
│   │   ├── plot.py               # 📈 图表创建与管理 (11)
│   │   ├── customize.py          # 🎨 图表外观定制 (25)
│   │   ├── analysis.py           # 📐 线性/非线性拟合 (3)
│   │   ├── export.py             # 💾 导出与项目管理 (6)
│   │   ├── system.py             # 🔧 系统与连接管理 (4)
│   │   └── advanced.py           # ⚡ LabTalk 逃生舱 (2)
│   └── utils/
│       ├── constants.py          # 枚举、默认值、拟合函数定义
│       ├── helpers.py            # 图层/工作表/图表解析、错误处理装饰器
│       └── validators.py         # 参数校验与统一返回结构
└── tests/
    ├── test_helpers.py           # helpers 辅助函数测试
    ├── test_phase3.py            # LabTalk 防护与 resolve 范式测试
    ├── test_session.py           # 会话阅读与 MCP Resources 测试
    ├── test_tools.py             # tool 注册与集成测试
    └── test_ui.py                # 本地状态面板配置逻辑测试
```

## ❓ 常见问题

<details>
<summary><b>Origin 连接失败</b></summary>

请确认：

- OriginLab 2021+ 已安装在本机
- 持有有效许可证
- 当前用户有权限启动 Origin
- 没有其他程序占用 Origin COM 接口

</details>

<details>
<summary><b>MCP 客户端看不到 tools</b></summary>

1. 确认配置文件路径正确
2. 确认依赖已安装（`uv sync` 或 `pip install -e .`）
3. 重启 MCP 客户端

</details>

<details>
<summary><b>拟合结果不准确</b></summary>

1. 检查数据是否有异常值（可用 `get_worksheet_data` 预览）
2. 尝试通过 `initial_params` 提供更好的初始参数
3. 使用 `fixed_params` 固定已知参数
4. 确认选择了正确的拟合函数（`list_fit_functions` 可查看可用函数）

</details>

## 📜 许可证

[MIT](LICENSE) © 2025 garethbeaumo
