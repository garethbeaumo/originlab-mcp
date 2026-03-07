<p align="center">
  <h1 align="center">🔬 OriginLab MCP Server</h1>
  <p align="center">
    <strong>让 AI 成为你的 OriginLab 助手</strong>
  </p>
  <p align="center">
    通过 <a href="https://modelcontextprotocol.io">MCP 协议</a> 将 OriginLab 的数据分析与可视化能力无缝接入 Antigravity 等 AI 客户端
  </p>
  <p align="center">
    <a href="#-快速开始">快速开始</a> · <a href="#-功能一览">功能一览</a> · <a href="#-使用示例">使用示例</a> · <a href="#-客户端配置">客户端配置</a>
  </p>
</p>

---

## ✨ 什么是 OriginLab MCP Server？

OriginLab MCP Server 是一个连接 AI 与 OriginLab 的桥梁。它让你可以用**自然语言**完成数据导入、图表绘制、样式定制和结果导出——无需手动操作 Origin 界面。

```
你：「把桌面上的 experiment.csv 导入 Origin，用第一列做 X、第二列做 Y，画散点图，导出为 PNG。」

AI：好的，我来帮你完成。
    ✅ 已导入 experiment.csv → Sheet1（200 行 × 5 列）
    ✅ 已创建散点图 → Graph1
    ✅ 已导出 → C:\Users\Desktop\Graph1.png
```

### 工作原理

```
AI 客户端 (Antigravity / Claude / Cursor)
       ↓ MCP 协议 (stdio)
OriginLab MCP Server (Python)
       ↓ originpro
OriginLab (COM)
```

## 📋 前置条件

| 条件      | 要求                            |
| --------- | ------------------------------- |
| 操作系统  | Windows                         |
| OriginLab | 2021 或更新版本，持有有效许可证 |
| Python    | 3.10+                           |

## 🚀 快速开始

提供两种安装方式，任选其一：

<details open>
<summary><b>方式 A：使用 uv（推荐，更快）</b></summary>

**1. 安装 uv**

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

**2. 安装依赖**

```powershell
cd C:\path\to\originlab
uv sync
```

**3. 启动**

```powershell
uv run originlab-mcp
```

</details>

<details open>
<summary><b>方式 B：使用 pip</b></summary>

**1. 创建虚拟环境（可选但推荐）**

```powershell
cd C:\path\to\originlab
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

## 🧰 功能一览

<table>
<tr>
<td width="140"><b>📊 数据管理</b></td>
<td>

`import_csv` · `import_excel` · `import_data_from_text`<br/>
`list_worksheets` · `get_worksheet_info` · `get_worksheet_data`<br/>
`set_column_designations` · `set_column_labels`

</td>
</tr>
<tr>
<td><b>📈 绘图</b></td>
<td>

`create_plot` · `add_plot_to_graph` · `create_double_y_plot`<br/>
`list_graphs` · `list_graph_templates`

</td>
</tr>
<tr>
<td><b>🎨 图表定制</b></td>
<td>

`set_axis_range` · `set_axis_scale` · `set_axis_title`<br/>
`set_plot_color` · `set_plot_colormap` · `set_plot_symbols`

</td>
</tr>
<tr>
<td><b>💾 导出与项目</b></td>
<td>

`export_graph` · `save_project` · `open_project` · `new_project`

</td>
</tr>
<tr>
<td><b>🔧 系统状态</b></td>
<td>

`get_origin_info`

</td>
</tr>
<tr>
<td><b>⚡ 高级</b></td>
<td>

`execute_labtalk` — 执行任意 LabTalk 命令（逃生舱）

</td>
</tr>
</table>

## 💬 使用示例

配置好客户端后，在 AI 对话中直接用自然语言操作：

| 你说的话                              | AI 调用的 tool                      |
| ------------------------------------- | ----------------------------------- |
| 「把 data.csv 导入 Origin」           | `import_csv`                        |
| 「这个表有几列？」                    | `get_worksheet_info`                |
| 「第一列设 X，第二三列设 Y」          | `set_column_designations`           |
| 「画个散点图」                        | `create_plot`                       |
| 「X 轴标题改成 Time (s)，曲线改红色」 | `set_axis_title` + `set_plot_color` |
| 「导出 PNG 到桌面」                   | `export_graph`                      |

典型的完整工作流：

```text
导入数据 → 查看结构 → 设置列角色 → 创建图表 → 定制外观 → 导出结果
```

## 🔌 客户端配置

> **你不需要手动启动 MCP Server。** 配置好后，AI 客户端会在需要 OriginLab 工具时自动在后台拉起子进程，通过 stdin/stdout 通信，整个过程对用户无感知。

路径请替换为你的实际项目路径。

### Antigravity (Gemini) — 推荐

在项目根目录创建 `.gemini/settings.json`：

**使用 uv：**

```json
{
  "mcpServers": {
    "originlab": {
      "command": "uv",
      "args": ["--directory", "C:\\path\\to\\originlab", "run", "originlab-mcp"]
    }
  }
}
```

**使用 pip：**

```json
{
  "mcpServers": {
    "originlab": {
      "command": "C:\\path\\to\\originlab\\.venv\\Scripts\\originlab-mcp.exe"
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
<summary><b>Codex (OpenAI)</b></summary>

在项目根目录创建 `.codex/config.json`，内容格式同上。

</details>

## 🧪 测试

```powershell
# uv
uv run pytest tests/ -v

# pip（激活虚拟环境后）
pytest tests/ -v
```

基础测试不依赖 Origin 安装，可在任何环境运行。

## 📁 项目结构

```
originlab/
├── pyproject.toml                # 项目配置与依赖
├── PLAN.md                       # 详细设计方案与 AI-First 原则
├── src/originlab_mcp/
│   ├── server.py                 # MCP Server 入口
│   ├── origin_manager.py         # Origin COM 连接管理（单例 + 线程安全）
│   ├── tools/
│   │   ├── data.py               # 📊 数据导入与工作表管理
│   │   ├── plot.py               # 📈 图表创建
│   │   ├── customize.py          # 🎨 图表外观定制
│   │   ├── export.py             # 💾 导出与项目管理
│   │   ├── system.py             # 🔧 系统状态
│   │   └── advanced.py           # ⚡ LabTalk 逃生舱
│   └── utils/
│       ├── constants.py          # 枚举、默认值
│       └── validators.py         # 参数校验与统一返回结构
└── tests/
    └── test_tools.py             # 单元测试
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

## 📜 许可证

[MIT](LICENSE) © 2026 garethbeaumo
