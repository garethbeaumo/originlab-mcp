# OriginLab MCP Server 开发计划

## 1. 项目目标

本项目的目标不是简单封装一组 OriginLab Python API，而是构建一个适合 AI 稳定调用的 MCP Server，使 AI 客户端能够通过自然语言完成以下任务：

- 导入和查看数据
- 理解工作表结构
- 设置列角色和标签
- 创建常见图表
- 修改图表外观
- 导出图表和保存项目
- 在标准工具不足时使用受控的高级能力

这份计划以 `AI 开发` 和 `AI 调用` 为核心，重点解决以下问题：

- AI 是否容易选对 tool
- AI 是否容易补齐参数
- tool 返回结果是否足以支持下一步调用
- 出错时 AI 是否能基于错误信息自我修复
- 工具集合是否支持自然语言任务的多步组合

## 2. 产品定义

### 2.1 核心定位

`OriginLab MCP Server` 是一个运行在 Windows 上的 Python MCP 服务端。它通过 `originpro` 控制本机已安装的 OriginLab，并将常见操作暴露为标准化 tools，供 Claude Desktop、Cursor 等 MCP 客户端调用。

### 2.2 非目标

以下内容不作为当前阶段目标：

- 替代 OriginLab 全量 GUI 功能
- 暴露底层全部 `originpro` API
- 支持任意复杂脚本编排
- 自动理解所有脏数据、歧义列结构和异常文件格式
- 在无 Origin 安装或无有效许可证的环境运行

### 2.3 成功标准

当 AI 接收到类似下面的自然语言时，应能稳定完成任务：

- “把桌面上的 `data.csv` 导入 Origin”
- “这个表有几列，分别是什么”
- “把第一列设成 X，第二列和第三列设成 Y”
- “用 A 做 X，B 做 Y，画散点图”
- “把 X 轴标题改成 `Time (s)`，第一条曲线改成红色”
- “导出成 PNG 到桌面”

成功不只意味着功能可用，也意味着：

- 参数语义清晰
- 返回结构统一
- 错误易于恢复
- tool 间可顺畅衔接

## 3. 技术路线

### 3.1 技术选择

| 项目        | 选择                | 理由                                                    |
| ----------- | ------------------- | ------------------------------------------------------- |
| AI 接入方式 | MCP Server          | 标准化协议，兼容 MCP 客户端                             |
| Origin 控制 | `originpro`         | OriginLab 官方 Python 包，通过 COM 控制 Origin          |
| MCP 框架    | FastMCP / `mcp` SDK | Python 侧工具定义清晰，便于声明式暴露 tools             |
| 传输方式    | stdio               | 通过标准输入/输出与 MCP 客户端通信，当前不支持 SSE/HTTP |
| 包管理      | `uv`                | Python 依赖管理和运行速度更稳定                         |

### 3.2 工作原理

```text
AI 客户端 -> MCP 协议 (stdio) -> MCP Server (Python) -> originpro -> OriginLab (COM)
```

当前阶段仅支持 stdio 传输模式，即 MCP 客户端通过启动子进程并通过 stdin/stdout 通信。不支持 SSE 或 HTTP 模式。如需远程访问，后续可扩展为 HTTP+SSE 方案。

### 3.3 前置条件

- Windows 操作系统
- OriginLab 2021 或更新版本
- 有效 OriginLab 许可证
- Python 3.10+
- `uv`

## 4. AI-First 设计原则

这是本项目最重要的设计约束。所有 tool 设计、参数命名、返回格式、异常语义都必须服从这些原则。

### 4.1 单一职责

每个 tool 只做一件明确的事，避免：

- 一个 tool 同时导入、设列、绘图
- 一个 tool 混合“查询状态”和“执行修改”
- 一个 tool 暴露过多互斥参数

这样做的目的，是降低 AI 选择错误 tool 的概率。

### 4.2 任务语义优先

参数名和返回值要面向任务语义，而不是底层 COM 细节。

优先：

- `sheet_name`
- `graph_name`
- `plot_type`
- `x_col`
- `y_cols`
- `output_path`

避免直接暴露难理解的内部对象或低层状态。

### 4.3 结构化返回

所有 tool 的返回值都必须是结构化数据，至少包含：

- `ok`: 是否成功
- `message`: 给 AI 的简洁结果说明
- `data`: 结构化结果主体

必要时增加：

- `resource`: 本次操作生成或引用的核心对象
- `warnings`: 非致命问题
- `next_suggestions`: 推荐下一步动作

### 4.4 错误可恢复

错误信息不能只说“失败”，必须告诉 AI：

- 失败原因
- 出错对象
- 可行修复方向

例如：

- “worksheet `RawData` 不存在，请先调用 `list_worksheets` 确认名称”
- “`x_col=5` 超出列范围，当前工作表只有 3 列”

### 4.5 默认值收敛

能合理推断的内容可以提供默认值，但默认规则必须稳定、可预测。

例如：

- 不传 `sheet_name` 时使用当前活动工作表
- 不传 `plot_type` 时默认使用 `line`
- 不传导出格式时按 `output_path` 扩展名推断

默认行为必须在文档中写清楚，避免 AI 误判。

### 4.6 对 AI 友好的命名

tool 名称要具备以下特征：

- 动词开头
- 语义直接
- 尽量避免缩写
- 同类能力命名模式一致

例如：

- `import_csv`
- `get_worksheet_info`
- `create_plot`
- `set_axis_title`
- `export_graph`

### 4.7 高风险能力隔离

任意 LabTalk 执行属于高风险兜底能力，必须：

- 独立命名
- 文档中明确“仅在标准 tool 不足时使用”
- 返回结果中标注这是逃生舱路径

## 5. AI 调用模型

本项目需要优先支持 AI 最常见的任务链，而不是孤立工具。

### 5.1 标准任务链

最常见的调用路径如下：

1. 导入数据
2. 查看工作表或列结构
3. 设置列角色
4. 创建图表
5. 定制图表
6. 导出图片或保存项目

对应自然语言示例：

```text
把 C:\data\exp.csv 导入 Origin
-> 看看这个表有几列
-> 第一列设成 X，第二列设成 Y
-> 画散点图
-> 把 X 轴标题改成 Time (s)
-> 导出成 PNG
```

### 5.2 AI 调用策略

为了让 AI 更稳，tool 设计要支持以下策略：

- 先查询，再修改
- 先验证对象存在，再引用对象
- 先返回资源名，再允许后续 tool 直接使用
- 在结果中提供下一步建议，减少 AI 盲猜

例如：

- `import_csv` 成功后返回工作表名
- `create_plot` 成功后返回图表名
- `list_worksheets` 返回所有候选工作表，减少名称歧义

### 5.3 错误恢复路径

系统至少要支持以下恢复模式：

- 对象不存在时，引导先列举对象
- 参数越界时，返回允许范围
- 图表类型不支持时，返回支持列表
- 文件路径错误时，返回可检查项

示例：

```json
{
  "ok": false,
  "message": "sheet_name 'Sheet9' not found",
  "error": {
    "type": "not_found",
    "target": "worksheet",
    "value": "Sheet9",
    "hint": "Call list_worksheets to inspect available worksheet names."
  }
}
```

## 6. Tool 体系设计

### 6.1 能力分组

保留原有 5 组标准能力和 1 组高级能力：

- 数据类 tools
- 绘图类 tools
- 图表定制类 tools
- 导出与项目管理类 tools
- 系统状态类 tools
- 高级逃生舱 tools

### 6.2 设计原则下的分组目标

| 分组   | 目标                                         | AI 调用特点                  |
| ------ | -------------------------------------------- | ---------------------------- |
| 数据类 | 让 AI 理解“有哪些表、有哪些列、列是什么角色” | 查询频率高，是大多数流程起点 |
| 绘图类 | 把工作表数据转成图表                         | 强依赖前序查询结果           |
| 定制类 | 修改图表外观和轴属性                         | 常作为绘图后的连续操作       |
| 导出类 | 形成最终成果物                               | 面向结果交付                 |
| 状态类 | 让 AI 判断服务和 Origin 当前状态             | 用于调试和连接确认           |
| 高级类 | 处理标准工具覆盖不到的场景                   | 默认不优先选择               |

## 7. Tool 统一规范

### 7.1 输入参数规范

所有 tool 应遵守以下约束：

- `sheet_name` 表示工作表逻辑名称
- `graph_name` 表示图表逻辑名称
- 列引用优先支持整数索引，必要时扩展支持字母列名
- 布尔参数只表示明确开关，不承载多义语义
- 可选参数用于补充，不应用于切换完全不同模式

推荐规则：

- 列索引统一为 `0-based`
- 文档明确是否支持 `"A"`, `"B"` 这类列名
- 如果同时支持整数和字母列，返回中写明已解析后的列

### 7.2 返回结构规范

推荐所有 tool 统一返回以下结构：

```json
{
  "ok": true,
  "message": "Plot created successfully.",
  "data": {},
  "resource": {},
  "warnings": [],
  "next_suggestions": []
}
```

字段说明：

- `ok`: 布尔值
- `message`: 简短结果摘要
- `data`: 与本次操作直接相关的结构化数据
- `resource`: 新建或更新后的核心对象标识
- `warnings`: 非致命问题列表
- `next_suggestions`: 建议 AI 下一步可调用的 tool

失败时推荐结构：

```json
{
  "ok": false,
  "message": "Column index out of range.",
  "error": {
    "type": "invalid_input",
    "target": "x_col",
    "value": 5,
    "hint": "Current worksheet has 3 columns. Use a value between 0 and 2."
  }
}
```

### 7.3 错误类别

建议统一使用以下错误类别：

- `not_found`
- `invalid_input`
- `unsupported`
- `conflict`
- `environment_error`
- `internal_error`

### 7.4 下一步建议规范

为了帮助 AI 自动衔接，关键 tool 应尽量返回 `next_suggestions`：

- `import_csv` -> `get_worksheet_info`, `set_column_designations`, `create_plot`
- `get_worksheet_info` -> `set_column_designations`, `create_plot`
- `create_plot` -> `set_axis_title`, `set_plot_color`, `add_plot_to_graph`, `export_graph`
- `add_plot_to_graph` -> `set_axis_title`, `set_plot_color`, `export_graph`
- `export_graph` -> `save_project`

### 7.5 活动对象策略

多数 tool 的 `sheet_name` 和 `graph_name` 参数都是可选的，省略时使用"当前活动对象"。为避免歧义，活动对象的管理必须遵守以下统一规则：

#### 规则

- **创建操作自动切换活动对象**：`import_csv`、`create_plot` 等创建操作完成后，新创建的工作表/图表自动成为活动对象
- **查询操作不改变活动对象**：`list_worksheets`、`get_worksheet_info`、`list_graphs` 等查询操作不会改变当前活动工作表或活动图表
- **修改操作不改变活动对象**：`set_column_designations`、`set_axis_title` 等修改操作只作用于指定对象，不改变活动对象状态
- **返回中始终标注活动对象**：每个 tool 的返回结构中，`resource` 字段应包含当前活动工作表和活动图表的名称，让 AI 始终可以追踪上下文

#### 示例

```text
用户：导入 data.csv -> import_csv 返回 sheet_name="Sheet1"，活动工作表变为 Sheet1
用户：画散点图   -> create_plot 使用活动工作表 Sheet1，返回 graph_name="Graph1"，活动图表变为 Graph1
用户：改成红色   -> set_plot_color 使用活动图表 Graph1，活动对象不变
```

### 7.6 Tool Description 编写规范

每个 tool 注册到 MCP Server 时的 description 是 AI 理解工具的核心依据。所有 tool 的 description 应遵守以下模板：

```text
[一句话功能概述]

何时使用：[正面规则，描述适用场景]
何时不用：[反面规则，描述不适用场景或应使用其他 tool 的情况]

默认行为：
- [可选参数的默认值说明]

示例：
- [一个典型调用示例，包含参数和预期结果]
```

好的 description 示例：

```text
创建图表。根据指定工作表的列数据创建一个新图表。

何时使用：需要从零开始创建一个新图表时使用。
何时不用：如果要在已有图表上追加曲线，请使用 add_plot_to_graph。

默认行为：
- sheet_name 省略时使用当前活动工作表
- plot_type 省略时默认为 line

示例：
- create_plot(sheet_name="Sheet1", plot_type="scatter", x_col=0, y_cols=1)
```

## 8. Tool 清单与设计要求

### 8.1 数据类 tools

目标：让 AI 能导入数据、理解工作表、设置列语义。

#### `import_csv`

- 功能：导入 CSV 文件到工作表
- 输入：`file_path`, `sheet_name?`
- 返回重点：
  - 实际导入的工作表名
  - 行数和列数
  - 是否创建了新表
- 下一步建议：
  - `get_worksheet_info`
  - `set_column_designations`
  - `create_plot`

#### `import_excel`

- 功能：导入 Excel 文件
- 输入：`file_path`, `sheet_name?`
- 补充要求：
  - 明确导入的是工作簿中的哪一页
  - 如果后续支持多 sheet，参数设计要避免歧义

#### `import_data_from_text`

- 功能：从 AI 直接提供的文本数据创建工作表
- 输入：`data`, `separator?`, `sheet_name?`, `has_header?`
- `data` 格式约定：
  - 默认接受以 `\n` 分行、以 `separator` 分列的纯文本（类似 CSV 格式）
  - `separator` 默认为 `,`，可设为 `\t`（TSV）、`|` 等
  - `has_header` 默认为 `true`，第一行作为列名
  - 示例：`"Name,Value\nA,1\nB,2\nC,3"`
- 适用场景：
  - 用户在聊天里直接贴数据
  - AI 自己生成的小样本数据
- 返回重点：
  - 实际创建的工作表名
  - 行数和列数
  - 解析使用的分隔符

#### `list_worksheets`

- 功能：列出当前项目中的工作表
- 返回重点：
  - 工作表名列表
  - 当前活动工作表

#### `get_worksheet_info`

- 功能：返回工作表结构
- 输入：`sheet_name`
- 返回重点：
  - 行数、列数
  - 列名
  - 列角色
  - 可选标签信息

#### `get_worksheet_data`

- 功能：返回工作表样本数据
- 输入：`sheet_name`, `max_rows?`
- 返回重点：
  - 截断后的表格数据
  - 是否发生截断

#### `set_column_designations`

- 功能：设置列角色
- 输入：`sheet_name`, `designations`
- 规则：
  - 支持 `X/Y/Z/YErr`
  - 返回变更后的列角色摘要

#### `set_column_labels`

- 功能：设置列长名、单位、注释
- 输入：`sheet_name`, `col`, `lname?`, `units?`, `comments?`
- 返回重点：
  - 更新前后差异

### 8.2 绘图类 tools

目标：让 AI 能把结构化数据稳定转换成图表。

#### `create_plot`

- 功能：创建新图表（支持单条或多条曲线）
- 输入：`sheet_name?`, `plot_type?`, `x_col`, `y_cols`, `z_col?`
- 参数说明：
  - `y_cols`：支持单个整数（单曲线）或整数数组（多曲线），统一用 `y_cols` 而非 `y_col`
  - `sheet_name` 省略时使用活动工作表
  - `plot_type` 省略时默认 `line`
- 返回重点：
  - 图表名
  - 图层信息
  - 每条曲线的列映射
  - 曲线数量
- 支持的 `plot_type`：
  - `line`
  - `scatter`
  - `line_symbol`
  - `column`
  - `area`
  - `auto`

> 注意：`create_plot` 始终创建一个**新图表**。如果需要在已有图表上追加曲线，请使用 `add_plot_to_graph`。

#### `add_plot_to_graph`

- 功能：在已有图表上追加一条或多条曲线
- 输入：`graph_name?`, `sheet_name?`, `x_col`, `y_cols`, `plot_type?`
- 参数说明：
  - `graph_name` 省略时使用活动图表
  - `sheet_name` 省略时使用活动工作表
- 返回重点：
  - 目标图表名
  - 新增曲线数量
  - 每条曲线的列映射
  - 当前图表总曲线数

#### `create_double_y_plot`

- 功能：创建双 Y 轴图
- 输入：`sheet_name`, `x_col`, `y1_col`, `y2_col`
- 返回重点：
  - 左右轴所绑定的列

#### `list_graphs`

- 功能：列出当前项目中的图表
- 返回重点：
  - 图表名列表
  - 当前活动图表

#### `list_graph_templates`

- 功能：列出支持的图表模板
- 返回重点：
  - 模板名
  - 推荐使用场景

### 8.3 图表定制类 tools

目标：让 AI 能对图表执行高频、低歧义的样式调整。

> 本组所有 tool 的 `graph_name` 参数均为可选，省略时默认使用当前活动图表。`plot_index` 默认为 `0`（第一条曲线）。

#### `set_axis_range`

- 输入：`graph_name`, `axis`, `min`, `max`
- 返回重点：
  - 修改后的范围

#### `set_axis_scale`

- 输入：`graph_name`, `axis`, `scale_type`
- 支持：
  - `linear`
  - `log`
  - `ln`

#### `set_axis_title`

- 输入：`graph_name`, `axis`, `title`
- 返回重点：
  - 轴标题更新结果

#### `set_plot_color`

- 输入：`graph_name`, `plot_index`, `color`
- 规则：
  - 优先支持十六进制颜色值
  - 必要时后续扩展颜色名映射

#### `set_plot_colormap`

- 输入：`graph_name`, `plot_index`, `colormap`

#### `set_plot_symbols`

- 输入：`graph_name`, `plot_index`, `shape_list`

### 8.4 导出与项目管理类 tools

目标：让 AI 能完成结果交付和项目持久化。

#### `export_graph`

- 输入：`graph_name`, `output_path`, `format?`, `width?`
- 支持格式：
  - `png`
  - `svg`
  - `pdf`
- 返回重点：
  - 输出文件路径
  - 实际导出格式

#### `save_project`

- 输入：`file_path?`
- 返回重点：
  - 实际保存路径

#### `open_project`

- 输入：`file_path`, `readonly?`
- 返回重点：
  - 是否成功打开
  - 打开的项目路径

#### `new_project`

- 功能：新建空白项目
- 返回重点：
  - 当前项目状态已重置

### 8.5 系统状态类 tools

目标：让 AI 能判断当前连接状态和运行环境。

#### `get_origin_info`

- 功能：返回 Origin 连接状态和基础环境信息
- 返回重点：
  - 是否已连接
  - Origin 安装路径
  - 用户文件路径
  - 当前活动项目摘要

### 8.6 高级逃生舱 tools

目标：处理标准 tool 覆盖不到的少量特殊场景。

#### `execute_labtalk`

- 功能：执行任意 LabTalk 命令
- 输入：`command`
- 设计要求：
  - 文档中明确标注高风险
  - 默认不作为首选调用路径
  - 返回执行结果和风险提示

> 注意：`execute_labtalk` 只作为兜底能力存在，标准任务优先通过显式 tool 完成。

## 9. 模块结构

```text
originlab/
├── PLAN.md
├── README.md
├── pyproject.toml
├── src/
│   └── originlab_mcp/
│       ├── __init__.py
│       ├── server.py
│       ├── origin_manager.py
│       ├── tools/
│       │   ├── __init__.py
│       │   ├── data.py
│       │   ├── plot.py
│       │   ├── customize.py
│       │   ├── export.py
│       │   ├── system.py
│       │   └── advanced.py
│       └── utils/
│           ├── __init__.py
│           ├── constants.py
│           └── validators.py
└── tests/
    ├── __init__.py
    └── test_tools.py
```

### 9.1 模块职责

#### `server.py`

- 创建 FastMCP 实例
- 注册全部 tools
- 处理服务启动与关闭生命周期

#### `origin_manager.py`

- 管理 Origin COM 连接
- 封装连接建立、断开、懒连接和状态检查
- 作为 tools 层访问 Origin 的统一入口
- **并发安全**：Origin COM 不支持多线程并发调用，所有 COM 操作必须通过串行锁（`threading.Lock`）或操作队列保证顺序执行
- **连接生命周期**：
  - 采用懒连接策略，首次调用 tool 时建立连接，不在 Server 启动时立即连接
  - 每次 COM 调用前检测连接有效性，断开时自动重连
  - MCP Server 关闭时（shutdown hook）主动释放 Origin COM 实例，避免残留进程
- **活动对象追踪**：维护当前活动工作表名和活动图表名，供 tool 读取默认值

#### `tools/*.py`

- 每个文件按能力域组织 tool
- 负责参数解析、调用协调、结果格式化
- 不应在 tool 函数中堆积复杂底层逻辑

各文件对应关系：

| 文件           | 对应能力组                     |
| -------------- | ------------------------------ |
| `data.py`      | 数据类 tools（§8.1）           |
| `plot.py`      | 绘图类 tools（§8.2）           |
| `customize.py` | 图表定制类 tools（§8.3）       |
| `export.py`    | 导出与项目管理类 tools（§8.4） |
| `system.py`    | 系统状态类 tools（§8.5）       |
| `advanced.py`  | 高级逃生舱 tools（§8.6）       |

#### `utils/constants.py`

- 维护图类型、错误类型、默认值、支持的枚举

#### `utils/validators.py`

- 做输入参数验证
- 输出统一错误格式

## 10. 分阶段开发计划

虽然本文档以 tool 设计为主，但实现仍按阶段推进。

### 阶段一：项目基础与连接能力

目标：MCP Server 能启动，Origin 能连接，AI 能判断服务可用性。

任务：

- 创建 `pyproject.toml`
- 建立包结构
- 实现 `OriginManager`
- 实现 `server.py`
- 注册 `get_origin_info`
- 统一基础返回结构和错误结构

验收：

```powershell
uv run originlab-mcp
```

应满足：

- MCP Server 能启动
- Origin 窗口能弹出
- AI 可调用 `get_origin_info`

### 阶段二：数据理解能力

目标：AI 能导入数据并理解工作表结构。

任务：

- `import_csv`
- `import_excel`
- `import_data_from_text`
- `list_worksheets`
- `get_worksheet_info`
- `get_worksheet_data`
- `set_column_designations`
- `set_column_labels`

验收场景：

- 导入文件后能返回实际工作表名
- AI 能查询列信息
- AI 能调整列角色

### 阶段三：绘图能力

目标：AI 能把工作表稳定转换为常见图表。

任务：

- `create_plot`
- `create_multi_plot`
- `create_double_y_plot`
- `list_graphs`
- `list_graph_templates`

验收场景：

- 根据指定列创建单图
- 在同一图上增加多条曲线
- 创建双 Y 轴图

### 阶段四：图表定制能力

目标：AI 能执行高频外观修改。

任务：

- `set_axis_range`
- `set_axis_scale`
- `set_axis_title`
- `set_plot_color`
- `set_plot_colormap`
- `set_plot_symbols`

验收场景：

- 修改坐标轴标题
- 修改轴范围和缩放
- 修改曲线颜色和符号

### 阶段五：导出与项目管理能力

目标：AI 能完成结果交付和项目存档。

任务：

- `export_graph`
- `save_project`
- `open_project`
- `new_project`

验收场景：

- 导出图表到指定路径
- 保存和重新打开项目

### 阶段六：高级能力与兜底路径

目标：补足标准 tool 覆盖不到的特殊场景。

任务：

- `execute_labtalk`
- 补充风险提示和调用约束

验收场景：

- AI 在标准 tool 无法满足需求时，能够退回到高级能力

## 11. 风险与边界

### 11.1 技术风险

- `originpro` 在不同 Origin 版本上的行为可能存在差异
- COM 连接可能受本机环境、权限、安装状态影响
- 某些图表模板和属性在不同版本中可用性不一致

### 11.2 AI 调用风险

- 用户提供的工作表名、图名可能与实际不一致
- 用户自然语言中的列语义可能存在歧义
- AI 可能过早调用高风险 tool

### 11.3 控制策略

- 优先提供查询类 tool，减少盲写
- 统一错误结构，明确提示修复路径
- 将高风险能力隔离到 `execute_labtalk`
- 先支持高频、低歧义任务，再逐步扩展

## 12. 验收用例

### 用例 1：导入并绘图

```text
用户：把 C:\data\experiment.csv 导入 Origin，用第一列做 X、第二列做 Y，画散点图
```

期望调用链：

- `import_csv`
- `set_column_designations`
- `create_plot`

### 用例 2：查看结构并补充标签

```text
用户：看看这个表有几列，分别是什么，再给第二列加单位 mV
```

期望调用链：

- `get_worksheet_info`
- `set_column_labels`

### 用例 3：图表定制

```text
用户：把 X 轴标题改成 Time (s)，Y 轴改成 Voltage (mV)，第一条曲线改成红色
```

期望调用链：

- `set_axis_title`
- `set_axis_title`
- `set_plot_color`

### 用例 4：导出成果

```text
用户：导出当前图表到桌面，格式为 PNG，宽度 1200
```

期望调用链：

- `export_graph`

### 用例 5：项目保存

```text
用户：把当前工作保存成 C:\data\analysis.opju
```

期望调用链：

- `save_project`

### 用例 6：工作表名错误（错误恢复）

```text
用户：把 Sheet9 的数据画成折线图
```

期望行为：

1. AI 调用 `create_plot(sheet_name="Sheet9", ...)`
2. 返回错误：`{"ok": false, "error": {"type": "not_found", "target": "worksheet", "value": "Sheet9", "hint": "Call list_worksheets to inspect available worksheet names."}}`
3. AI 自动调用 `list_worksheets`
4. AI 根据返回的工作表列表选择正确的名称，重新调用 `create_plot`

### 用例 7：文件路径不存在（错误恢复）

```text
用户：导入 C:\data\result.csv
```

期望行为：

1. AI 调用 `import_csv(file_path="C:\\data\\result.csv")`
2. 返回错误：`{"ok": false, "error": {"type": "invalid_input", "target": "file_path", "value": "C:\\data\\result.csv", "hint": "File does not exist. Please verify the file path and try again."}}`
3. AI 向用户确认正确的文件路径

### 用例 8：列索引越界（错误恢复）

```text
用户：用第 5 列做 Y 画图
```

期望行为：

1. AI 调用 `create_plot(x_col=0, y_cols=[4])`
2. 返回错误：`{"ok": false, "error": {"type": "invalid_input", "target": "y_cols", "value": [4], "hint": "Column index 4 out of range. Current worksheet has 3 columns (0-2). Call get_worksheet_info to inspect column details."}}`
3. AI 调用 `get_worksheet_info` 查看实际列结构
4. AI 向用户确认应使用哪一列

## 13. 对应 originpro API 参考

以下为主要参考 API，实际开发时仍需根据版本验证：

```python
wks = op.new_sheet(lname=name)
wks.from_file(path)
wks.from_df(df)
wks.to_df()
wks.rows
wks.cols
wks.cols_axis('XYY')
wks.set_label(col, val, type)
wks.from_list(col, data, lname=..., units=..., axis=...)

gr = op.new_graph(template='line')
gl = gr[0]
plot = gl.add_plot(wks, coly=1, colx=0, type='l')
gl.rescale()
gl.group()

gr = op.new_graph(template='doubley')
gr[0].add_plot(wks, 'B', 'A')
gr[1].add_plot(wks, 'C', 'A')

gl.set_xlim(0, 100)
gl.set_ylim(0, 50)
gl.xscale(1)
ax = gl.axis('x')
ax.title = 'Time (s)'
plot.color = '#ff5833'
plot.colormap = 'Candy'
plot.shapelist = [3, 2, 1]

graph.save_fig('output.png', width=800)
op.save('project.opju')
op.open(file='project.opju')
op.new()

op.lt_exec('window -a Graph1')
op.path('e')
op.path('u')
```

## 14. 使用方式

### 安装

```powershell
cd /path/to/originlab
uv sync
```

### 启动

```powershell
uv run originlab-mcp
```

### Claude Desktop 配置示例

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

## 15. 当前结论

本项目的优先级不再是“尽快罗列更多 Origin 功能”，而是：

1. 先把 tool 设计成 AI 容易理解和组合的形态
2. 让返回值和错误信息能驱动下一步调用
3. 让能力集合覆盖一条完整的数据分析任务链
4. 在标准化 tool 之外保留受控兜底能力

后续实现应始终围绕这个原则推进，避免退化为仅面向人工脚本调用的 API 包装层。
