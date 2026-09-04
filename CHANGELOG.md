# Changelog

本文件记录 OriginLab MCP Server 的版本变更。

格式参照 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

---

## [Unreleased]

### ✨ 新功能

- **本地状态面板**：新增 `originlab-mcp-ui` 命令，可在浏览器中查看 MCP Server、Origin 连接和常见客户端配置状态
- **客户端配置助手**：状态面板支持为 Antigravity / Gemini、Cursor、Codex、Trae、Claude Desktop 写入 `originlab` MCP 配置，并在覆盖已有文件前生成备份
- **一键安装入口**：新增 `scripts/install-and-open.ps1`，为 AI 助手和本机用户提供安装依赖并打开状态面板的单命令入口
- **系统阅读**：新增 `read_origin_session` 工具，只读当前 Origin 项目快照（工作表、图表、矩阵、Notes、活动对象、项目路径）
- **MCP Resources**：新增 `originlab://session`、`originlab://worksheets`、`originlab://graphs` 及工作表/图表模板资源，供客户端通过 `resources/read` 阅读会话
- **状态面板会话视图**：本地 UI 增加「阅读会话」，可列出当前工作表和图表
- **ToolAnnotations**：为全部 66 个 MCP tools 补充 `readOnlyHint` / `destructiveHint` / `idempotentHint` / `openWorldHint`，帮助客户端优先选择只读工具并规避破坏性操作
- **FakeOrigin 契约测试**：新增可复用的内存 OriginPro mock（对齐 Excel MCP mock-backend 思路），覆盖 import→plot→style→export 与项目生命周期工作流
- **Tool 选择引导**：统一 `change_plot_type` 的 When to use / When not to use，补齐 `save_project` / `close_origin` 的 `next_suggestions`，并交叉引用 `create_plot` / `add_plot_to_graph` / `change_plot_type`
- **LabTalk 安全门控**：`execute_labtalk` 对删除/重置/系统类命令要求 `confirm=true`，并在错误中返回 `suggested_alternatives`
- **错误恢复提示**：`error_response` / 常见 `ToolError` 增加 `suggested_alternatives`，引导 agent 改用标准 tools

### 🐛 Bug 修复

- **Origin 生命周期**：MCP Server 关闭时默认只 detach，不再直接退出用户的 Origin 实例
- **活动对象恢复**：当内存中的 active worksheet / graph 丢失时，尝试从当前 Origin 会话恢复活动对象
- **OriginPro API 兼容**：统一封装列标签和曲线获取逻辑，兼容没有 `WSheet.get_col()` / `GLayer.plot()` 的 originpro 版本

- **change_plot_type**：支持在原有图表窗口和图层中重建曲线类型，例如把点线图原位改为柱状图，不创建新的图表页

### 🧹 代码质量

- **mypy**：修复 `resources.py` / `ui.py` 类型错误，并为 Windows-only 的 `originpro` 添加 ignore_missing_imports
- **平台依赖**：`originpro` 标记为 `sys_platform == 'win32'`，Linux/macOS 可直接 `uv sync` 跑测试

### 🏗️ 工程

- **GitHub Actions CI**：新增 `.github/workflows/ci.yml`，在 Python 3.10–3.12 上跑 ruff / mypy / pytest（跳过 Windows-only 的 originpro 安装），并上传 coverage.xml

### 🧪 测试

- 新增生命周期、plot_list 兼容、原图类型替换和本地状态面板配置逻辑回归测试
- 新增 Origin 会话快照、MCP Resources 与状态面板阅读接口测试
- 新增 ToolAnnotations 覆盖回归测试（66 tools 全量校验）
- 新增 FakeOrigin COM 契约测试与共享 `tests/fakes` 测试双件（DummyMCP / attach_fake_origin）
- 新增 tool 描述与 `next_suggestions` 一致性回归测试
- 扩展 FakeOrigin 覆盖线性/非线性拟合、LabTalk 与多图层工作流；CI 产出 coverage.xml
- 新增 LabTalk confirm 门控与 `suggested_alternatives` 错误恢复测试

### 📝 文档

- 默认 README 改为英文，并保留中文文档为 README.zh.md
- README / README.en / README.zh 补充本地状态面板、Trae 配置和项目结构说明
- README 补充 `read_origin_session`、MCP Resources 与会话阅读说明，tool 总数 65 → 66

---

## [0.2.1] - 2026-04-08

### 🐛 Bug 修复

- **sort_worksheet**：修正列索引为 Origin 的 1-offset，添加边界检查，防止越界操作
- **set_legend**：使用 `layer_index` 参数替代硬编码 `layer -s 1`，正确支持多图层场景
- **get_plot**：拒绝负数索引（如 `-1`），防止 Python 负索引绕过验证静默通过
- **import_data_from_text**：使用 `max(所有行宽)` 计算列数，修复不等长行丢失尾列的问题
- **remove_plot_from_graph**：删除曲线前先验证 `plot_index` 有效性

### 🧹 代码质量

- **CSV 导出**：使用 `csv.writer` 替代手动逗号拼接，正确处理含逗号、引号、换行的单元格数据
- **异常链**：`analysis.py` 中的 `raise` 语句添加 `from e`，保留完整异常上下文（B904）
- **现代 import**：`Callable` / `Iterable` 从 `collections.abc` 导入（Python 3.9+ 推荐）
- **代码简化**：`contextlib.suppress` 替代 `try/except/pass` 模式
- **Lint 清零**：移除所有未使用 import、修正 import 排序、消除 `l` 变量名警告（ruff 全部通过）

### 🏗️ 工程清理

- **行尾统一**：13 个源文件从 CRLF/LF 混合统一为 LF
- **`.gitattributes`**：新增文件，强制 `eol=lf` 防止未来行尾混乱
- **`.gitignore`**：添加 `out/`、`.ruff_cache/`、`.mypy_cache/` 排除
- **重复测试**：删除 `test_phase3.py` 中与 `test_helpers.py` 重复的 `test_none_like_empty`

### 🧪 测试

- 新增 6 个回归测试：CSV 转义、文本导入列宽、排序索引、图例图层、负索引拒绝、get_plot 负数
- 测试总数：81 → 86（全部通过，0.12s）

### 📝 文档

- **README**：重构格式（HTML 表格 → Markdown），tool 总数 56 → 59，补齐 system 分类（1 → 4），添加 shields.io 徽章
- **CHANGELOG**：新增 v0.2.1 条目

---

## [0.2.0] - 2026-04-08

### 🏗️ 架构重构

- **依赖注入**：移除 `OriginManager` 单例模式，改为在 `server.py` 中创建唯一实例并显式注入所有 tool 模块
- **统一错误处理**：tool 函数覆盖 `@tool_error_handler` 装饰器，消除 ~47 处手写 `try/except` 样板代码
- **统一 resolve 范式**：plot.py 中 8 处手写资源获取逻辑替换为 `_resolve_xxx_name()` 辅助函数
- **Protocol 补齐**：`GraphLayerProtocol`、`GraphProtocol`、`OriginProProtocol` 新增 9 个方法定义

### ✨ 新功能

- **多图层支持**：customize.py 22 个工具新增 `layer_index` 参数（默认 0，向后兼容），支持操作双 Y 轴等多图层图表
- **`get_graph_layer()` 辅助函数**：统一图层获取逻辑，含越界检查和清晰错误提示
- **`LayerIndexError` 异常**：新增图层索引越界异常，提供修复建议
- **线型设置**：`set_plot_line_style` — 支持实线、虚线、点线、点划线等 8 种线型
- **线宽设置**：`set_plot_line_width` — 精确控制曲线线宽
- **误差棒样式**：`set_error_bar_style` — 设置线宽、端帽、颜色、方向
- **图例控制**：`set_legend` — 显示/隐藏、位置、字号

### 🔒 安全性

- **LabTalk 注入防护**：新增 `sanitize_labtalk_name()` 函数，通过正则表达式严格限制对象名，防止命令注入
- **命令长度限制**：LabTalk 命令限制 2000 字符上限
- 所有涉及对象名拼接的 LabTalk 命令强制应用防护

### 🧹 清理

- 移除废弃的 `validators.validate_column_index()` 和 `validate_column_indices()`（已被 `helpers.validate_column_indices` 异常模式替代）
- 消除所有 `gr[0]` 硬编码图层访问（26 处 → 0 处）
- 清理死导入和未使用符号

### 🧪 测试

- 新增 26 个测试用例（55 → 81，全部通过）
- 覆盖 LabTalk 注入防护、resolve 范式、图层越界、装饰器行为等场景
- 所有测试适配依赖注入模式

### ⚠️ 破坏性变更

- `OriginManager` 不再是单例模式，`OriginManager()` 每次返回新实例
- `register_xxx_tools()` 签名变更：`(mcp)` → `(mcp, manager)`
- 移除 `OriginManager.reset_for_testing()` 方法
- 移除 `validators.validate_column_index()` 和 `validate_column_indices()`

---

## [0.1.0] - 2025-xx-xx

### 初始版本

- 50+ 个 MCP tools，覆盖数据导入、工作表管理、图表创建与定制、数据分析（线性/非线性拟合）、导出
- 支持 Antigravity、Claude Desktop、Cursor 等 MCP 客户端
- LabTalk 命令逃生舱
