# Changelog

本文件记录 OriginLab MCP Server 的版本变更。

格式参照 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

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
