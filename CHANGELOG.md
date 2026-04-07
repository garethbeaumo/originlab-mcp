# Changelog

本文件记录 OriginLab MCP Server 的版本变更。

格式参照 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

---

## [0.2.0] - 2026-04-08

### 🏗️ 架构重构

- **依赖注入**：移除 `OriginManager` 单例模式，改为在 `server.py` 中创建唯一实例并显式注入所有 tool 模块
- **统一错误处理**：53/56 个 tool 函数覆盖 `@tool_error_handler` 装饰器，消除 ~47 处手写 `try/except` 样板代码
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

- 53 个 MCP tools，覆盖数据导入、工作表管理、图表创建与定制、数据分析（线性/非线性拟合）、导出
- 支持 Antigravity、Claude Desktop、Cursor 等 MCP 客户端
- LabTalk 命令逃生舱
