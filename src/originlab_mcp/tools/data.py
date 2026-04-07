"""数据类 tools。

让 AI 能导入数据、理解工作表结构、设置列语义。

包含:
    import_csv: 导入 CSV 文件到工作表
    import_excel: 导入 Excel 文件到工作表
    import_data_from_text: 从文本数据创建工作表
    list_worksheets: 列出所有工作表
    get_worksheet_info: 返回工作表结构信息
    get_worksheet_data: 返回工作表样本数据
    set_column_designations: 设置列角色
    set_column_labels: 设置列标签
"""

from __future__ import annotations

from typing import Any

from originlab_mcp.exceptions import (
    ColumnIndexError,
    ToolError,
)
from originlab_mcp.utils.constants import (
    DEFAULT_HAS_HEADER,
    DEFAULT_MAX_PREVIEW_ROWS,
    DEFAULT_SEPARATOR,
    ColumnDesignation,
)
from originlab_mcp.utils.helpers import (
    find_worksheet as _find_worksheet,
)
from originlab_mcp.utils.helpers import (
    resolve_worksheet_name as _resolve_worksheet_name,
)
from originlab_mcp.utils.helpers import (
    tool_error_handler,
)
from originlab_mcp.utils.validators import (
    error_response,
    success_response,
    validate_file_path,
)

# 注: _resolve_worksheet_name 和 _find_worksheet 从 utils.helpers 导入


def _get_column_display_name(col: Any, index: int) -> str:
    """返回用于预览输出的列名。"""
    long_name = ""
    if hasattr(col, "get_label"):
        try:
            long_name = col.get_label("L")
        except Exception:
            long_name = ""

    if long_name:
        return str(long_name)

    short_name = getattr(col, "name", "")
    if short_name:
        return str(short_name)

    return f"Col{index + 1}"


def _make_unique_column_name(name: str, used: set[str]) -> str:
    """为重复列名追加稳定后缀，避免 records 键冲突。"""
    if name not in used:
        used.add(name)
        return name

    suffix = 2
    while f"{name}_{suffix}" in used:
        suffix += 1
    unique_name = f"{name}_{suffix}"
    used.add(unique_name)
    return unique_name


def _normalize_designations(spec: str) -> str:
    """将用户输入的列角色字符串归一化为 Origin 接受的单字符格式。"""
    normalized: list[str] = []
    valid_chars = {
        ColumnDesignation.X.value,
        ColumnDesignation.Y.value,
        ColumnDesignation.Z.value,
        ColumnDesignation.Y_ERROR.value,
        ColumnDesignation.LABEL.value,
        ColumnDesignation.DISREGARD.value,
    }

    i = 0
    while i < len(spec):
        if spec[i].isspace():
            i += 1
            continue

        if spec[i:i + 4].lower() == "yerr":
            normalized.append(ColumnDesignation.Y_ERROR.value)
            i += 4
            continue

        token = spec[i].upper()
        if token in valid_chars:
            normalized.append(token)
            i += 1
            continue

        raise ToolError(
            f"Unsupported designation '{spec[i:]}'",
            error_type="invalid_input",
            target="designations",
            value=spec,
            hint="请提供列角色字符串。支持: X, Y, Z, E(误差), L, N；兼容旧写法 YErr。",
        )

    return "".join(normalized)


def register_data_tools(mcp: Any, manager: Any) -> None:
    """注册数据类 tools 到 MCP Server。

    Args:
        mcp: FastMCP 实例。
        manager: OriginManager 实例（依赖注入）。
    """

    # =================================================================
    # import_csv
    # =================================================================

    @mcp.tool()
    @tool_error_handler("导入CSV", "请检查文件格式是否为有效的 CSV。")
    def import_csv(file_path: str, sheet_name: str | None = None) -> dict:
        """导入 CSV 文件到工作表。

        何时使用：需要将本地 CSV 文件导入 Origin 工作表时使用。
        何时不用：导入 Excel 文件请用 import_excel；直接传入文本数据请用 import_data_from_text。

        默认行为：
        - sheet_name 省略时创建新工作表

        示例：
        - import_csv(file_path="C:\\data\\exp.csv")
        - import_csv(file_path="C:\\data\\exp.csv", sheet_name="RawData")
        """
        err = validate_file_path(file_path)
        if err:
            return error_response(
                message=err,
                error_type="invalid_input",
                target="file_path",
                value=file_path,
                hint="文件不存在，请检查文件路径。",
            )

        def _import(op: Any) -> dict[str, Any]:
            created_new = False
            if sheet_name:
                wks = op.find_sheet("w", sheet_name)
                if wks is None:
                    wks = op.new_sheet(lname=sheet_name)
                    created_new = True
            else:
                wks = op.new_sheet()
                created_new = True

            wks.from_file(file_path)
            book_name = wks.get_book().name
            actual_name = wks.name
            full_name = f"[{book_name}]{actual_name}"
            manager.active_worksheet = full_name

            return {
                "sheet_name": full_name,
                "rows": wks.rows,
                "cols": wks.cols,
                "created_new_sheet": created_new,
                "source_file": file_path,
            }

        result = manager.execute(_import)

        return success_response(
            message=(
                f"CSV 已导入到工作表 '{result['sheet_name']}'，"
                f"共 {result['rows']} 行 {result['cols']} 列。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "get_worksheet_info",
                "set_column_designations",
                "create_plot",
            ],
        )

    # =================================================================
    # import_excel
    # =================================================================

    @mcp.tool()
    @tool_error_handler("导入Excel", "请检查文件格式是否为有效的 Excel 文件。")
    def import_excel(file_path: str, sheet_name: str | None = None) -> dict:
        """导入 Excel 文件到工作表。

        何时使用：需要将本地 Excel (.xlsx/.xls) 文件导入 Origin 工作表时使用。
        何时不用：导入 CSV 请用 import_csv。

        默认行为：
        - sheet_name 省略时创建新工作表
        - 默认导入 Excel 文件的第一个工作页

        示例：
        - import_excel(file_path="C:\\data\\results.xlsx")
        """
        err = validate_file_path(file_path)
        if err:
            return error_response(
                message=err,
                error_type="invalid_input",
                target="file_path",
                value=file_path,
                hint="文件不存在，请检查文件路径。",
            )

        def _import(op: Any) -> dict[str, Any]:
            if sheet_name:
                wks = op.find_sheet("w", sheet_name)
                if wks is None:
                    wks = op.new_sheet(lname=sheet_name)
            else:
                wks = op.new_sheet()

            wks.from_file(file_path)
            book_name = wks.get_book().name
            actual_name = wks.name
            full_name = f"[{book_name}]{actual_name}"
            manager.active_worksheet = full_name

            return {
                "sheet_name": full_name,
                "rows": wks.rows,
                "cols": wks.cols,
                "source_file": file_path,
            }

        result = manager.execute(_import)

        return success_response(
            message=(
                f"Excel 已导入到工作表 '{result['sheet_name']}'，"
                f"共 {result['rows']} 行 {result['cols']} 列。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "get_worksheet_info",
                "set_column_designations",
                "create_plot",
            ],
        )

    # =================================================================
    # import_data_from_text
    # =================================================================

    @mcp.tool()
    @tool_error_handler("导入文本数据", "请检查数据格式。数据应以换行符分行，以指定的分隔符分列。")
    def import_data_from_text(
        data: str,
        separator: str = DEFAULT_SEPARATOR,
        sheet_name: str | None = None,
        has_header: bool = DEFAULT_HAS_HEADER,
    ) -> dict:
        """从文本数据创建工作表。适用于用户直接贴数据或 AI 生成的小样本数据。

        何时使用：用户在聊天中直接提供数据文本，或 AI 需要快速创建示例数据时使用。
        何时不用：数据保存在本地文件中时，请用 import_csv 或 import_excel。

        默认行为：
        - separator 默认为 ","
        - has_header 默认为 true，第一行作为列名
        - sheet_name 省略时自动命名

        示例：
        - import_data_from_text(data="Name,Value\\nA,1\\nB,2\\nC,3")
        - import_data_from_text(data="X\\tY\\n1\\t10\\n2\\t20", separator="\\t")
        """
        if not data or not data.strip():
            return error_response(
                message="data 不能为空",
                error_type="invalid_input",
                target="data",
                hint="请提供以换行符分行、以分隔符分列的文本数据。",
            )

        lines = data.strip().split("\n")
        if len(lines) < 1:
            return error_response(
                message="data 至少需要包含一行",
                error_type="invalid_input",
                target="data",
                hint="请提供至少一行数据。",
            )

        # 解析表头和数据行
        header = None
        start_row = 0
        if has_header and len(lines) > 1:
            header = [h.strip() for h in lines[0].split(separator)]
            start_row = 1

        rows_data = []
        for line in lines[start_row:]:
            row = [cell.strip() for cell in line.split(separator)]
            rows_data.append(row)

        num_cols = max(
            (len(row) for row in rows_data),
            default=0,
        )
        if header:
            num_cols = max(num_cols, len(header))

        def _import(op: Any) -> dict[str, Any]:
            wks = (
                op.new_sheet(lname=sheet_name) if sheet_name
                else op.new_sheet()
            )

            for col_idx in range(num_cols):
                col_data = []
                for row in rows_data:
                    val = row[col_idx] if col_idx < len(row) else ""
                    try:
                        val = float(val)
                        if val == int(val):
                            val = int(val)
                    except (ValueError, TypeError):
                        pass
                    col_data.append(val)

                lname = (
                    header[col_idx]
                    if header and col_idx < len(header)
                    else None
                )
                wks.from_list(col_idx, col_data, lname=lname)

            book_name = wks.get_book().name
            actual_name = wks.name
            full_name = f"[{book_name}]{actual_name}"
            manager.active_worksheet = full_name
            return {
                "sheet_name": full_name,
                "rows": len(rows_data),
                "cols": num_cols,
                "separator_used": separator,
                "has_header": has_header,
            }

        result = manager.execute(_import)

        return success_response(
            message=(
                f"文本数据已导入到工作表 '{result['sheet_name']}'，"
                f"共 {result['rows']} 行 {result['cols']} 列。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "get_worksheet_info",
                "set_column_designations",
                "create_plot",
            ],
        )

    # =================================================================
    # list_worksheets
    # =================================================================

    @mcp.tool()
    @tool_error_handler("列出工作表", "请确认 Origin 已连接。可调用 get_origin_info 检查状态。")
    def list_worksheets() -> dict:
        """列出当前项目中的所有工作表。

        何时使用：需要查看有哪些可用的工作表，或确认工作表名称时使用。
        何时不用：已经知道工作表名称时无需调用。

        示例：
        - list_worksheets()
        """
        def _list(op: Any) -> list[dict[str, Any]]:
            worksheets = []
            for book in op.pages("Book"):
                for sheet in book:
                    worksheets.append({
                        "book_name": book.name,
                        "sheet_name": sheet.name,
                        "full_name": f"[{book.name}]{sheet.name}",
                        "rows": sheet.rows,
                        "cols": sheet.cols,
                    })
            return worksheets

        sheets = manager.execute(_list)

        return success_response(
            message=f"当前项目共有 {len(sheets)} 个工作表。",
            data={"worksheets": sheets, "count": len(sheets)},
            resource=manager.get_resource_context(),
            next_suggestions=["get_worksheet_info", "import_csv"],
        )

    # =================================================================
    # get_worksheet_info
    # =================================================================

    @mcp.tool()
    @tool_error_handler("获取工作表信息", "请确认工作表名称正确。调用 list_worksheets 查看可用工作表。")
    def get_worksheet_info(sheet_name: str | None = None) -> dict:
        """返回工作表的结构信息（列名、列角色、行列数等）。

        何时使用：需要了解工作表有多少列、列名是什么、列角色如何设置时使用。
        何时不用：只需看工作表数据内容时，请用 get_worksheet_data。

        默认行为：
        - sheet_name 省略时使用当前活动工作表

        示例：
        - get_worksheet_info()
        - get_worksheet_info(sheet_name="Sheet1")
        """
        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _info(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)

            columns = []
            for i in range(wks.cols):
                col = wks.get_col(i)
                col_info: dict[str, Any] = {
                    "index": i,
                    "name": (
                        col.name if hasattr(col, "name") else f"Col{i+1}"
                    ),
                    "long_name": (
                        col.get_label("L")
                        if hasattr(col, "get_label") else ""
                    ),
                    "units": (
                        col.get_label("U")
                        if hasattr(col, "get_label") else ""
                    ),
                    "comments": (
                        col.get_label("C")
                        if hasattr(col, "get_label") else ""
                    ),
                }
                columns.append(col_info)

            # 获取列角色
            try:
                designations = wks.get_labels("D")
                for i, d in enumerate(designations):
                    if i < len(columns):
                        columns[i]["designation"] = d
            except Exception:
                pass

            return {
                "sheet_name": target_name,
                "rows": wks.rows,
                "cols": wks.cols,
                "columns": columns,
            }

        result = manager.execute(_info)

        return success_response(
            message=(
                f"工作表 '{target_name}' "
                f"共 {result['rows']} 行 {result['cols']} 列。"
            ),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "set_column_designations",
                "get_worksheet_data",
                "create_plot",
            ],
        )

    # =================================================================
    # get_worksheet_data
    # =================================================================

    @mcp.tool()
    @tool_error_handler("获取工作表数据", "请确认工作表名称正确。")
    def get_worksheet_data(
        sheet_name: str | None = None,
        max_rows: int = DEFAULT_MAX_PREVIEW_ROWS,
    ) -> dict:
        """返回工作表的样本数据（用于预览）。

        何时使用：需要查看工作表中的实际数据内容时使用。
        何时不用：只需了解列结构和角色，请用 get_worksheet_info。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - max_rows 默认 20，最多返回指定行数的数据

        示例：
        - get_worksheet_data()
        - get_worksheet_data(sheet_name="Sheet1", max_rows=10)
        """
        if max_rows < 0:
            return error_response(
                message="max_rows 不能小于 0",
                error_type="invalid_input",
                target="max_rows",
                value=max_rows,
                hint="请传入 0 或正整数。",
            )

        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _data(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            total_cols = wks.cols
            used_names: set[str] = set()
            col_names: list[str] = []
            col_values: list[list[Any]] = []
            for ci in range(total_cols):
                col_obj = wks.get_col(ci)
                base_name = _get_column_display_name(col_obj, ci)
                col_name = _make_unique_column_name(base_name, used_names)
                col_names.append(col_name)
                col_values.append(list(wks.to_list(ci)))

            total_rows = max((len(values) for values in col_values), default=0)
            truncated = total_rows > max_rows
            rows_to_return = min(total_rows, max_rows)

            data_records = []
            for ri in range(rows_to_return):
                record = {
                    col_names[ci]: (
                        col_values[ci][ri] if ri < len(col_values[ci]) else None
                    )
                    for ci in range(total_cols)
                }
                data_records.append(record)

            return {
                "sheet_name": target_name,
                "total_rows": total_rows,
                "returned_rows": rows_to_return,
                "truncated": truncated,
                "columns": col_names,
                "data": data_records,
            }

        result = manager.execute(_data)

        msg = f"返回 {result['returned_rows']} 行数据"
        if result["truncated"]:
            msg += f"（共 {result['total_rows']} 行，已截断）"
        msg += "。"

        return success_response(
            message=msg,
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["set_column_designations", "create_plot"],
        )

    # =================================================================
    # set_column_designations
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置列角色", "请检查角色字符串长度是否与列数匹配，且仅使用 X/Y/Z/E/L/N 或兼容写法 YErr。")
    def set_column_designations(
        designations: str,
        sheet_name: str | None = None,
    ) -> dict:
        """设置工作表的列角色。

        何时使用：需要指定哪些列是 X、Y、Z 或误差列时使用。
        何时不用：只需查看列角色时，请用 get_worksheet_info。

        默认行为：
        - sheet_name 省略时使用当前活动工作表

        参数说明：
        - designations: 列角色字符串，如 "XYY" 表示第一列为 X，后两列为 Y

        示例：
        - set_column_designations(designations="XYY")
        - set_column_designations(designations="XYYY", sheet_name="Sheet1")
        """
        if not designations:
            return error_response(
                message="designations 不能为空",
                error_type="invalid_input",
                target="designations",
                hint="请提供列角色字符串。支持: X, Y, Z, E(误差), L, N；兼容旧写法 YErr。",
            )

        target_name = _resolve_worksheet_name(sheet_name, manager)
        normalized_designations = _normalize_designations(designations)

        def _set(op: Any) -> list[dict[str, Any]]:
            wks = _find_worksheet(op, target_name)
            wks.cols_axis(normalized_designations)

            # 从 Origin 读取更新后的实际角色
            updated = []
            try:
                actual_desigs = wks.get_labels("D")
            except Exception:
                actual_desigs = []

            for i in range(wks.cols):
                col = wks.get_col(i)
                col_name = (
                    col.name if hasattr(col, "name") else f"Col{i+1}"
                )
                col_info: dict[str, Any] = {
                    "index": i,
                    "name": col_name,
                }
                if i < len(actual_desigs):
                    col_info["designation"] = actual_desigs[i]
                updated.append(col_info)

            return updated

        columns = manager.execute(_set)

        return success_response(
            message=f"列角色已更新为 '{normalized_designations}'。",
            data={
                "sheet_name": target_name,
                "designations": normalized_designations,
                "requested_designations": designations,
                "columns": columns,
            },
            resource=manager.get_resource_context(),
            next_suggestions=["create_plot", "get_worksheet_info"],
        )

    # =================================================================
    # set_column_labels
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置列标签", "请检查列索引和参数值。")
    def set_column_labels(
        col: int,
        sheet_name: str | None = None,
        lname: str | None = None,
        units: str | None = None,
        comments: str | None = None,
    ) -> dict:
        """设置指定列的长名、单位和注释。

        何时使用：需要为列设置描述性的长名(Long Name)、单位(Units)或注释(Comments)时使用。
        何时不用：设置列的 X/Y 角色请用 set_column_designations。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - 未提供的标签不会被修改

        示例：
        - set_column_labels(col=1, lname="Voltage", units="mV")
        - set_column_labels(col=0, lname="Time", units="s", comments="elapsed time")
        """
        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _set_labels(op: Any) -> dict[str, str]:
            wks = _find_worksheet(op, target_name)

            if col < 0 or col >= wks.cols:
                raise ColumnIndexError(col, wks.cols)

            changes: dict[str, str] = {}
            if lname is not None:
                wks.set_label(col, lname, "L")
                changes["long_name"] = lname
            if units is not None:
                wks.set_label(col, units, "U")
                changes["units"] = units
            if comments is not None:
                wks.set_label(col, comments, "C")
                changes["comments"] = comments

            return changes

        result = manager.execute(_set_labels)

        return success_response(
            message=f"列 {col} 的标签已更新。",
            data={
                "sheet_name": target_name,
                "col": col,
                "changes": result,
            },
            resource=manager.get_resource_context(),
            next_suggestions=["get_worksheet_info", "create_plot"],
        )

    # =================================================================
    # sort_worksheet
    # =================================================================

    @mcp.tool()
    @tool_error_handler("排序工作表", "请检查列索引是否正确。")
    def sort_worksheet(
        col: int,
        descending: bool = False,
        sheet_name: str | None = None,
    ) -> dict:
        """按指定列排序工作表数据。

        何时使用：需要对工作表数据按某列进行升序或降序排列时使用。
        何时不用：只需查看数据时请用 get_worksheet_data。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - descending 默认为 False（升序）

        示例：
        - sort_worksheet(col=0)
        - sort_worksheet(col=1, descending=True)
        """
        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _sort(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            if col < 0 or col >= wks.cols:
                raise ColumnIndexError(col, wks.cols)

            # originpro sort 使用 1-offset 索引
            wks.sort(col + 1, descending)
            return {
                "sheet_name": target_name,
                "sort_col": col,
                "descending": descending,
            }

        result = manager.execute(_sort)
        order = "降序" if descending else "升序"

        return success_response(
            message=f"工作表已按列 {col} {order}排序。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["get_worksheet_data", "create_plot"],
        )

    # =================================================================
    # clear_worksheet
    # =================================================================

    @mcp.tool()
    @tool_error_handler("清除工作表", "请检查工作表名称和列索引。")
    def clear_worksheet(
        sheet_name: str | None = None,
        start_col: int | None = None,
        end_col: int | None = None,
    ) -> dict:
        """清除工作表数据。

        何时使用：需要清除工作表中全部或部分列的数据时使用。
        何时不用：需要删除整列（含结构）请用 delete_columns。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - 不指定 start_col 和 end_col 时清除全部数据

        示例：
        - clear_worksheet()
        - clear_worksheet(start_col=1, end_col=3)
        """
        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _clear(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            if start_col is not None and end_col is not None:
                wks.clear(start_col, c2=end_col)
            elif start_col is not None:
                wks.clear(start_col)
            else:
                wks.clear()
            return {
                "sheet_name": target_name,
                "start_col": start_col,
                "end_col": end_col,
            }

        result = manager.execute(_clear)

        if start_col is not None:
            msg = f"已清除列 {start_col}"
            if end_col is not None:
                msg += f" 至 {end_col}"
            msg += " 的数据。"
        else:
            msg = "已清除工作表全部数据。"

        return success_response(
            message=msg,
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["import_csv", "import_data_from_text"],
        )

    # =================================================================
    # set_column_formula
    # =================================================================

    @mcp.tool()
    @tool_error_handler("设置列公式", "请检查公式语法是否正确。公式中可引用列名（如 A, B, C）。")
    def set_column_formula(
        col: int | str,
        formula: str,
        sheet_name: str | None = None,
    ) -> dict:
        """设置工作表列的计算公式。

        何时使用：需要通过公式计算列数据时使用（如基于其他列的数学运算）。
        何时不用：直接导入数据请用 import_csv 或 import_data_from_text。

        默认行为：
        - sheet_name 省略时使用当前活动工作表

        参数说明：
        - formula: Origin 列公式表达式，引用列名（如 "A+1", "sin(A)*2", "B/C"）

        示例：
        - set_column_formula(col=1, formula="A+1")
        - set_column_formula(col="C", formula="sin(A)*B")
        """
        if not formula:
            return error_response(
                message="formula 不能为空",
                error_type="invalid_input",
                target="formula",
                hint="请提供 Origin 列公式表达式，如 'A+1'。",
            )

        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _set_formula(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            wks.set_formula(col, formula)
            return {
                "sheet_name": target_name,
                "col": col,
                "formula": formula,
            }

        result = manager.execute(_set_formula)

        return success_response(
            message=f"列 {col} 的公式已设置为 '{formula}'。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["get_worksheet_data", "create_plot"],
        )

    # =================================================================
    # get_cell_value
    # =================================================================

    @mcp.tool()
    @tool_error_handler("读取单元格", "请检查行列索引是否在范围内。")
    def get_cell_value(
        row: int,
        col: int | str,
        sheet_name: str | None = None,
    ) -> dict:
        """读取工作表中指定单元格的值。

        何时使用：需要精确读取某个单元格数据时使用。
        何时不用：需要查看多行数据请用 get_worksheet_data。

        默认行为：
        - sheet_name 省略时使用当前活动工作表

        示例：
        - get_cell_value(row=0, col=0)
        - get_cell_value(row=5, col="B")
        """
        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _get_cell(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            value = wks.cell(row, col)
            return {
                "sheet_name": target_name,
                "row": row,
                "col": col,
                "value": value,
            }

        result = manager.execute(_get_cell)

        return success_response(
            message=f"单元格 ({row}, {col}) 的值为 '{result['value']}'。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["get_worksheet_data"],
        )

    # =================================================================
    # delete_columns
    # =================================================================

    @mcp.tool()
    @tool_error_handler("删除列", "请检查列索引是否正确。")
    def delete_columns(
        col: int | str,
        count: int = 1,
        sheet_name: str | None = None,
    ) -> dict:
        """删除工作表中的列。

        何时使用：需要删除工作表中不需要的列时使用。
        何时不用：只需清除列数据（保留列结构）请用 clear_worksheet。

        默认行为：
        - sheet_name 省略时使用当前活动工作表
        - count 默认为 1（删除一列）

        示例：
        - delete_columns(col=0)
        - delete_columns(col=2, count=3)
        """
        target_name = _resolve_worksheet_name(sheet_name, manager)

        def _del(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)
            wks.del_col(col, count)
            return {
                "sheet_name": target_name,
                "deleted_from": col,
                "count": count,
                "remaining_cols": wks.cols,
            }

        result = manager.execute(_del)

        return success_response(
            message=f"已删除 {count} 列，剩余 {result['remaining_cols']} 列。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=["get_worksheet_info"],
        )

    # =================================================================
    # add_worksheet
    # =================================================================

    @mcp.tool()
    @tool_error_handler("添加工作表", "请检查工作簿名称是否正确。")
    def add_worksheet(
        sheet_name: str | None = None,
        book_name: str | None = None,
    ) -> dict:
        """在工作簿中添加新的工作表。

        何时使用：需要在已有工作簿中创建新工作表时使用。
        何时不用：导入数据时会自动创建工作表，无需手动调用。

        默认行为：
        - book_name 省略时使用当前活动工作表所在的工作簿
        - sheet_name 省略时自动命名

        示例：
        - add_worksheet(sheet_name="Results")
        - add_worksheet(sheet_name="Summary", book_name="Book1")
        """
        def _add(op: Any) -> dict[str, Any]:
            if book_name:
                # 查找指定的工作簿
                target_book = None
                for book in op.pages("Book"):
                    if book.name == book_name:
                        target_book = book
                        break
                if target_book is None:
                    raise ToolError(
                        f"工作簿 '{book_name}' 不存在",
                        error_type="not_found",
                        target="book_name",
                        value=book_name,
                        hint="请调用 list_worksheets 查看可用的工作簿。",
                    )
                new_wks = target_book.add_sheet(sheet_name or "")
            else:
                # 使用活动工作表所在的工作簿
                active_name = manager.active_worksheet
                if active_name:
                    wks = op.find_sheet("w", active_name)
                    if wks:
                        target_book = wks.get_book()
                        new_wks = target_book.add_sheet(sheet_name or "")
                    else:
                        new_wks = op.new_sheet(lname=sheet_name)
                else:
                    new_wks = op.new_sheet(lname=sheet_name)

            actual_book = new_wks.get_book().name
            actual_name = new_wks.name
            full_name = f"[{actual_book}]{actual_name}"
            manager.active_worksheet = full_name

            return {
                "book_name": actual_book,
                "sheet_name": actual_name,
                "full_name": full_name,
            }

        result = manager.execute(_add)

        return success_response(
            message=f"新工作表 '{result['full_name']}' 已创建。",
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "import_data_from_text",
                "set_column_designations",
            ],
        )
