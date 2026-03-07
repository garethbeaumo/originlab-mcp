"""
数据类 tools

让 AI 能导入数据、理解工作表结构、设置列语义。

包含：
- import_csv
- import_excel
- import_data_from_text
- list_worksheets
- get_worksheet_info
- get_worksheet_data
- set_column_designations
- set_column_labels
"""

from __future__ import annotations

import os

from originlab_mcp.origin_manager import OriginManager
from originlab_mcp.utils.constants import (
    DEFAULT_HAS_HEADER,
    DEFAULT_MAX_PREVIEW_ROWS,
    DEFAULT_SEPARATOR,
    ColumnDesignation,
)
from originlab_mcp.utils.validators import (
    error_response,
    normalize_y_cols,
    success_response,
    validate_column_index,
    validate_designation,
    validate_file_path,
)


def register_data_tools(mcp) -> None:
    """注册数据类 tools 到 MCP Server。"""

    manager = OriginManager()

    # =================================================================
    # import_csv
    # =================================================================

    @mcp.tool()
    def import_csv(file_path: str, sheet_name: str | None = None) -> dict:
        """导入 CSV 文件到工作表。

        何时使用：需要将本地 CSV 文件导入 Origin 工作表时使用。
        何时不用：导入 Excel 文件请用 import_excel；直接传入文本数据请用 import_data_from_text。

        默认行为：
        - sheet_name 省略时创建新工作表

        示例：
        - import_csv(file_path="C:\\\\data\\\\exp.csv")
        - import_csv(file_path="C:\\\\data\\\\exp.csv", sheet_name="RawData")
        """
        # 验证文件路径
        err = validate_file_path(file_path)
        if err:
            return error_response(
                message=err,
                error_type="invalid_input",
                target="file_path",
                value=file_path,
                hint="File does not exist. Please verify the file path and try again.",
            )

        try:
            def _import(op):
                # 创建或获取工作表
                if sheet_name:
                    wks = op.find_sheet("w", sheet_name)
                    if wks is None:
                        wks = op.new_sheet(lname=sheet_name)
                    created_new = wks is not None
                else:
                    wks = op.new_sheet()
                    created_new = True

                wks.from_file(file_path)

                actual_name = wks.name
                rows = wks.rows
                cols = wks.cols

                # 更新活动工作表
                manager.active_worksheet = actual_name

                return actual_name, rows, cols, created_new

            name, rows, cols, created = manager.execute(_import)

            return success_response(
                message=f"CSV 已导入到工作表 '{name}'，共 {rows} 行 {cols} 列。",
                data={
                    "sheet_name": name,
                    "rows": rows,
                    "cols": cols,
                    "created_new_sheet": created,
                    "source_file": file_path,
                },
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "get_worksheet_info",
                    "set_column_designations",
                    "create_plot",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"导入 CSV 失败: {e}",
                error_type="internal_error",
                target="file_path",
                value=file_path,
                hint="请检查文件格式是否为有效的 CSV，或尝试使用 import_data_from_text 手动传入数据。",
            )

    # =================================================================
    # import_excel
    # =================================================================

    @mcp.tool()
    def import_excel(file_path: str, sheet_name: str | None = None) -> dict:
        """导入 Excel 文件到工作表。

        何时使用：需要将本地 Excel (.xlsx/.xls) 文件导入 Origin 工作表时使用。
        何时不用：导入 CSV 请用 import_csv。

        默认行为：
        - sheet_name 省略时创建新工作表
        - 默认导入 Excel 文件的第一个工作页

        示例：
        - import_excel(file_path="C:\\\\data\\\\results.xlsx")
        """
        err = validate_file_path(file_path)
        if err:
            return error_response(
                message=err,
                error_type="invalid_input",
                target="file_path",
                value=file_path,
                hint="File does not exist. Please verify the file path and try again.",
            )

        try:
            def _import(op):
                if sheet_name:
                    wks = op.find_sheet("w", sheet_name)
                    if wks is None:
                        wks = op.new_sheet(lname=sheet_name)
                else:
                    wks = op.new_sheet()

                wks.from_file(file_path)

                actual_name = wks.name
                manager.active_worksheet = actual_name
                return actual_name, wks.rows, wks.cols

            name, rows, cols = manager.execute(_import)

            return success_response(
                message=f"Excel 已导入到工作表 '{name}'，共 {rows} 行 {cols} 列。",
                data={
                    "sheet_name": name,
                    "rows": rows,
                    "cols": cols,
                    "source_file": file_path,
                },
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "get_worksheet_info",
                    "set_column_designations",
                    "create_plot",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"导入 Excel 失败: {e}",
                error_type="internal_error",
                target="file_path",
                value=file_path,
                hint="请检查文件格式是否为有效的 Excel 文件。",
            )

    # =================================================================
    # import_data_from_text
    # =================================================================

    @mcp.tool()
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

        try:
            lines = data.strip().split("\n")
            if len(lines) < 1:
                return error_response(
                    message="data 至少需要包含一行",
                    error_type="invalid_input",
                    target="data",
                    hint="请提供至少一行数据。",
                )

            # 解析数据
            header = None
            start_row = 0
            if has_header and len(lines) > 1:
                header = [h.strip() for h in lines[0].split(separator)]
                start_row = 1

            rows_data = []
            for line in lines[start_row:]:
                row = [cell.strip() for cell in line.split(separator)]
                rows_data.append(row)

            num_cols = len(rows_data[0]) if rows_data else (len(header) if header else 0)

            def _import(op):
                wks = op.new_sheet(lname=sheet_name) if sheet_name else op.new_sheet()

                # 逐列写入数据
                for col_idx in range(num_cols):
                    col_data = []
                    for row in rows_data:
                        val = row[col_idx] if col_idx < len(row) else ""
                        # 尝试转换为数值
                        try:
                            val = float(val)
                            if val == int(val):
                                val = int(val)
                        except (ValueError, TypeError):
                            pass
                        col_data.append(val)

                    lname = header[col_idx] if header and col_idx < len(header) else None
                    wks.from_list(col_idx, col_data, lname=lname)

                actual_name = wks.name
                manager.active_worksheet = actual_name
                return actual_name, len(rows_data), num_cols

            name, n_rows, n_cols = manager.execute(_import)

            return success_response(
                message=f"文本数据已导入到工作表 '{name}'，共 {n_rows} 行 {n_cols} 列。",
                data={
                    "sheet_name": name,
                    "rows": n_rows,
                    "cols": n_cols,
                    "separator_used": separator,
                    "has_header": has_header,
                },
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "get_worksheet_info",
                    "set_column_designations",
                    "create_plot",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"从文本数据导入失败: {e}",
                error_type="internal_error",
                target="data",
                hint="请检查数据格式。数据应以换行符分行，以指定的分隔符分列。",
            )

    # =================================================================
    # list_worksheets
    # =================================================================

    @mcp.tool()
    def list_worksheets() -> dict:
        """列出当前项目中的所有工作表。

        何时使用：需要查看有哪些可用的工作表，或确认工作表名称时使用。
        何时不用：已经知道工作表名称时无需调用。

        示例：
        - list_worksheets()
        """
        try:
            def _list(op):
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
                data={
                    "worksheets": sheets,
                    "count": len(sheets),
                },
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "get_worksheet_info",
                    "import_csv",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"列出工作表失败: {e}",
                error_type="internal_error",
                target="worksheets",
                hint="请确认 Origin 已连接。可调用 get_origin_info 检查状态。",
            )

    # =================================================================
    # get_worksheet_info
    # =================================================================

    @mcp.tool()
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
        target_name = sheet_name or manager.active_worksheet

        if not target_name:
            return error_response(
                message="未指定工作表且无活动工作表",
                error_type="invalid_input",
                target="sheet_name",
                hint="请指定 sheet_name，或先调用 import_csv / list_worksheets 获取可用工作表。",
            )

        try:
            def _info(op):
                wks = op.find_sheet("w", target_name)
                if wks is None:
                    return None

                columns = []
                for i in range(wks.cols):
                    col = wks.get_col(i)
                    col_info = {
                        "index": i,
                        "name": col.name if hasattr(col, "name") else f"Col{i+1}",
                        "long_name": col.get_label("L") if hasattr(col, "get_label") else "",
                        "units": col.get_label("U") if hasattr(col, "get_label") else "",
                        "comments": col.get_label("C") if hasattr(col, "get_label") else "",
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

            if result is None:
                return error_response(
                    message=f"工作表 '{target_name}' 不存在",
                    error_type="not_found",
                    target="worksheet",
                    value=target_name,
                    hint="Call list_worksheets to inspect available worksheet names.",
                )

            return success_response(
                message=f"工作表 '{target_name}' 共 {result['rows']} 行 {result['cols']} 列。",
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_column_designations",
                    "get_worksheet_data",
                    "create_plot",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"获取工作表信息失败: {e}",
                error_type="internal_error",
                target="sheet_name",
                value=target_name,
                hint="请确认工作表名称正确。调用 list_worksheets 查看可用工作表。",
            )

    # =================================================================
    # get_worksheet_data
    # =================================================================

    @mcp.tool()
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
        target_name = sheet_name or manager.active_worksheet

        if not target_name:
            return error_response(
                message="未指定工作表且无活动工作表",
                error_type="invalid_input",
                target="sheet_name",
                hint="请指定 sheet_name，或先调用 import_csv / list_worksheets。",
            )

        try:
            def _data(op):
                wks = op.find_sheet("w", target_name)
                if wks is None:
                    return None

                df = wks.to_df()
                total_rows = len(df)
                truncated = total_rows > max_rows

                if truncated:
                    df = df.head(max_rows)

                # 转为简单的列表结构
                records = df.to_dict(orient="records")

                return {
                    "sheet_name": target_name,
                    "total_rows": total_rows,
                    "returned_rows": len(records),
                    "truncated": truncated,
                    "columns": list(df.columns),
                    "data": records,
                }

            result = manager.execute(_data)

            if result is None:
                return error_response(
                    message=f"工作表 '{target_name}' 不存在",
                    error_type="not_found",
                    target="worksheet",
                    value=target_name,
                    hint="Call list_worksheets to inspect available worksheet names.",
                )

            msg = f"返回 {result['returned_rows']} 行数据"
            if result["truncated"]:
                msg += f"（共 {result['total_rows']} 行，已截断）"
            msg += "。"

            return success_response(
                message=msg,
                data=result,
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "set_column_designations",
                    "create_plot",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"获取工作表数据失败: {e}",
                error_type="internal_error",
                target="sheet_name",
                value=target_name,
                hint="请确认工作表名称正确。",
            )

    # =================================================================
    # set_column_designations
    # =================================================================

    @mcp.tool()
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
        target_name = sheet_name or manager.active_worksheet

        if not target_name:
            return error_response(
                message="未指定工作表且无活动工作表",
                error_type="invalid_input",
                target="sheet_name",
                hint="请指定 sheet_name 或先导入数据。",
            )

        if not designations:
            return error_response(
                message="designations 不能为空",
                error_type="invalid_input",
                target="designations",
                hint="请提供列角色字符串，如 'XYY'。支持的角色: X, Y, Z, YErr, L, N",
            )

        try:
            def _set(op):
                wks = op.find_sheet("w", target_name)
                if wks is None:
                    return None, "not_found"

                wks.cols_axis(designations)

                # 读取更新后的角色
                updated = []
                for i in range(wks.cols):
                    col = wks.get_col(i)
                    col_name = col.name if hasattr(col, "name") else f"Col{i+1}"
                    updated.append({"index": i, "name": col_name})

                return updated, "ok"

            result, status = manager.execute(_set)

            if status == "not_found":
                return error_response(
                    message=f"工作表 '{target_name}' 不存在",
                    error_type="not_found",
                    target="worksheet",
                    value=target_name,
                    hint="Call list_worksheets to inspect available worksheet names.",
                )

            return success_response(
                message=f"列角色已更新为 '{designations}'。",
                data={
                    "sheet_name": target_name,
                    "designations": designations,
                    "columns": result,
                },
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "create_plot",
                    "get_worksheet_info",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"设置列角色失败: {e}",
                error_type="internal_error",
                target="designations",
                value=designations,
                hint="请检查角色字符串长度是否与列数匹配。",
            )

    # =================================================================
    # set_column_labels
    # =================================================================

    @mcp.tool()
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
        target_name = sheet_name or manager.active_worksheet

        if not target_name:
            return error_response(
                message="未指定工作表且无活动工作表",
                error_type="invalid_input",
                target="sheet_name",
                hint="请指定 sheet_name 或先导入数据。",
            )

        try:
            def _set_labels(op):
                wks = op.find_sheet("w", target_name)
                if wks is None:
                    return None, "not_found"

                if col < 0 or col >= wks.cols:
                    return col, "out_of_range"

                changes = {}
                if lname is not None:
                    wks.set_label(col, lname, "L")
                    changes["long_name"] = lname
                if units is not None:
                    wks.set_label(col, units, "U")
                    changes["units"] = units
                if comments is not None:
                    wks.set_label(col, comments, "C")
                    changes["comments"] = comments

                return changes, "ok"

            result, status = manager.execute(_set_labels)

            if status == "not_found":
                return error_response(
                    message=f"工作表 '{target_name}' 不存在",
                    error_type="not_found",
                    target="worksheet",
                    value=target_name,
                    hint="Call list_worksheets to inspect available worksheet names.",
                )

            if status == "out_of_range":
                return error_response(
                    message=f"列索引 {col} 超出范围",
                    error_type="invalid_input",
                    target="col",
                    value=col,
                    hint="Call get_worksheet_info to inspect column count.",
                )

            return success_response(
                message=f"列 {col} 的标签已更新。",
                data={
                    "sheet_name": target_name,
                    "col": col,
                    "changes": result,
                },
                resource=manager.get_resource_context(),
                next_suggestions=[
                    "get_worksheet_info",
                    "create_plot",
                ],
            )
        except Exception as e:
            return error_response(
                message=f"设置列标签失败: {e}",
                error_type="internal_error",
                target="col",
                value=col,
                hint="请检查列索引和参数值。",
            )
