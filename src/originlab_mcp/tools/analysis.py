"""数据分析类 tools。

让 AI 能利用 Origin 内置的拟合引擎进行数据分析。

包含:
    linear_fit: 线性拟合
    nonlinear_fit: 非线性拟合
    list_fit_functions: 列出可用拟合函数
"""

from __future__ import annotations

from typing import Any

from originlab_mcp.exceptions import (
    FitConvergenceError,
    ToolError,
)
from originlab_mcp.utils.constants import COMMON_FIT_FUNCTIONS
from originlab_mcp.utils.helpers import (
    find_worksheet as _find_worksheet,
)
from originlab_mcp.utils.helpers import (
    resolve_worksheet_name,
    tool_error_handler,
)
from originlab_mcp.utils.validators import (
    success_response,
)


def register_analysis_tools(mcp: Any, manager: Any) -> None:
    """注册数据分析类 tools 到 MCP Server。

    Args:
        mcp: FastMCP 实例。
        manager: OriginManager 实例（依赖注入）。
    """

    # =================================================================
    # linear_fit
    # =================================================================

    @mcp.tool()
    @tool_error_handler("线性拟合", "请检查列索引和数据是否有效。")
    def linear_fit(
        x_col: int,
        y_col: int,
        sheet_name: str | None = None,
        yerr_col: int | None = None,
        fix_slope: float | None = None,
        fix_intercept: float | None = None,
        confidence_band: bool = False,
    ) -> dict:
        """Perform linear fit on worksheet data (y = Intercept + Slope * x).

        When to use: For linear regression analysis to obtain slope and intercept.
        When not to use: For clearly non-linear data, use nonlinear_fit.

        Default behavior:
        - sheet_name omitted: uses current active worksheet
        - Returns slope, intercept with errors, and R² statistics

        Examples:
        - linear_fit(x_col=0, y_col=1)
        - linear_fit(x_col=0, y_col=1, yerr_col=2, confidence_band=True)
        - linear_fit(x_col=0, y_col=1, fix_intercept=0)
        """
        target_name = resolve_worksheet_name(sheet_name, manager)

        def _fit(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)

            lr = op.LinearFit()

            # 设置数据
            if yerr_col is not None:
                lr.set_data(wks, x_col, y_col, yerr_col)
            else:
                lr.set_data(wks, x_col, y_col)

            # 可选：固定斜率或截距
            if fix_slope is not None:
                lr.fix_slope = fix_slope
            if fix_intercept is not None:
                lr.fix_intercept = fix_intercept

            # 执行拟合
            if confidence_band:
                band = 3  # 同时生成置信带和预测带
                r, c = lr.report(band)
                result = lr.result() if hasattr(lr, 'result') else {}
                report_info = {
                    "report_range": str(r) if r else None,
                    "curves_range": str(c) if c else None,
                }
            else:
                result = lr.result()
                report_info = {}

            # 提取拟合参数
            fit_result: dict[str, Any] = {
                "sheet_name": target_name,
                "x_col": x_col,
                "y_col": y_col,
                "method": "linear",
            }

            if result:
                # 提取 Parameters
                params = result.get("Parameters", {})
                fit_result["parameters"] = {}
                for pname, pinfo in params.items():
                    if isinstance(pinfo, dict):
                        fit_result["parameters"][pname] = {
                            "value": pinfo.get("Value"),
                            "error": pinfo.get("Error"),
                        }

                # 提取统计量
                stats = result.get("Statistics", {})
                if stats:
                    fit_result["statistics"] = stats

            if report_info:
                fit_result["report"] = report_info

            if fix_slope is not None:
                fit_result["fixed_slope"] = fix_slope
            if fix_intercept is not None:
                fit_result["fixed_intercept"] = fix_intercept

            return fit_result

        result = manager.execute(_fit)

        # 构建消息
        params = result.get("parameters", {})
        slope_info = params.get("Slope", {})
        intercept_info = params.get("Intercept", {})
        msg_parts = ["线性拟合完成。"]
        if slope_info.get("value") is not None:
            msg_parts.append(
                f"Slope = {slope_info['value']:.6g}"
                + (f" ± {slope_info['error']:.6g}" if slope_info.get('error') else "")
            )
        if intercept_info.get("value") is not None:
            msg_parts.append(
                f"Intercept = {intercept_info['value']:.6g}"
                + (f" ± {intercept_info['error']:.6g}" if intercept_info.get('error') else "")
            )

        return success_response(
            message=" ".join(msg_parts),
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "create_plot",
                "export_graph",
                "nonlinear_fit",
            ],
        )

    # =================================================================
    # nonlinear_fit
    # =================================================================

    @mcp.tool()
    @tool_error_handler("非线性拟合", "请检查函数名、列索引和数据是否有效。调用 list_fit_functions 查看可用函数。")
    def nonlinear_fit(
        function_name: str,
        x_col: int,
        y_col: int,
        sheet_name: str | None = None,
        yerr_col: int | None = None,
        initial_params: dict | None = None,
        fixed_params: dict | None = None,
        generate_report: bool = False,
    ) -> dict:
        """Perform nonlinear curve fitting on worksheet data.

        When to use: To fit data with specific functions (e.g. Gauss, Lorentz, ExpDec1).
        When not to use: For simple linear relationships, use linear_fit.

        Default behavior:
        - sheet_name omitted: uses current active worksheet
        - function_name must be an Origin built-in fit function name

        Parameter notes:
        - function_name: Origin built-in function name (e.g. "Gauss", "Lorentz", "ExpDec1", "Boltzmann")
        - initial_params: optional initial parameter values dict (e.g. {"xc": 0.5, "w": 1.0})
        - fixed_params: optional fixed parameter dict (e.g. {"y0": 0}). Value of false unfixes the parameter

        Examples:
        - nonlinear_fit(function_name="Gauss", x_col=0, y_col=1)
        - nonlinear_fit(function_name="ExpDec1", x_col=0, y_col=1, initial_params={"t1": 5.0})
        - nonlinear_fit(function_name="Gauss", x_col=0, y_col=1, fixed_params={"y0": 0})
        """
        target_name = resolve_worksheet_name(sheet_name, manager)

        def _fit(op: Any) -> dict[str, Any]:
            wks = _find_worksheet(op, target_name)

            try:
                model = op.NLFit(function_name)
            except Exception as e:
                raise ToolError(
                    f"无法创建拟合模型 '{function_name}': {e}",
                    error_type="not_found",
                    target="function_name",
                    value=function_name,
                    hint="请调用 list_fit_functions 查看常用函数名，或确认 Origin 中已安装该拟合函数。",
                ) from e

            # 设置数据
            if yerr_col is not None:
                model.set_data(wks, x_col, y_col, yerr=yerr_col)
            else:
                model.set_data(wks, x_col, y_col)

            # 设置初始参数
            if initial_params:
                for pname, pval in initial_params.items():
                    model.set_param(pname, pval)

            # 固定参数
            if fixed_params:
                for pname, pval in fixed_params.items():
                    model.fix_param(pname, pval)

            # 执行拟合
            try:
                model.fit()
            except Exception as e:
                raise FitConvergenceError(function_name, str(e)) from e

            # 获取结果
            fit_result: dict[str, Any] = {
                "sheet_name": target_name,
                "x_col": x_col,
                "y_col": y_col,
                "function_name": function_name,
            }

            if generate_report:
                try:
                    r, c = model.report()
                    fit_result["report"] = {
                        "report_range": str(r) if r else None,
                        "curves_range": str(c) if c else None,
                    }
                except Exception:
                    pass
                # report 之后不能再调用 result
            else:
                try:
                    result = model.result()
                    if result:
                        params = result.get("Parameters", {})
                        fit_result["parameters"] = {}
                        for pname, pinfo in params.items():
                            if isinstance(pinfo, dict):
                                fit_result["parameters"][pname] = {
                                    "value": pinfo.get("Value"),
                                    "error": pinfo.get("Error"),
                                }

                        stats = result.get("Statistics", {})
                        if stats:
                            fit_result["statistics"] = stats
                except Exception:
                    pass

            return fit_result

        result = manager.execute(_fit)

        # 构建消息
        params = result.get("parameters", {})
        param_strs = []
        for pname, pinfo in params.items():
            val = pinfo.get("value")
            err = pinfo.get("error")
            if val is not None:
                s = f"{pname} = {val:.6g}"
                if err is not None:
                    s += f" ± {err:.6g}"
                param_strs.append(s)

        msg = f"非线性拟合完成（函数: {function_name}）。"
        if param_strs:
            msg += " " + ", ".join(param_strs[:4])
            if len(param_strs) > 4:
                msg += f"...（共 {len(param_strs)} 个参数）"

        return success_response(
            message=msg,
            data=result,
            resource=manager.get_resource_context(),
            next_suggestions=[
                "create_plot",
                "export_graph",
                "linear_fit",
            ],
        )

    # =================================================================
    # list_fit_functions
    # =================================================================

    @mcp.tool()
    def list_fit_functions() -> dict:
        """List commonly used fit functions with parameter descriptions.

        When to use: When unsure which fit function to use; browse available functions and parameter info.
        When not to use: If the fit function name is already known.

        Examples:
        - list_fit_functions()
        """
        functions = []
        for fname, finfo in COMMON_FIT_FUNCTIONS.items():
            functions.append({
                "name": fname,
                "parameters": list(finfo.params),
                "description": finfo.description,
            })

        return success_response(
            message=f"共有 {len(functions)} 个常用拟合函数可用。Origin 还支持更多内置函数。",
            data={"functions": functions},
            resource=manager.get_resource_context(),
            next_suggestions=["linear_fit", "nonlinear_fit"],
        )
