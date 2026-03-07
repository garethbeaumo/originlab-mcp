"""
Tool 基础测试

测试 validators 模块的返回结构和校验函数。
这些测试不依赖 Origin COM，可以在任何环境运行。
"""

import pytest

from originlab_mcp.utils.constants import (
    ColumnDesignation,
    ErrorType,
    ExportFormat,
    PlotType,
    ScaleType,
)
from originlab_mcp.utils.validators import (
    error_response,
    normalize_y_cols,
    success_response,
    validate_column_index,
    validate_column_indices,
    validate_designation,
    validate_export_format,
    validate_plot_type,
    validate_scale_type,
)


# ===================================================================
# success_response 测试
# ===================================================================


class TestSuccessResponse:
    def test_basic(self):
        r = success_response(message="done")
        assert r["ok"] is True
        assert r["message"] == "done"
        assert r["data"] == {}
        assert r["resource"] == {}
        assert r["warnings"] == []
        assert r["next_suggestions"] == []

    def test_with_data(self):
        r = success_response(
            message="ok",
            data={"key": "val"},
            resource={"active_worksheet": "Sheet1"},
            warnings=["warn1"],
            next_suggestions=["next_tool"],
        )
        assert r["data"]["key"] == "val"
        assert r["resource"]["active_worksheet"] == "Sheet1"
        assert len(r["warnings"]) == 1
        assert "next_tool" in r["next_suggestions"]


# ===================================================================
# error_response 测试
# ===================================================================


class TestErrorResponse:
    def test_basic(self):
        r = error_response(
            message="not found",
            error_type=ErrorType.NOT_FOUND,
            target="worksheet",
            value="Sheet9",
            hint="Call list_worksheets",
        )
        assert r["ok"] is False
        assert r["error"]["type"] == "not_found"
        assert r["error"]["target"] == "worksheet"
        assert r["error"]["value"] == "Sheet9"
        assert r["error"]["hint"] == "Call list_worksheets"

    def test_string_type(self):
        r = error_response(
            message="error",
            error_type="custom_error",
            target="x",
        )
        assert r["error"]["type"] == "custom_error"


# ===================================================================
# 验证函数测试
# ===================================================================


class TestValidateColumnIndex:
    def test_valid(self):
        assert validate_column_index(0, 3) is None
        assert validate_column_index(2, 3) is None

    def test_out_of_range(self):
        err = validate_column_index(5, 3)
        assert err is not None
        assert "5" in err
        assert "3" in err

    def test_negative(self):
        err = validate_column_index(-1, 3)
        assert err is not None


class TestValidateColumnIndices:
    def test_valid(self):
        assert validate_column_indices([0, 1, 2], 3) is None

    def test_one_bad(self):
        err = validate_column_indices([0, 1, 5], 3)
        assert err is not None


class TestValidatePlotType:
    def test_valid(self):
        for pt in PlotType:
            assert validate_plot_type(pt.value) is None

    def test_invalid(self):
        err = validate_plot_type("heatmap")
        assert err is not None
        assert "heatmap" in err


class TestValidateDesignation:
    def test_valid(self):
        for d in ColumnDesignation:
            assert validate_designation(d.value) is None

    def test_invalid(self):
        err = validate_designation("W")
        assert err is not None


class TestValidateScaleType:
    def test_valid(self):
        for s in ScaleType:
            assert validate_scale_type(s.value) is None

    def test_invalid(self):
        err = validate_scale_type("exp")
        assert err is not None


class TestValidateExportFormat:
    def test_valid(self):
        for f in ExportFormat:
            assert validate_export_format(f.value) is None

    def test_invalid(self):
        err = validate_export_format("bmp")
        assert err is not None


# ===================================================================
# normalize_y_cols 测试
# ===================================================================


class TestNormalizeYCols:
    def test_single_int(self):
        assert normalize_y_cols(1) == [1]

    def test_list(self):
        assert normalize_y_cols([1, 2, 3]) == [1, 2, 3]

    def test_empty_list(self):
        assert normalize_y_cols([]) == []
