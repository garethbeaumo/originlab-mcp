"""MCP Resources for read-only Origin session inspection.

These resources implement the MCP `resources/read` surface so clients can
inspect the current Origin project without calling mutation tools.
"""

from __future__ import annotations

import json
from typing import Any

from originlab_mcp.session import (
    build_session_snapshot,
    collect_graphs,
    collect_worksheets,
    decode_resource_segment,
    read_graph_detail,
    read_worksheet_detail,
    worksheet_full_name,
)
from originlab_mcp.utils.constants import DEFAULT_MAX_PREVIEW_ROWS

JSON_MIME = "application/json"


def dumps_resource(payload: dict[str, Any]) -> str:
    """Serialize a resource payload as stable JSON text."""
    return json.dumps(payload, ensure_ascii=False, indent=2)


def _error_payload(message: str, hint: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {"ok": False, "message": message}
    if hint:
        payload["hint"] = hint
    return payload


def _connection_hint() -> str:
    return (
        "请确认 OriginLab 已安装并持有有效许可证。"
        "可调用 get_origin_info 或 read_origin_session 检查状态。"
    )


def register_session_resources(mcp, manager) -> None:
    """Register Origin session MCP resources.

    Args:
        mcp: FastMCP instance.
        manager: OriginManager instance (dependency injection).
    """

    def _read(builder):
        try:
            def _run(op):
                manager.peek_active_context()
                return builder(op)

            data = manager.execute(_run)
            return dumps_resource({"ok": True, "data": data})
        except Exception as exc:
            return dumps_resource(_error_payload(str(exc), _connection_hint()))

    @mcp.resource(
        "originlab://session",
        name="OriginSession",
        mime_type=JSON_MIME,
    )
    def origin_session() -> str:
        """Read a snapshot of the current Origin project.

        Includes workbooks, worksheets, graphs, matrices, notes, active
        objects, and project metadata. This resource is read-only.
        """
        return _read(
            lambda op: build_session_snapshot(
                op,
                active_worksheet=manager.active_worksheet,
                active_graph=manager.active_graph,
            )
        )

    @mcp.resource(
        "originlab://worksheets",
        name="OriginWorksheets",
        mime_type=JSON_MIME,
    )
    def origin_worksheets() -> str:
        """List worksheets in the current Origin project."""
        return _read(collect_worksheets)

    @mcp.resource(
        "originlab://graphs",
        name="OriginGraphs",
        mime_type=JSON_MIME,
    )
    def origin_graphs() -> str:
        """List graphs in the current Origin project."""
        return _read(collect_graphs)

    @mcp.resource(
        "originlab://worksheet/{book}/{sheet}",
        name="OriginWorksheet",
        mime_type=JSON_MIME,
    )
    def origin_worksheet(book: str, sheet: str) -> str:
        """Read one worksheet's columns and a data preview.

        URI example: originlab://worksheet/Book1/Sheet1
        """
        sheet_name = worksheet_full_name(
            decode_resource_segment(book),
            decode_resource_segment(sheet),
        )
        return _read(
            lambda op: read_worksheet_detail(
                op,
                sheet_name,
                include_preview=True,
                max_preview_rows=DEFAULT_MAX_PREVIEW_ROWS,
            )
        )

    @mcp.resource(
        "originlab://graph/{name}",
        name="OriginGraph",
        mime_type=JSON_MIME,
    )
    def origin_graph(name: str) -> str:
        """Read one graph's layers and curve list.

        URI example: originlab://graph/Graph1
        """
        return _read(lambda op: read_graph_detail(op, decode_resource_segment(name)))
