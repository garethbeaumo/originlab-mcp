"""Read-only Origin session snapshot helpers.

These functions inspect the current Origin project without changing it.
They accept an originpro-like `op` object so they can be unit-tested
without a live Origin COM session.
"""

from __future__ import annotations

from contextlib import suppress
from typing import Any
from urllib.parse import unquote

from originlab_mcp.utils.constants import DEFAULT_MAX_PREVIEW_ROWS
from originlab_mcp.utils.helpers import (
    find_graph,
    find_worksheet,
    get_column_info,
    get_plot_count,
)

PAGE_KIND_BOOK = "w"
PAGE_KIND_GRAPH = "g"
PAGE_KIND_MATRIX = "m"
PAGE_KIND_NOTES = "Notes"
PAGE_KIND_EXCEL = "Excel"


def decode_resource_segment(value: str) -> str:
    """Decode a URI template path segment."""
    return unquote(value or "").strip()


def worksheet_full_name(book_name: str, sheet_name: str) -> str:
    """Return Origin range-style worksheet name `[Book]Sheet`."""
    return f"[{book_name}]{sheet_name}"


def iter_pages(op: Any, kind: str) -> list[Any]:
    """Return pages of the given Origin type, or an empty list on failure."""
    try:
        pages = op.pages(kind)
    except Exception:
        return []
    if pages is None:
        return []
    try:
        return list(pages)
    except TypeError:
        return []


def _page_name(page: Any, fallback: str = "") -> str:
    name = getattr(page, "name", None)
    if name:
        return str(name)
    return str(page) if page is not None else fallback


def _optional_len(obj: Any) -> int | None:
    if obj is None or not hasattr(obj, "__len__"):
        return None
    try:
        return len(obj)
    except Exception:
        return None


def _safe_path(op: Any, kind: str) -> str | None:
    with suppress(Exception):
        value = op.path(kind)
        if value:
            return str(value)
    return None


def _safe_lt_str(op: Any, name: str) -> str | None:
    getter = getattr(op, "get_lt_str", None)
    if getter is None:
        return None
    with suppress(Exception):
        value = getter(name)
        if value:
            return str(value)
    return None


def collect_project_meta(op: Any) -> dict[str, str]:
    """Collect project file metadata from LabTalk string variables."""
    meta: dict[str, str] = {}
    mapping = (
        ("path", ("pe_path$", "filename$", "doc.path$")),
        ("name", ("pe_name$", "doc.name$", "page.name$")),
    )
    for key, variables in mapping:
        for variable in variables:
            value = _safe_lt_str(op, variable)
            if value:
                meta[key] = value
                break
    return meta


def collect_worksheets(op: Any) -> list[dict[str, Any]]:
    """List workbooks/worksheets in the current project."""
    worksheets: list[dict[str, Any]] = []
    for book in iter_pages(op, PAGE_KIND_BOOK):
        book_name = _page_name(book)
        try:
            sheets = list(book)
        except TypeError:
            sheets = []
        if not sheets:
            worksheets.append(
                {
                    "book_name": book_name,
                    "sheet_name": None,
                    "full_name": f"[{book_name}]",
                    "rows": None,
                    "cols": None,
                }
            )
            continue
        for sheet in sheets:
            sheet_name = _page_name(sheet)
            worksheets.append(
                {
                    "book_name": book_name,
                    "sheet_name": sheet_name,
                    "full_name": worksheet_full_name(book_name, sheet_name),
                    "rows": getattr(sheet, "rows", None),
                    "cols": getattr(sheet, "cols", None),
                }
            )
    return worksheets


def collect_graphs(op: Any) -> list[dict[str, Any]]:
    """List graph pages in the current project."""
    graphs: list[dict[str, Any]] = []
    for graph in iter_pages(op, PAGE_KIND_GRAPH):
        item: dict[str, Any] = {
            "graph_name": _page_name(graph),
            "layers": _optional_len(graph),
        }
        long_name = getattr(graph, "lname", None)
        if long_name:
            item["long_name"] = str(long_name)
        graphs.append(item)
    return graphs


def collect_matrix_sheets(op: Any) -> list[dict[str, Any]]:
    """List matrix books/sheets when Origin exposes them."""
    items: list[dict[str, Any]] = []
    for book in iter_pages(op, PAGE_KIND_MATRIX):
        book_name = _page_name(book)
        try:
            sheets = list(book)
        except TypeError:
            sheets = []
        if not sheets:
            items.append(
                {
                    "book_name": book_name,
                    "sheet_name": None,
                    "full_name": book_name,
                    "rows": getattr(book, "rows", None),
                    "cols": getattr(book, "cols", None),
                }
            )
            continue
        for sheet in sheets:
            sheet_name = _page_name(sheet)
            items.append(
                {
                    "book_name": book_name,
                    "sheet_name": sheet_name,
                    "full_name": worksheet_full_name(book_name, sheet_name),
                    "rows": getattr(sheet, "rows", None),
                    "cols": getattr(sheet, "cols", None),
                }
            )
    return items


def collect_named_pages(op: Any, kind: str) -> list[dict[str, Any]]:
    """List named Origin pages such as Notes or Excel books."""
    items: list[dict[str, Any]] = []
    for page in iter_pages(op, kind):
        item: dict[str, Any] = {"name": _page_name(page)}
        long_name = getattr(page, "lname", None)
        if long_name:
            item["long_name"] = str(long_name)
        layers = _optional_len(page)
        if layers is not None:
            item["layers"] = layers
        items.append(item)
    return items


def _unique_column_name(name: str, used: set[str]) -> str:
    if name not in used:
        used.add(name)
        return name
    suffix = 2
    while f"{name}_{suffix}" in used:
        suffix += 1
    unique_name = f"{name}_{suffix}"
    used.add(unique_name)
    return unique_name


def collect_worksheet_columns(wks: Any) -> list[dict[str, Any]]:
    """Return column metadata, including designations when available."""
    columns: list[dict[str, Any]] = []
    total_cols = int(getattr(wks, "cols", 0) or 0)
    for index in range(min(total_cols, MAX_PREVIEW_COLS)):
        info = get_column_info(wks, index)
        clipped = [key for key, value in info.items() if isinstance(value, str) and len(value) > 256]
        item = {key: value[:256] if isinstance(value, str) else value for key, value in info.items()}
        if clipped:
            item["truncated_fields"] = clipped
        columns.append(item)
    with suppress(Exception):
        designations = [wks.get_label(index, "D") for index in range(len(columns))]
        for index, designation in enumerate(designations):
            if index < len(columns):
                columns[index]["designation"] = str(designation)[:256]
                if len(str(designation)) > 256:
                    columns[index].setdefault("truncated_fields", []).append("designation")
    return columns


# Bounds apply before any cell data crosses COM. Text limits bound the response;
# Origin still returns complete cell strings within this bounded rectangle.
MAX_PREVIEW_ROWS = 100
MAX_PREVIEW_COLS = 32
MAX_PREVIEW_CELL_CHARS = 256
MAX_PREVIEW_TEXT_CHARS = 20_000


def collect_worksheet_preview(
    wks: Any,
    max_rows: int = DEFAULT_MAX_PREVIEW_ROWS,
    max_cols: int = MAX_PREVIEW_COLS,
) -> dict[str, Any]:
    """Read only a bounded rectangle; total_rows remains the used row extent."""
    if max_rows < 0 or max_cols < 0:
        raise ValueError("max_rows 和 max_cols 不能小于 0")

    total_cols = int(wks.cols)
    # WSheet.rows is wks.maxRows (used extent), not allocated shape[0].
    total_rows = int(wks.rows)
    allocated_rows = int(wks.shape[0])
    rows_to_return = min(total_rows, max_rows, MAX_PREVIEW_ROWS)
    cols_to_return = min(total_cols, max_cols, MAX_PREVIEW_COLS)
    used_names: set[str] = set()
    col_names: list[str] = []
    text_remaining = MAX_PREVIEW_TEXT_CHARS
    shortened = 0

    def bounded_text(value: Any, limit: int = MAX_PREVIEW_CELL_CHARS) -> Any:
        nonlocal text_remaining, shortened
        if not isinstance(value, str):
            return value
        budget = min(limit, text_remaining)
        text_remaining -= min(len(value), budget)
        if len(value) <= budget:
            return value
        shortened += 1
        return value[:max(0, budget - 1)] + ("…" if budget else "")

    for index in range(cols_to_return):
        # Do not hide COM errors as empty labels or empty data.
        name = wks.get_label(index, "L") or wks.get_label(index, "G") or f"Col{index + 1}"
        col_names.append(_unique_column_name(bounded_text(str(name), 128), used_names))

    # -1 means 'to the end' in to_list2: never issue a read for zero rows/cols.
    block = wks.to_list2(0, rows_to_return - 1, 0, cols_to_return - 1) if rows_to_return and cols_to_return else []
    if not cols_to_return:
        rows_to_return = 0
    if len(block) != rows_to_return or any(len(row) != cols_to_return for row in block):
        raise RuntimeError("Origin 预览返回的区域尺寸与请求不符")
    records = [
        {name: bounded_text(row[index]) for index, name in enumerate(col_names)}
        for row in block
    ]
    omitted_rows = total_rows - rows_to_return
    omitted_cols = total_cols - cols_to_return
    omissions = []
    if omitted_rows:
        omissions.append(f"省略 {omitted_rows} 行")
    if omitted_cols:
        omissions.append(f"省略 {omitted_cols} 列")
    if shortened:
        omissions.append(f"截短 {shortened} 个文本值或列名")
    return {
        "total_rows": total_rows,
        "row_count_semantics": "used_extent_including_gaps",
        "allocated_rows": allocated_rows,
        "total_cols": total_cols,
        "returned_rows": rows_to_return,
        "returned_cols": cols_to_return,
        "omitted_rows": omitted_rows,
        "omitted_cols": omitted_cols,
        "shortened_text_values": shortened,
        "omissions": omissions,
        "truncated": bool(omissions),
        "columns": col_names,
        "data": records,
    }


def read_worksheet_detail(
    op: Any,
    sheet_name: str,
    *,
    include_preview: bool = True,
    max_preview_rows: int = DEFAULT_MAX_PREVIEW_ROWS,
) -> dict[str, Any]:
    """Read worksheet structure and an optional data preview."""
    wks = find_worksheet(op, sheet_name)
    detail: dict[str, Any] = {
        "sheet_name": sheet_name,
        "rows": getattr(wks, "rows", None),
        "cols": getattr(wks, "cols", None),
        "columns": collect_worksheet_columns(wks),
        "omitted_columns": max(0, int(wks.cols) - MAX_PREVIEW_COLS),
        "column_text_limit": 256,
    }
    if include_preview:
        detail["preview"] = collect_worksheet_preview(wks, max_preview_rows)
    return detail


def collect_graph_layers(graph: Any) -> list[dict[str, Any]]:
    """Collect per-layer curve summaries for a graph page."""
    layers: list[dict[str, Any]] = []
    layer_count = _optional_len(graph) or 0
    for layer_index in range(layer_count):
        layer = graph[layer_index]
        plots: list[dict[str, Any]] = []
        with suppress(Exception):
            plot_items = layer.plot_list()
            for plot_index, plot in enumerate(plot_items):
                plot_info: dict[str, Any] = {"index": plot_index}
                with suppress(Exception):
                    if hasattr(plot, "lt_range"):
                        plot_info["range"] = plot.lt_range()
                with suppress(Exception):
                    red, green, blue = plot.color
                    plot_info["color"] = f"#{red:02x}{green:02x}{blue:02x}"
                plots.append(plot_info)
        if not plots:
            with suppress(Exception):
                plot_count = get_plot_count(layer)
                plots = [{"index": index} for index in range(plot_count)]
        layers.append(
            {
                "index": layer_index,
                "plot_count": len(plots),
                "plots": plots,
            }
        )
    return layers


def read_graph_detail(op: Any, graph_name: str) -> dict[str, Any]:
    """Read graph layers and curve list."""
    graph = find_graph(op, graph_name)
    layers = collect_graph_layers(graph)
    return {
        "graph_name": graph_name,
        "long_name": getattr(graph, "lname", None),
        "layer_count": len(layers),
        "layers": layers,
    }


def build_session_snapshot(
    op: Any,
    *,
    active_worksheet: str | None = None,
    active_graph: str | None = None,
    include_preview: bool = False,
    max_preview_rows: int = DEFAULT_MAX_PREVIEW_ROWS,
) -> dict[str, Any]:
    """Build a read-only snapshot of the current Origin session."""
    worksheets = collect_worksheets(op)
    graphs = collect_graphs(op)
    matrices = collect_matrix_sheets(op)
    notes = collect_named_pages(op, PAGE_KIND_NOTES)
    excel_books = collect_named_pages(op, PAGE_KIND_EXCEL)
    snapshot: dict[str, Any] = {
        "exe_path": _safe_path(op, "e"),
        "user_path": _safe_path(op, "u"),
        "project": collect_project_meta(op),
        "active_worksheet": active_worksheet,
        "active_graph": active_graph,
        "counts": {
            "worksheets": len(worksheets),
            "graphs": len(graphs),
            "matrices": len(matrices),
            "notes": len(notes),
            "excel_books": len(excel_books),
        },
        "worksheets": worksheets,
        "graphs": graphs,
        "matrices": matrices,
        "notes": notes,
        "excel_books": excel_books,
    }
    if include_preview and active_worksheet:
        snapshot["active_worksheet_preview"] = read_worksheet_detail(
            op,
            active_worksheet,
            include_preview=True,
            max_preview_rows=max_preview_rows,
        )
    return snapshot
