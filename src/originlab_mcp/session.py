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
    get_column_label,
    get_plot_count,
)

PAGE_KIND_BOOK = "Book"
PAGE_KIND_GRAPH = "Graph"
PAGE_KIND_MATRIX = "Matrix"
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


def _column_display_name(wks: Any, index: int) -> str:
    long_name = get_column_label(wks, index, "L")
    if long_name:
        return long_name
    short_name = get_column_label(wks, index, "G")
    if short_name:
        return short_name
    return f"Col{index + 1}"


def collect_worksheet_columns(wks: Any) -> list[dict[str, Any]]:
    """Return column metadata, including designations when available."""
    columns: list[dict[str, Any]] = []
    total_cols = int(getattr(wks, "cols", 0) or 0)
    for index in range(total_cols):
        columns.append(get_column_info(wks, index))
    with suppress(Exception):
        designations = wks.get_labels("D")
        for index, designation in enumerate(designations):
            if index < len(columns):
                columns[index]["designation"] = designation
    return columns


def collect_worksheet_preview(
    wks: Any,
    max_rows: int = DEFAULT_MAX_PREVIEW_ROWS,
) -> dict[str, Any]:
    """Return a truncated row preview for a worksheet."""
    if max_rows < 0:
        raise ValueError("max_rows 不能小于 0")

    total_cols = int(getattr(wks, "cols", 0) or 0)
    used_names: set[str] = set()
    col_names: list[str] = []
    col_values: list[list[Any]] = []
    for index in range(total_cols):
        col_names.append(_unique_column_name(_column_display_name(wks, index), used_names))
        try:
            values = list(wks.to_list(index))
        except Exception:
            values = []
        col_values.append(values)

    total_rows = max((len(values) for values in col_values), default=0)
    rows_to_return = min(total_rows, max_rows)
    records = []
    for row_index in range(rows_to_return):
        records.append(
            {
                col_names[col_index]: (
                    col_values[col_index][row_index]
                    if row_index < len(col_values[col_index])
                    else None
                )
                for col_index in range(total_cols)
            }
        )
    return {
        "total_rows": total_rows,
        "returned_rows": rows_to_return,
        "truncated": total_rows > max_rows,
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
        with suppress(Exception):
            snapshot["active_worksheet_preview"] = read_worksheet_detail(
                op,
                active_worksheet,
                include_preview=True,
                max_preview_rows=max_preview_rows,
            )
    return snapshot
