"""In-memory OriginPro stand-in for Linux/Windows unit contract tests.

Mirrors the Excel MCP pattern of a mock desktop backend so multi-tool
workflows can be exercised without a live COM session.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class FakeAxis:
    title: str = ""


@dataclass
class FakePlot:
    color: str | None = None
    symbol_size: float | None = None
    shapelist: Any = None
    colormap: str | None = None
    transparency: float | None = None
    source: dict[str, Any] = field(default_factory=dict)


class FakeLayer:
    def __init__(self) -> None:
        self._plots: list[FakePlot] = []
        self._axes = {"x": FakeAxis(), "y": FakeAxis(), "z": FakeAxis()}
        self.xscale: str | None = None
        self.yscale: str | None = None
        self.rescaled = False
        self.grouped = False
        self.removed_labels: list[str] = []

    def add_plot(self, wks: Any, **kwargs: Any) -> FakePlot:
        plot = FakePlot(source={"wks": getattr(wks, "name", None), **kwargs})
        self._plots.append(plot)
        return plot

    def rescale(self) -> None:
        self.rescaled = True

    def group(self, *args: Any) -> None:
        self.grouped = True

    def plot_list(self) -> list[FakePlot]:
        return list(self._plots)

    def plot(self, index: int) -> FakePlot | None:
        if 0 <= index < len(self._plots):
            return self._plots[index]
        return None

    def remove_plot(self, index: int) -> None:
        if 0 <= index < len(self._plots):
            self._plots.pop(index)

    @property
    def num_plots(self) -> int:
        return len(self._plots)

    def set_xlim(self, begin: float = 0.0, end: float = 1.0, step: float = 0.0) -> None:
        self._xlim = (begin, end, step)

    def set_ylim(self, begin: float = 0.0, end: float = 1.0, step: float = 0.0) -> None:
        self._ylim = (begin, end, step)

    def axis(self, name: str) -> FakeAxis:
        return self._axes[name.lower()]

    def remove_label(self, label_name: str) -> None:
        self.removed_labels.append(label_name)


class FakeGraph:
    def __init__(self, name: str, template: str = "line") -> None:
        self.name = name
        self.lname = name
        self.template = template
        layer_count = 2 if template.lower() in {"doubley", "doubleY"} else 1
        self._layers = [FakeLayer() for _ in range(layer_count)]
        self.saved: list[tuple[str, dict[str, Any]]] = []
        self.copied: list[dict[str, Any]] = []
        self.added_layer_types: list[int] = []

    def __getitem__(self, index: int) -> FakeLayer:
        return self._layers[index]

    def __len__(self) -> int:
        return len(self._layers)

    def save_fig(self, path: str, **kwargs: Any) -> str:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        Path(path).write_bytes(b"fake-origin-export")
        self.saved.append((path, kwargs))
        return path

    def add_layer(self, layer_type: int = 0) -> FakeLayer:
        layer = FakeLayer()
        self._layers.append(layer)
        self.added_layer_types.append(layer_type)
        return layer

    def copy_page(self, fmt: str, dpi: int, quality: int, transparent: bool) -> None:
        self.copied.append(
            {"fmt": fmt, "dpi": dpi, "quality": quality, "transparent": transparent}
        )


class FakeLinearFit:
    """Minimal originpro.LinearFit stand-in with deterministic results."""

    def __init__(self) -> None:
        self._wks: FakeSheet | None = None
        self._x_col = 0
        self._y_col = 0
        self._yerr_col: int | None = None
        self.fix_slope: float | None = None
        self.fix_intercept: float | None = None
        self.calls: list[str] = []

    def set_data(
        self,
        wks: FakeSheet,
        x_col: int,
        y_col: int,
        yerr_col: int | None = None,
    ) -> None:
        self.calls.append("set_data")
        self._wks = wks
        self._x_col = x_col
        self._y_col = y_col
        self._yerr_col = yerr_col

    def _fit_params(self) -> dict[str, Any]:
        if self._wks is None:
            return {
                "Parameters": {
                    "Slope": {"Value": 1.0, "Error": 0.01},
                    "Intercept": {"Value": 0.0, "Error": 0.01},
                },
                "Statistics": {"RSqCOD": 1.0},
            }

        xs = [float(v) for v in self._wks.to_list(self._x_col)]
        ys = [float(v) for v in self._wks.to_list(self._y_col)]
        n = min(len(xs), len(ys))
        if n == 0:
            slope = 0.0
            intercept = 0.0
        elif self.fix_slope is not None and self.fix_intercept is not None:
            slope = float(self.fix_slope)
            intercept = float(self.fix_intercept)
        elif self.fix_slope is not None:
            slope = float(self.fix_slope)
            intercept = sum(ys[i] - slope * xs[i] for i in range(n)) / n
        elif self.fix_intercept is not None:
            intercept = float(self.fix_intercept)
            denom = sum(xs[i] * xs[i] for i in range(n)) or 1.0
            slope = sum((ys[i] - intercept) * xs[i] for i in range(n)) / denom
        else:
            mean_x = sum(xs[:n]) / n
            mean_y = sum(ys[:n]) / n
            denom = sum((xs[i] - mean_x) ** 2 for i in range(n)) or 1.0
            slope = sum((xs[i] - mean_x) * (ys[i] - mean_y) for i in range(n)) / denom
            intercept = mean_y - slope * mean_x

        return {
            "Parameters": {
                "Slope": {"Value": slope, "Error": 0.01},
                "Intercept": {"Value": intercept, "Error": 0.01},
            },
            "Statistics": {"RSqCOD": 0.99, "N": n},
        }

    def result(self) -> dict[str, Any]:
        self.calls.append("result")
        return self._fit_params()

    def report(self, band: int = 0) -> tuple[str, str]:
        self.calls.append(f"report:{band}")
        return ("[Book1]FitReport", "[Book1]FitCurves")


class FakeNLFit:
    """Minimal originpro.NLFit stand-in."""

    SUPPORTED = {
        "Gauss",
        "Lorentz",
        "ExpDec1",
        "Boltzmann",
        "Allometric1",
        "PolyLine",
    }

    def __init__(self, function_name: str) -> None:
        if function_name not in self.SUPPORTED:
            raise ValueError(f"unknown fit function: {function_name}")
        self.function_name = function_name
        self._params: dict[str, float] = {}
        self._fixed: dict[str, float] = {}
        self._fitted = False
        self.calls: list[str] = []

    def set_data(
        self,
        wks: FakeSheet,
        x_col: int,
        y_col: int,
        yerr: int | None = None,
    ) -> None:
        self.calls.append("set_data")
        self._wks = wks
        self._x_col = x_col
        self._y_col = y_col
        self._yerr = yerr

    def set_param(self, name: str, value: float) -> None:
        self.calls.append(f"set_param:{name}")
        self._params[name] = float(value)

    def fix_param(self, name: str, value: Any) -> None:
        self.calls.append(f"fix_param:{name}")
        if value is False:
            self._fixed.pop(name, None)
        else:
            self._fixed[name] = float(value)

    def fit(self) -> None:
        self.calls.append("fit")
        self._fitted = True
        if "xc" not in self._params:
            self._params["xc"] = 0.0
        if "w" not in self._params:
            self._params["w"] = 1.0
        if "A" not in self._params:
            self._params["A"] = 1.0
        self._params.update(self._fixed)

    def result(self) -> dict[str, Any]:
        self.calls.append("result")
        if not self._fitted:
            self.fit()
        return {
            "Parameters": {
                name: {"Value": value, "Error": 0.02}
                for name, value in self._params.items()
            },
            "Statistics": {"RSqCOD": 0.98},
        }

    def report(self) -> tuple[str, str]:
        self.calls.append("report")
        return ("[Book1]NLFitReport", "[Book1]NLFitCurves")


class FakeBook:
    def __init__(self, name: str) -> None:
        self.name = name
        self.lname = name
        self._sheets: list[FakeSheet] = []

    def add_sheet(self, name: str = "Sheet1") -> FakeSheet:
        sheet = FakeSheet(name=name, book=self)
        self._sheets.append(sheet)
        return sheet

    def __iter__(self):
        return iter(self._sheets)


class FakeSheet:
    def __init__(self, name: str, book: FakeBook) -> None:
        self.name = name
        self._book = book
        self._columns: list[list[Any]] = []
        self._labels: dict[str, list[str]] = {
            "G": [],
            "L": [],
            "U": [],
            "C": [],
            "D": [],
        }
        self.formulas: dict[int, str] = {}
        self.sort_calls: list[tuple[int, bool]] = []
        self.cleared: list[tuple[int, int | None]] = []
        self.axis_specs: list[str] = []

    @property
    def cols(self) -> int:
        return len(self._columns)

    @property
    def rows(self) -> int:
        if not self._columns:
            return 0
        return max(len(col) for col in self._columns)

    def get_book(self) -> FakeBook:
        return self._book

    def _ensure_col(self, index: int) -> None:
        while len(self._columns) <= index:
            self._columns.append([])
            for key in self._labels:
                self._labels[key].append("")
            short = chr(ord("A") + len(self._columns) - 1)
            self._labels["G"][-1] = short

    def from_list(self, col_index: int, data: list, lname: str | None = None) -> None:
        self._ensure_col(col_index)
        self._columns[col_index] = list(data)
        if lname is not None:
            self._labels["L"][col_index] = str(lname)

    def from_file(self, path: str) -> None:
        with open(path, newline="", encoding="utf-8") as handle:
            rows = list(csv.reader(handle))
        if not rows:
            return
        header = rows[0]
        body = rows[1:] if len(rows) > 1 else []
        width = max(len(header), max((len(r) for r in body), default=0))
        for col_idx in range(width):
            values = [r[col_idx] if col_idx < len(r) else "" for r in body]
            parsed: list[Any] = []
            for value in values:
                try:
                    number = float(value)
                    parsed.append(int(number) if number == int(number) else number)
                except (TypeError, ValueError):
                    parsed.append(value)
            lname = header[col_idx] if col_idx < len(header) else None
            self.from_list(col_idx, parsed, lname=lname)

    def to_list(self, col_index: int) -> list:
        self._ensure_col(col_index)
        return list(self._columns[col_index])

    def cols_axis(self, spec: str) -> None:
        self.axis_specs.append(spec)
        # Origin uses 1-based designations in a compact string; keep a simple map.
        mapping = {"X": "X", "Y": "Y", "Z": "Z", "E": "E", "yE": "E", "xE": "xE"}
        # Spec examples like "XYY" or "XYE"
        for idx, ch in enumerate(spec):
            self._ensure_col(idx)
            self._labels["D"][idx] = mapping.get(ch, ch)

    def get_label(self, col_index: int, label_type: str = "L") -> str:
        self._ensure_col(col_index)
        return self._labels.get(label_type.upper(), [""] * self.cols)[col_index]

    def get_labels(self, label_type: str) -> list[str]:
        key = label_type.upper()
        while len(self._labels[key]) < self.cols:
            self._labels[key].append("")
        return list(self._labels[key][: self.cols])

    def set_label(self, col: int, value: str, label_type: str) -> None:
        self._ensure_col(col)
        self._labels[label_type.upper()][col] = value

    def sort(self, col: int, descending: bool) -> None:
        self.sort_calls.append((col, descending))

    def clear(self, c1: int = 0, c2: int | None = None) -> None:
        self.cleared.append((c1, c2))
        end = self.cols if c2 is None else c2
        for idx in range(c1, end):
            if idx < len(self._columns):
                self._columns[idx] = []

    def set_formula(self, col: int | str, formula: str) -> None:
        index = int(col) if not isinstance(col, int) else col
        self.formulas[index] = formula

    def cell(self, row: int, col: int | str) -> Any:
        index = int(col) if not isinstance(col, int) else col
        self._ensure_col(index)
        if row < 0 or row >= len(self._columns[index]):
            return None
        return self._columns[index][row]

    def del_col(self, col: int | str, count: int = 1) -> None:
        index = int(col) if not isinstance(col, int) else col
        for _ in range(count):
            if 0 <= index < len(self._columns):
                self._columns.pop(index)
                for key in self._labels:
                    if index < len(self._labels[key]):
                        self._labels[key].pop(index)


class FakeOrigin:
    """Stateful fake `originpro` module used by OriginManager.execute()."""

    def __init__(self) -> None:
        self.books: list[FakeBook] = []
        self.graphs: list[FakeGraph] = []
        self.matrices: list[Any] = []
        self.notes: list[Any] = []
        self.excel_books: list[Any] = []
        self._sheet_counter = 0
        self._graph_counter = 0
        self._book_counter = 0
        self.lt_commands: list[str] = []
        self.lt_strings: dict[str, str] = {}
        self.lt_floats: dict[str, float] = {}
        self.saved_paths: list[str] = []
        self.opened_paths: list[tuple[str, bool]] = []
        self.show_calls: list[bool] = []
        self.detached = False
        self.exited = False
        self.attach_calls = 0
        self.exe_path = r"C:\Program Files\OriginLab\Origin\Origin.exe"
        self.user_path = r"C:\Users\Test\Documents\OriginLab"
        self.project_path: str | None = None
        self.linear_fits: list[FakeLinearFit] = []
        self.nl_fits: list[FakeNLFit] = []

    def LinearFit(self) -> FakeLinearFit:
        fit = FakeLinearFit()
        self.linear_fits.append(fit)
        return fit

    def NLFit(self, function_name: str) -> FakeNLFit:
        fit = FakeNLFit(function_name)
        self.nl_fits.append(fit)
        return fit

    def _next_book_name(self) -> str:
        self._book_counter += 1
        return f"Book{self._book_counter}"

    def _next_sheet_name(self) -> str:
        self._sheet_counter += 1
        return f"Sheet{self._sheet_counter}"

    def _next_graph_name(self) -> str:
        self._graph_counter += 1
        return f"Graph{self._graph_counter}"

    def new_sheet(self, **kwargs: Any) -> FakeSheet:
        book = FakeBook(self._next_book_name())
        sheet_name = kwargs.get("lname") or kwargs.get("name") or self._next_sheet_name()
        sheet = book.add_sheet(str(sheet_name))
        self.books.append(book)
        return sheet

    def new_graph(self, template: str = "line") -> FakeGraph:
        graph = FakeGraph(self._next_graph_name(), template=template)
        self.graphs.append(graph)
        return graph

    def find_sheet(self, kind: str, name: str) -> FakeSheet | None:
        if kind not in {"w", "W", "worksheet", "Worksheet"}:
            return None
        needle = (name or "").strip()
        for book in self.books:
            for sheet in book:
                full = f"[{book.name}]{sheet.name}"
                if needle in {sheet.name, full, book.name}:
                    return sheet
        return None

    def find_graph(self, name: str) -> FakeGraph | None:
        for graph in self.graphs:
            if graph.name == name or graph.lname == name:
                return graph
        return None

    def pages(self, kind: str):
        mapping = {
            "Book": self.books,
            "Graph": self.graphs,
            "Matrix": self.matrices,
            "Notes": self.notes,
            "Excel": self.excel_books,
        }
        return list(mapping.get(kind, []))

    def save(self, path: str = "") -> None:
        if path:
            self.project_path = path
            self.saved_paths.append(path)
        elif self.project_path:
            self.saved_paths.append(self.project_path)
        else:
            self.saved_paths.append("")

    def open(self, file: str = "", readonly: bool = False) -> None:
        self.opened_paths.append((file, readonly))
        self.project_path = file or None
        # Simulate replacing the live project.
        self.books.clear()
        self.graphs.clear()
        book = FakeBook("Book1")
        book.add_sheet("Sheet1")
        self.books.append(book)

    def new(self) -> None:
        self.books.clear()
        self.graphs.clear()
        self.project_path = None
        self._sheet_counter = 0
        self._graph_counter = 0
        self._book_counter = 0
        book = FakeBook(self._next_book_name())
        book.add_sheet(self._next_sheet_name())
        self.books.append(book)

    def lt_exec(self, command: str) -> Any:
        self.lt_commands.append(command)
        return None

    def get_lt_str(self, name: str) -> str:
        return self.lt_strings.get(name, "")

    def lt_float(self, name: str) -> float:
        return float(self.lt_floats.get(name, 0.0))

    def path(self, kind: str) -> str:
        if kind == "e":
            return self.exe_path
        if kind == "u":
            return self.user_path
        return ""

    def set_show(self, visible: bool) -> None:
        self.show_calls.append(visible)

    def attach(self) -> None:
        self.detached = False
        self.attach_calls = getattr(self, "attach_calls", 0) + 1

    def detach(self) -> None:
        self.detached = True

    def exit(self) -> None:
        self.exited = True
        self.detached = True
