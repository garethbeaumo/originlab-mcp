"""Test-only stdio host: production registrations with a controlled backend.

Never imported by the application. No Origin installation or COM is used.
"""

import threading

from originlab_mcp import server
from tests.fakes import FakeOrigin
from tests.fakes.origin import FakeGraph, FakeLayer, FakeLinearFit, FakeNLFit


class FailureLinearFit(FakeLinearFit):
    def result(self):
        if self._wks.name == "ResultFailure":
            raise RuntimeError("controlled fit result failure")
        return super().result()

    def report(self, band=0):
        if self._wks.name == "ReportFailure":
            raise RuntimeError("controlled fit report failure")
        return super().report(band)


class FailureNLFit(FakeNLFit):
    def result(self):
        if self._wks.name == "ResultFailure":
            raise RuntimeError("controlled fit result failure")
        return super().result()

    def report(self):
        if self._wks.name == "ReportFailure":
            raise RuntimeError("controlled fit report failure")
        return super().report()


class FailureLayer(FakeLayer):
    def add_plot(self, wks, **kwargs):
        if wks.name == "PlotPartial" and kwargs["coly"] == 2:
            raise RuntimeError("controlled second curve failure")
        return super().add_plot(wks, **kwargs)


class FailureGraph(FakeGraph):
    def save_fig(self, path, **kwargs):
        return ""


class ProtocolOrigin(FakeOrigin):
    def __init__(self):
        super().__init__()
        for name in ("FitData", "ResultFailure", "ReportFailure", "PlotPartial"):
            sheet = self.new_sheet(lname=name)
            for index, values in enumerate(([1, 2, 3], [3, 5, 7], [4, 6, 8])):
                sheet.from_list(index, values)
        self.graphs.append(FailureGraph("ExportFailure"))

    def LinearFit(self):
        return FailureLinearFit()

    def NLFit(self, function_name):
        return FailureNLFit(function_name)

    def new_graph(self, template="line"):
        graph = super().new_graph(template)
        graph._layers = [FailureLayer()]
        return graph

    def pages(self, kind: str):
        if kind == "w":
            raise RuntimeError("controlled backend failure")
        return super().pages(kind)

    def lt_exec(self, command: str):
        # A finite blocked call exercises the real manager's soft deadline.
        threading.Event().wait(0.2)
        return None


if __name__ == "__main__":
    server._manager._op = ProtocolOrigin()
    server._manager._connected = True
    server._manager._idle_timeout = 0
    server.main()
