"""Local control panel for OriginLab MCP diagnostics and configuration."""

from __future__ import annotations

import json
import os
import shutil
import signal
import subprocess
import sys
import threading
import time
import webbrowser
from collections import deque
from dataclasses import asdict, dataclass
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

HOST = "127.0.0.1"
PORT = 8765
MAX_LOG_LINES = 300
SERVER_NAME = "originlab"


@dataclass(frozen=True)
class ClientDefinition:
    id: str
    name: str
    path: Path
    scope: str
    hint: str


@dataclass
class ClientConfigStatus:
    id: str
    name: str
    path: str
    scope: str
    exists: bool
    configured: bool
    hint: str


class UIState:
    def __init__(self) -> None:
        self.lock = threading.RLock()
        self.process: subprocess.Popen[str] | None = None
        self.logs: deque[str] = deque(maxlen=MAX_LOG_LINES)
        self.origin_status: dict[str, Any] = {
            "state": "not_tested",
            "message": "Origin connection has not been tested yet.",
        }

    def log(self, message: str) -> None:
        stamp = time.strftime("%H:%M:%S")
        with self.lock:
            self.logs.append(f"[{stamp}] {message}")


STATE = UIState()


def _creationflags() -> int:
    if os.name != "nt":
        return 0
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="ignore")


def _read_json_file(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    data = json.loads(_read_text(path))
    if not isinstance(data, dict):
        raise ValueError("The config root must be a JSON object.")
    return data


def _server_config() -> dict[str, Any]:
    return {
        "command": sys.executable,
        "args": ["-m", "originlab_mcp.server"],
    }


def _client_definitions() -> list[ClientDefinition]:
    root = _repo_root()
    appdata = Path(os.environ.get("APPDATA", ""))
    return [
        ClientDefinition(
            id="antigravity",
            name="Antigravity / Gemini",
            path=root / ".gemini" / "settings.json",
            scope="Project",
            hint="Writes .gemini/settings.json.",
        ),
        ClientDefinition(
            id="cursor",
            name="Cursor",
            path=root / ".cursor" / "mcp.json",
            scope="Project",
            hint="Writes .cursor/mcp.json.",
        ),
        ClientDefinition(
            id="codex",
            name="Codex",
            path=root / ".codex" / "config.json",
            scope="Project",
            hint="Writes .codex/config.json.",
        ),
        ClientDefinition(
            id="trae",
            name="Trae",
            path=root / ".trae" / "mcp.json",
            scope="Project",
            hint="Writes .trae/mcp.json.",
        ),
        ClientDefinition(
            id="claude",
            name="Claude Desktop",
            path=appdata / "Claude" / "claude_desktop_config.json",
            scope="User",
            hint="Writes the Claude Desktop user config.",
        ),
    ]


def _client_definition(client_id: str) -> ClientDefinition | None:
    for definition in _client_definitions():
        if definition.id == client_id:
            return definition
    return None


def _is_originlab_configured(path: Path) -> bool:
    if not path.is_file():
        return False
    try:
        data = _read_json_file(path)
    except (json.JSONDecodeError, ValueError):
        return "originlab" in _read_text(path).lower()

    servers = data.get("mcpServers")
    return isinstance(servers, dict) and SERVER_NAME in servers


def detect_client_configs() -> list[ClientConfigStatus]:
    statuses = []
    for definition in _client_definitions():
        exists = definition.path.is_file()
        configured = _is_originlab_configured(definition.path)
        statuses.append(
            ClientConfigStatus(
                id=definition.id,
                name=definition.name,
                path=str(definition.path),
                scope=definition.scope,
                exists=exists,
                configured=configured,
                hint="Configured" if configured else definition.hint,
            )
        )
    return statuses


def configure_client(client_id: str) -> dict[str, Any]:
    definition = _client_definition(client_id)
    if definition is None:
        return {
            "ok": False,
            "message": f"Unknown client: {client_id}",
            "client_id": client_id,
        }

    try:
        data = _read_json_file(definition.path)
    except json.JSONDecodeError as exc:
        return {
            "ok": False,
            "message": f"Config is not valid JSON: {exc}",
            "path": str(definition.path),
        }
    except ValueError as exc:
        return {
            "ok": False,
            "message": str(exc),
            "path": str(definition.path),
        }

    servers = data.setdefault("mcpServers", {})
    if not isinstance(servers, dict):
        return {
            "ok": False,
            "message": "The mcpServers field must be a JSON object.",
            "path": str(definition.path),
        }

    existed = definition.path.is_file()
    previous = servers.get(SERVER_NAME)
    servers[SERVER_NAME] = _server_config()

    new_text = json.dumps(data, indent=2, ensure_ascii=False) + "\n"
    old_text = _read_text(definition.path) if existed else None
    backup_path = None
    changed = old_text != new_text

    if changed:
        definition.path.parent.mkdir(parents=True, exist_ok=True)
        if existed:
            stamp = time.strftime("%Y%m%d-%H%M%S")
            backup_path = definition.path.with_name(
                f"{definition.path.name}.bak-{stamp}",
            )
            shutil.copy2(definition.path, backup_path)
        definition.path.write_text(new_text, encoding="utf-8")

    STATE.log(
        f"{'Updated' if existed else 'Created'} {definition.name} config: "
        f"{definition.path}"
    )
    return {
        "ok": True,
        "message": f"{definition.name} configuration is ready.",
        "client": {
            "id": definition.id,
            "name": definition.name,
            "path": str(definition.path),
            "scope": definition.scope,
            "hint": definition.hint,
        },
        "path": str(definition.path),
        "created": not existed,
        "changed": changed,
        "updated_existing": previous is not None,
        "backup_path": str(backup_path) if backup_path else None,
        "server": _server_config(),
    }


def _origin_process_running() -> bool:
    if os.name != "nt":
        return False
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Origin64.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
            creationflags=_creationflags(),
        )
    except Exception:
        return False
    return "Origin64.exe" in result.stdout


def _stream_process_output(process: subprocess.Popen[str]) -> None:
    assert process.stdout is not None
    for line in process.stdout:
        STATE.log(line.rstrip())


def _mcp_command() -> list[str]:
    return [sys.executable, "-m", "originlab_mcp.server"]


def mcp_status() -> dict[str, Any]:
    with STATE.lock:
        process = STATE.process
        logs = list(STATE.logs)

    if process is None:
        return {"state": "stopped", "pid": None, "logs": logs}

    code = process.poll()
    if code is None:
        return {"state": "running", "pid": process.pid, "logs": logs}

    return {"state": "exited", "pid": process.pid, "exit_code": code, "logs": logs}


def start_mcp_server() -> dict[str, Any]:
    with STATE.lock:
        process = STATE.process
        if process is not None and process.poll() is None:
            return {"started": False, "message": "MCP server is already running."}

        command = _mcp_command()
        STATE.log("Starting MCP server debug process...")
        process = subprocess.Popen(
            command,
            cwd=str(_repo_root()),
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            creationflags=_creationflags(),
        )
        STATE.process = process

    threading.Thread(
        target=_stream_process_output,
        args=(process,),
        daemon=True,
    ).start()
    return {"started": True, "message": "MCP server started.", "pid": process.pid}


def stop_mcp_server() -> dict[str, Any]:
    with STATE.lock:
        process = STATE.process
        if process is None or process.poll() is not None:
            STATE.process = None
            return {"stopped": False, "message": "MCP server is not running."}

        STATE.log("Stopping MCP server debug process...")
        if os.name == "nt":
            process.terminate()
        else:
            process.send_signal(signal.SIGTERM)

    try:
        process.wait(timeout=8)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)

    with STATE.lock:
        STATE.process = None
    return {"stopped": True, "message": "MCP server stopped."}


def test_origin_connection() -> dict[str, Any]:
    try:
        import originpro as op

        if hasattr(op, "attach"):
            op.attach()
        op.set_show(True)
        exe_path = op.path("e")
        user_path = op.path("u")
        if hasattr(op, "detach"):
            op.detach()
        status = {
            "state": "connected",
            "message": "Origin connection test succeeded.",
            "exe_path": exe_path,
            "user_path": user_path,
        }
    except Exception as exc:
        status = {
            "state": "failed",
            "message": f"Origin connection test failed: {exc}",
        }

    with STATE.lock:
        STATE.origin_status = status
    STATE.log(status["message"])
    return status


def status_payload() -> dict[str, Any]:
    with STATE.lock:
        origin_status = dict(STATE.origin_status)
    origin_status["process_running"] = _origin_process_running()
    return {
        "ui": {"state": "running", "host": HOST, "port": PORT},
        "mcp": mcp_status(),
        "origin": origin_status,
        "configs": [asdict(item) for item in detect_client_configs()],
        "server": _server_config(),
    }


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>OriginLab MCP 控制台</title>
  <style>
    :root {
      color-scheme: light;
      font-family: "Segoe UI", "Microsoft YaHei", Arial, sans-serif;
      --bg: #f4f5f2;
      --panel: #ffffff;
      --line: #d9ded6;
      --text: #1f2722;
      --muted: #66706a;
      --accent: #0f766e;
      --accent-dark: #0b5f59;
      --warn: #b7791f;
      --bad: #c24135;
      --ok: #26834f;
      --code: #111816;
    }
    * { box-sizing: border-box; }
    body { margin: 0; background: var(--bg); color: var(--text); }
    header {
      border-bottom: 1px solid var(--line);
      background: #fbfcfa;
      padding: 18px 24px;
    }
    .topbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      max-width: 1180px;
      margin: 0 auto;
    }
    h1 { margin: 0; font-size: 22px; font-weight: 650; letter-spacing: 0; }
    main { max-width: 1180px; margin: 0 auto; padding: 20px 24px 28px; }
    .chips { display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end; }
    .chip {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      min-height: 28px;
      border: 1px solid var(--line);
      background: #fff;
      border-radius: 6px;
      padding: 4px 9px;
      color: var(--muted);
      font-size: 13px;
      white-space: nowrap;
    }
    .dot { width: 9px; height: 9px; border-radius: 50%; background: #9aa39d; flex: none; }
    .ok { background: var(--ok); }
    .warn { background: var(--warn); }
    .bad { background: var(--bad); }
    .layout { display: grid; grid-template-columns: 320px 1fr; gap: 16px; align-items: start; }
    .stack { display: grid; gap: 16px; }
    .card {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    .card h2 { margin: 0 0 12px; font-size: 15px; font-weight: 650; }
    .status-line { display: flex; align-items: center; gap: 8px; min-height: 28px; font-weight: 600; }
    .detail { color: var(--muted); font-size: 13px; line-height: 1.45; margin-top: 8px; }
    .path {
      font-family: Consolas, "Cascadia Mono", monospace;
      font-size: 12px;
      color: #4d5a54;
      word-break: break-all;
    }
    .actions { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }
    button, select {
      min-height: 34px;
      border: 1px solid #a9b4ad;
      background: #fff;
      color: var(--text);
      border-radius: 6px;
      padding: 7px 11px;
      font: inherit;
      font-size: 14px;
    }
    button { cursor: pointer; }
    button.primary { background: var(--accent); color: white; border-color: var(--accent); }
    button.primary:hover { background: var(--accent-dark); }
    button.danger { color: #a7372d; border-color: #e2aaa5; }
    button:disabled, select:disabled { opacity: .58; cursor: wait; }
    select { width: 100%; }
    .config-grid { display: grid; grid-template-columns: 260px 1fr; gap: 14px; }
    .preview {
      background: #f8faf7;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      display: grid;
      gap: 8px;
    }
    .preview-row { display: grid; grid-template-columns: 86px 1fr; gap: 10px; }
    .preview-label { color: var(--muted); font-size: 13px; }
    .notice {
      border-left: 3px solid var(--accent);
      background: #eef8f5;
      padding: 9px 11px;
      color: #25544e;
      font-size: 13px;
      line-height: 1.45;
      min-height: 38px;
    }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { border-bottom: 1px solid #e6eae4; text-align: left; padding: 9px 6px; vertical-align: top; }
    th { color: var(--muted); font-weight: 600; }
    .badge {
      display: inline-flex;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 2px 7px;
      font-size: 12px;
      color: var(--muted);
      background: #fff;
      white-space: nowrap;
    }
    .badge.ok-text { color: #17633a; border-color: #a9d5bd; background: #f1fbf5; }
    .badge.warn-text { color: #8a5a13; border-color: #e5c992; background: #fff8e8; }
    pre {
      margin: 0;
      background: var(--code);
      color: #dbe7e1;
      padding: 13px;
      border-radius: 8px;
      max-height: 260px;
      overflow: auto;
      font-size: 12px;
      line-height: 1.45;
    }
    @media (max-width: 940px) {
      .topbar { align-items: flex-start; flex-direction: column; }
      .chips { justify-content: flex-start; }
      .layout, .config-grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <div class="topbar">
      <h1>OriginLab MCP 控制台</h1>
      <div class="chips">
        <span class="chip"><span id="chipMcpDot" class="dot"></span><span id="chipMcp">MCP</span></span>
        <span class="chip"><span id="chipOriginDot" class="dot"></span><span id="chipOrigin">Origin</span></span>
        <span class="chip"><span class="dot ok"></span><span>127.0.0.1:8765</span></span>
      </div>
    </div>
  </header>
  <main>
    <div class="layout">
      <div class="stack">
        <section class="card">
          <h2>MCP Server</h2>
          <div class="status-line"><span id="mcpDot" class="dot"></span><span id="mcpText">Loading...</span></div>
          <div id="mcpPid" class="detail path"></div>
          <div class="actions">
            <button id="startBtn" class="primary">启动</button>
            <button id="stopBtn" class="danger">终止</button>
            <button id="refreshBtn">刷新</button>
          </div>
        </section>
        <section class="card">
          <h2>Origin</h2>
          <div class="status-line"><span id="originDot" class="dot"></span><span id="originText">Loading...</span></div>
          <div id="originDetail" class="detail path"></div>
          <div class="actions">
            <button id="testOriginBtn" class="primary">测试连接</button>
          </div>
        </section>
        <section class="card">
          <h2>日志</h2>
          <pre id="logs">Loading...</pre>
        </section>
      </div>
      <div class="stack">
        <section class="card">
          <h2>一键配置</h2>
          <div class="config-grid">
            <div>
              <select id="clientSelect" aria-label="选择客户端"></select>
              <div class="actions">
                <button id="configureBtn" class="primary">写入配置</button>
              </div>
            </div>
            <div class="preview">
              <div class="preview-row"><span class="preview-label">客户端</span><span id="previewName">-</span></div>
              <div class="preview-row"><span class="preview-label">状态</span><span id="previewStatus">-</span></div>
              <div class="preview-row"><span class="preview-label">路径</span><span id="previewPath" class="path">-</span></div>
              <div class="preview-row"><span class="preview-label">命令</span><span id="previewCommand" class="path">-</span></div>
            </div>
          </div>
          <div id="configNotice" class="notice" style="margin-top: 12px;">选择客户端后写入配置。</div>
        </section>
        <section class="card">
          <h2>配置状态</h2>
          <table>
            <thead><tr><th>Client</th><th>Scope</th><th>Status</th><th>Path</th></tr></thead>
            <tbody id="configRows"></tbody>
          </table>
        </section>
      </div>
    </div>
  </main>
  <script>
    const byId = (id) => document.getElementById(id);
    let latestStatus = null;

    const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (ch) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[ch]));
    const dotClass = (state) => (
      state === "running" || state === "connected" ? "dot ok" :
      state === "failed" || state === "exited" ? "dot bad" : "dot warn"
    );
    const setBusy = (busy) => {
      for (const id of ["startBtn", "stopBtn", "refreshBtn", "testOriginBtn", "configureBtn", "clientSelect"]) {
        byId(id).disabled = busy;
      }
    };
    async function call(path, options = {}) {
      setBusy(true);
      try {
        const res = await fetch(path, options);
        const data = await res.json();
        if (!res.ok || data.ok === false) throw new Error(data.message || res.statusText);
        return data;
      } finally {
        setBusy(false);
      }
    }
    function selectedConfig() {
      if (!latestStatus) return null;
      return latestStatus.configs.find((cfg) => cfg.id === byId("clientSelect").value) || latestStatus.configs[0];
    }
    function commandPreview() {
      if (!latestStatus) return "-";
      const command = latestStatus.server.command;
      const args = latestStatus.server.args || [];
      return [command, ...args].join(" ");
    }
    function renderClientOptions(configs) {
      const current = byId("clientSelect").value || "cursor";
      byId("clientSelect").innerHTML = configs.map((cfg) => (
        `<option value="${escapeHtml(cfg.id)}">${escapeHtml(cfg.name)}</option>`
      )).join("");
      byId("clientSelect").value = configs.some((cfg) => cfg.id === current) ? current : configs[0]?.id;
    }
    function renderPreview() {
      const cfg = selectedConfig();
      if (!cfg) return;
      byId("previewName").textContent = cfg.name;
      byId("previewStatus").textContent = cfg.configured ? "已配置" : (cfg.exists ? "文件存在，未配置 originlab" : "未创建");
      byId("previewPath").textContent = cfg.path;
      byId("previewCommand").textContent = commandPreview();
    }
    function renderConfigs(configs) {
      byId("configRows").innerHTML = configs.map((cfg) => {
        const status = cfg.configured ? "已配置" : (cfg.exists ? "未配置" : "未创建");
        const cls = cfg.configured ? "badge ok-text" : "badge warn-text";
        return `<tr>
          <td>${escapeHtml(cfg.name)}</td>
          <td><span class="badge">${escapeHtml(cfg.scope)}</span></td>
          <td><span class="${cls}">${status}</span></td>
          <td class="path">${escapeHtml(cfg.path)}</td>
        </tr>`;
      }).join("");
    }
    async function refresh() {
      const data = await call("/api/status");
      latestStatus = data;

      byId("mcpDot").className = dotClass(data.mcp.state);
      byId("chipMcpDot").className = dotClass(data.mcp.state);
      byId("mcpText").textContent = data.mcp.state;
      byId("chipMcp").textContent = `MCP ${data.mcp.state}`;
      byId("mcpPid").textContent = data.mcp.pid ? `PID ${data.mcp.pid}` : "";

      byId("originDot").className = dotClass(data.origin.state);
      byId("chipOriginDot").className = dotClass(data.origin.state);
      byId("originText").textContent = data.origin.state + (data.origin.process_running ? " · Origin64.exe running" : "");
      byId("chipOrigin").textContent = data.origin.process_running ? "Origin running" : `Origin ${data.origin.state}`;
      byId("originDetail").textContent = data.origin.exe_path || data.origin.message || "";

      renderClientOptions(data.configs);
      renderConfigs(data.configs);
      renderPreview();
      byId("logs").textContent = (data.mcp.logs || []).join("\\n") || "No logs yet.";
    }
    byId("clientSelect").onchange = renderPreview;
    byId("startBtn").onclick = async () => { await call("/api/start", { method: "POST" }); await refresh(); };
    byId("stopBtn").onclick = async () => { await call("/api/stop", { method: "POST" }); await refresh(); };
    byId("refreshBtn").onclick = refresh;
    byId("testOriginBtn").onclick = async () => { await call("/api/test-origin", { method: "POST" }); await refresh(); };
    byId("configureBtn").onclick = async () => {
      const client = byId("clientSelect").value;
      const result = await call("/api/configure", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ client })
      });
      byId("configNotice").textContent = result.changed
        ? `${result.message} ${result.backup_path ? "已备份原文件。" : ""}`
        : `${result.message} 已是最新。`;
      await refresh();
    };
    refresh().catch((err) => { byId("configNotice").textContent = err.message; });
    setInterval(() => refresh().catch(() => {}), 4000);
  </script>
</body>
</html>
"""


class UIHandler(BaseHTTPRequestHandler):
    def _send_json(self, payload: dict[str, Any], status: int = 200) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_html(self) -> None:
        data = INDEX_HTML.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _read_json_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if length == 0:
            return {}
        raw = self.rfile.read(length).decode("utf-8")
        data = json.loads(raw)
        if not isinstance(data, dict):
            raise ValueError("Request body must be a JSON object.")
        return data

    def do_GET(self) -> None:
        if self.path == "/" or self.path.startswith("/?"):
            self._send_html()
            return
        if self.path == "/api/status":
            self._send_json(status_payload())
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        if self.path == "/api/start":
            self._send_json(start_mcp_server())
            return
        if self.path == "/api/stop":
            self._send_json(stop_mcp_server())
            return
        if self.path == "/api/test-origin":
            self._send_json(test_origin_connection())
            return
        if self.path == "/api/configure":
            try:
                body = self._read_json_body()
            except (json.JSONDecodeError, ValueError) as exc:
                self._send_json({"ok": False, "message": str(exc)}, 400)
                return
            result = configure_client(str(body.get("client", "")))
            status = 200 if result.get("ok") else 400
            self._send_json(result, status)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args: Any) -> None:
        STATE.log(format % args)


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), UIHandler)
    STATE.log(f"UI listening on http://{HOST}:{PORT}/")
    print(f"OriginLab MCP UI: http://{HOST}:{PORT}/", flush=True)
    if os.environ.get("ORIGINLAB_MCP_UI_NO_BROWSER") != "1":
        webbrowser.open(f"http://{HOST}:{PORT}/")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        stop_mcp_server()
        server.server_close()


if __name__ == "__main__":
    main()
