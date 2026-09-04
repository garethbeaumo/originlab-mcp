<p align="center">
  <h1 align="center">OriginLab MCP Server</h1>
  <p align="center">
    <strong>Use AI as your OriginLab assistant</strong>
  </p>
  <p align="center">
    Connect OriginLab data analysis and visualization to Antigravity, Claude, Cursor, and other AI clients through the <a href="https://modelcontextprotocol.io">Model Context Protocol</a>.
  </p>
  <p align="center">
    <a href="https://github.com/garethbeaumo/originlab-mcp/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License"></a>
    <img src="https://img.shields.io/badge/python-3.10%2B-blue.svg" alt="Python">
    <img src="https://img.shields.io/badge/version-0.2.1-green.svg" alt="Version">
    <img src="https://img.shields.io/badge/platform-Windows-lightgrey.svg" alt="Platform">
    <img src="https://img.shields.io/badge/tools-66-orange.svg" alt="Tools">
  </p>
  <p align="center">
    <a href="#quick-start">Quick Start</a> · <a href="#features">Features</a> · <a href="#examples">Examples</a> · <a href="#client-configuration">Client Configuration</a>
  </p>
  <p align="center">
    <a href="README.zh.md">简体中文</a> · English
  </p>
</p>

## What Is OriginLab MCP Server?

OriginLab MCP Server is a bridge between AI clients and OriginLab. It lets you import data, create plots, customize figures, run analysis, and export results through natural-language requests instead of manually operating the Origin UI.

```text
User: Import experiment.csv from my desktop into Origin, use the first column as X
      and the second column as Y, create a scatter plot, run a Gaussian fit,
      then export the graph as PNG.

AI: Done.
    Imported experiment.csv -> Sheet1 (200 rows x 5 columns)
    Created scatter plot -> Graph1
    Gaussian fit completed -> xc=2.35, w=0.82, A=156.3, R²=0.9987
    Exported -> C:\Users\Desktop\Graph1.png
```

### How It Works

```text
AI Client (Antigravity / Claude / Cursor)
       ↓ MCP over stdio
OriginLab MCP Server (Python)
       ↓ originpro + COM
OriginLab
```

## Requirements

| Requirement | Details |
| :--- | :--- |
| Operating system | Windows |
| OriginLab | OriginLab 2021 or later with a valid license |
| Python | 3.10+ |

## Quick Start

For AI assistants or one-command local setup, run this from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-and-open.ps1
```

The script installs `uv` if needed, runs `uv sync`, starts the local status panel, and opens `http://127.0.0.1:8765/` automatically. From that page you can test Origin and write MCP client configs.

Choose one installation method.

<details open>
<summary><b>Option A: uv (recommended)</b></summary>

**1. Install uv**

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

**2. Install dependencies**

```powershell
cd C:\path\to\originlab-mcp
uv sync
```

**3. Start the server**

```powershell
uv run originlab-mcp
```

</details>

<details>
<summary><b>Option B: pip</b></summary>

**1. Create a virtual environment (optional but recommended)**

```powershell
cd C:\path\to\originlab-mcp
python -m venv .venv
.venv\Scripts\activate
```

**2. Install the project**

```powershell
pip install -e .
```

**3. Start the server**

```powershell
originlab-mcp
```

</details>

After startup, the server waits for MCP client requests over stdio. The first tool call automatically connects to the local OriginLab installation.

### Optional: Local Status Panel

To check whether the MCP server can start, whether Origin can be reached, and whether common client config files exist, start the local UI:

```powershell
uv run originlab-mcp-ui
```

Then open `http://127.0.0.1:8765/`. This page can start/stop a debug MCP server subprocess, test the Origin connection, **read the current Origin session** (worksheets and graphs), and write the `originlab` MCP configuration for Antigravity / Gemini, Cursor, Codex, Trae, and Claude Desktop. For normal use, the AI client should still start the server automatically from its configuration.

When updating an existing config file, the UI creates a `.bak-timestamp` backup first. To prevent the browser from opening automatically:

```powershell
$env:ORIGINLAB_MCP_UI_NO_BROWSER = "1"
uv run originlab-mcp-ui
```

## Features

The server provides **66 tools** covering the OriginLab data workflow.

### Data Management (14 Tools)

| Category | Tools |
| :--- | :--- |
| Import | `import_csv` · `import_excel` · `import_data_from_text` |
| Inspect | `list_worksheets` · `get_worksheet_info` · `get_worksheet_data` · `get_cell_value` |
| Edit | `set_column_designations` · `set_column_labels` · `set_column_formula` |
| Manage | `add_worksheet` · `sort_worksheet` · `clear_worksheet` · `delete_columns` |

### Plotting (11 Tools)

| Category | Tools |
| :--- | :--- |
| Create | `create_plot` · `create_double_y_plot` |
| Modify | `add_plot_to_graph` · `remove_plot_from_graph` · `change_plot_data` · `change_plot_type` |
| Layers | `add_graph_layer` · `group_plots` |
| Inspect | `list_graphs` · `list_graph_templates` · `get_graph_info` |

### Graph Customization (25 Tools)

| Category | Tools |
| :--- | :--- |
| Axes | `set_axis_range` · `set_axis_scale` · `set_axis_step` · `set_axis_title` |
| Lines | `set_plot_line_style` · `set_plot_line_width` |
| Fonts and ticks | `set_graph_font` · `set_tick_style` |
| Colors | `set_plot_color` · `set_plot_colormap` · `set_plot_transparency` |
| Symbols | `set_plot_symbols` · `set_symbol_size` · `set_symbol_interior` |
| Group increments | `set_color_increment` · `set_symbol_increment` |
| Error bars | `set_error_bar_style` |
| Fill | `set_fill_area` |
| Legend | `set_legend` |
| Preset styling | `apply_publication_style` |
| Annotations | `set_graph_title` · `add_text_label` · `add_line_to_graph` · `remove_graph_label` |

### Analysis (3 Tools)

`linear_fit` · `nonlinear_fit` · `list_fit_functions`

> Common Origin fit functions such as Gauss, Lorentz, ExpDec1, and Boltzmann are supported. You can provide initial parameters, fix parameters, and fit with error bars.

### Export and Project Management (6 Tools)

`export_graph` · `export_all_graphs` · `export_worksheet_to_csv` · `save_project` · `open_project` · `new_project`

### System Management (5 Tools)

`get_origin_info` · `read_origin_session` · `release_origin` · `reconnect_origin` · `close_origin`

> `read_origin_session` is a read-only snapshot of the current project: workbooks/worksheets, graphs, matrices, notes, active objects, and project path. Pass `include_preview=true` to include a truncated preview of the active worksheet.

### MCP Resources (Session Reading)

In addition to tools, clients can inspect the Origin session through MCP `resources/read` without changing the project:

| URI | Contents |
| :--- | :--- |
| `originlab://session` | Full project snapshot |
| `originlab://worksheets` | Worksheet list |
| `originlab://graphs` | Graph list |
| `originlab://worksheet/{book}/{sheet}` | One worksheet's columns and data preview |
| `originlab://graph/{name}` | One graph's layers and curves |

### Advanced (2 Tools)

`execute_labtalk` · `get_labtalk_variable`

`execute_labtalk` is an escape hatch for operations not covered by standard tools. `get_labtalk_variable` safely reads LabTalk variable values.

## Examples

| Request | Typical tool call |
| :--- | :--- |
| Inspect the current Origin project | `read_origin_session` |
| Import `data.csv` into Origin | `import_csv` |
| Show worksheet columns and metadata | `get_worksheet_info` |
| Set the first column as X and the next two columns as Y | `set_column_designations` |
| Create a scatter plot | `create_plot` |
| Add another curve from the third column | `add_plot_to_graph` |
| Change the X-axis title and set the curve color to red | `set_axis_title` + `set_plot_color` |
| Apply a publication-style preset | `apply_publication_style` |
| Apply publication styling to layer 2 with line width 3 and symbol size 12 | `apply_publication_style` |
| Run a Gaussian fit | `nonlinear_fit` |
| Set the Y axis to logarithmic scale | `set_axis_scale` |
| Export a graph as PNG | `export_graph` |
| Export all project graphs as SVG | `export_all_graphs` |
| Release Origin so I can use it manually | `release_origin` |

`apply_publication_style` supports `layer_index`, axis-title font size, tick-label font size, legend font size, major tick length, minor tick count, line width, symbol size, and other common publication-figure settings.

Typical workflow:

```text
Import data -> inspect structure -> set column designations -> create plot -> customize graph -> run analysis -> export results
```

## Client Configuration

> [!NOTE]
> You do not need to start the MCP server manually. Once configured, the AI client starts the server process when needed and communicates with it over stdin/stdout.

Replace paths with your actual project path.

You can also run `uv run originlab-mcp-ui`, choose a client in the local status panel, and let it write the config automatically. Manual examples are shown below.

### Antigravity (Gemini) - Recommended

Create `.gemini/settings.json` in the project root.

**Using uv:**

```json
{
  "mcpServers": {
    "originlab": {
      "command": "uv",
      "args": ["--directory", "C:\\path\\to\\originlab-mcp", "run", "originlab-mcp"]
    }
  }
}
```

**Using pip:**

```json
{
  "mcpServers": {
    "originlab": {
      "command": "C:\\path\\to\\originlab-mcp\\.venv\\Scripts\\originlab-mcp.exe"
    }
  }
}
```

<details>
<summary><b>Claude Desktop</b></summary>

Edit `%APPDATA%\Claude\claude_desktop_config.json` using the same format as above.

</details>

<details>
<summary><b>Cursor</b></summary>

Create `.cursor/mcp.json` in the project root using the same format as above.

</details>

<details>
<summary><b>Trae</b></summary>

Create `.trae/mcp.json` in the project root using the same format as above.

</details>

<details>
<summary><b>Codex (OpenAI)</b></summary>

Create `.codex/config.json` in the project root using the same format as above.

</details>

## Testing

```powershell
# uv
uv run python -m pytest tests/ -v

# pip, after activating the virtual environment
pytest tests/ -v
```

The basic test suite does not require OriginLab to be installed. On Linux/macOS, `originpro` is skipped automatically because it is marked Windows-only in `pyproject.toml`. GitHub Actions runs ruff, mypy, and pytest (with coverage XML) on Python 3.10–3.12.

## Project Structure

```text
originlab-mcp/
├── pyproject.toml                # Project configuration and dependencies
├── CHANGELOG.md                  # Release notes
├── .github/workflows/ci.yml      # Lint / typecheck / unit tests
├── scripts/
│   └── install-and-open.ps1      # One-command setup and UI launcher
├── src/originlab_mcp/
│   ├── server.py                 # MCP server entry point and dependency injection
│   ├── ui.py                     # Local status panel
│   ├── origin_manager.py         # Thread-safe Origin COM connection manager
│   ├── session.py                # Read-only Origin session snapshot
│   ├── resources.py              # MCP Resources (session reading)
│   ├── exceptions.py             # Custom exceptions
│   ├── types.py                  # Protocol type definitions
│   ├── tools/
│   │   ├── data.py               # Data import and worksheet management (14)
│   │   ├── plot.py               # Plot creation and graph management (11)
│   │   ├── customize.py          # Graph customization (25)
│   │   ├── analysis.py           # Linear and nonlinear fitting (3)
│   │   ├── export.py             # Export and project management (6)
│   │   ├── system.py             # System and connection management (5)
│   │   └── advanced.py           # LabTalk escape hatch (2)
│   └── utils/
│       ├── annotations.py        # MCP ToolAnnotations presets
│       ├── constants.py          # Enums, defaults, and fit-function metadata
│       ├── helpers.py            # Graph/sheet resolution and error handling helpers
│       └── validators.py         # Input validation and standard response builders
└── tests/
    ├── conftest.py               # Shared fixtures (manager / fake Origin)
    ├── fakes/                    # In-memory OriginPro + DummyMCP doubles
    ├── test_annotations.py       # ToolAnnotations coverage tests
    ├── test_origin_contract.py   # Multi-tool COM contract workflows
    ├── test_tool_guidance.py     # Description / next_suggestions contracts
    ├── test_helpers.py           # Helper tests
    ├── test_phase3.py            # LabTalk safety and resolve-pattern tests
    ├── test_session.py           # Session reading and MCP resource tests
    ├── test_tools.py             # Tool registration and integration tests
    └── test_ui.py                # Local status panel config tests
```

## FAQ

<details>
<summary><b>Origin connection failed</b></summary>

Check that:

- OriginLab 2021 or later is installed locally
- You have a valid OriginLab license
- The current user can start Origin
- No other program is blocking the Origin COM interface

</details>

<details>
<summary><b>The MCP client cannot see the tools</b></summary>

1. Confirm the client configuration path is correct.
2. Confirm dependencies are installed with `uv sync` or `pip install -e .`.
3. Restart the MCP client.

</details>

<details>
<summary><b>Fit results look inaccurate</b></summary>

1. Check for outliers with `get_worksheet_data`.
2. Provide better initial values through `initial_params`.
3. Fix known parameters with `fixed_params`.
4. Confirm that the fit function is appropriate. Use `list_fit_functions` to inspect common options.

</details>

## License

[MIT](LICENSE) © 2025 garethbeaumo
