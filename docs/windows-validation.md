# Windows + Origin acceptance check

The portable test suite uses fake Origin objects. It does **not** prove COM thread
ownership, Origin attachment, fitting, rendering, or file persistence on Windows.
Keep PR #8 draft until the checks below have been performed on the intended host.

## Required environment

- Windows with an interactive desktop and licensed OriginLab 2021+.
- A supported 64-bit Python (start with 3.12) and `uv`; use the same interpreter
  architecture as the installed Origin/OriginExt package.
- This checkout including its `uv.lock`. Record `git rev-parse HEAD`, Windows,
  Origin, Python, originpro and OriginExt versions with the results.
- Save and close valuable projects and other Origin instances first. The live
  script **creates a new project**, so use only a disposable Origin session.

## Commands (PowerShell, repository root)

```powershell
uv sync --frozen --group dev
uv run --no-sync ruff check src tests scripts/verify_mcp.py
uv run --no-sync mypy src
uv run --no-sync pytest -q
uv run --no-sync python scripts/verify_mcp.py

# Choose a directory that does not exist yet. This resets the active project.
uv run --no-sync python scripts/verify_mcp.py --live --allow-project-reset --output-dir C:\temp\origin-smoke-001 2>&1 | Tee-Object origin-smoke.log
```

The portable command checks actual stdio initialization, tools/prompts/resources
discovery and `originlab_doctor` without connecting to Origin. The live command
imports three rows of synthetic data, lists worksheets, creates and lists a
scatter plot, reads the session, saves, batch-exports a PNG and a CSV, checks
nonempty output files, reopens the project, and releases Origin. Inspect the
retained graph and CSV visually; verify that the saved/reopened worksheet contains
X = 1, 2, 3 and Y = 2, 4, 6.

The script sets allowed roots to its output directory and disables automatic
saves for its child process. This makes the initial reset explicit and prevents
background saves from obscuring which operation wrote a file. It does not test
periodic autosave. It never force-kills Origin.

## Manual checks before approval

1. In a trusted project, run the UI and configure Codex. Inspect
   `.codex/config.toml`, verify existing settings/comments remain, then check
   `codex mcp get originlab` and invoke `originlab_doctor` from the client. JSON
   clients keep their existing configuration format.
2. With only a disposable project, test `save_project` to an unwritable target
   and opening a corrupt `.opju`: each must return an error, not success. Verify
   the actual current project remains consistent after a failed open.
3. Test `ORIGINLAB_MCP_AUTOSAVE_REQUIRED=1`: missing/ambiguous live project path,
   failed save, and an out-of-root path must block `new_project`. A successful
   explicit save alone is not proof that live path discovery works on that Origin
   version. If an implicit save cannot establish a live path, supply `file_path`
   explicitly and record this limitation.
4. Open project A, release, manually open project B, reconnect, then inspect.
   No implicit or periodic save may write B into A's cached path. Repeat a UI
   project switch without releasing. If path metadata is unavailable, skipping
   the save (or blocking in required mode) is the expected conservative behavior.
5. Observe periodic autosave on a disposable project with interval 10 seconds,
   including editing repeatedly and idling. Record actual timestamps. Scheduling
   currently resets after calls and shares the default 300-second idle window;
   this is not yet validated as a reliable wall-clock backup mechanism.
6. Induce a modal dialog only in the disposable session. Record the soft timeout,
   clear the dialog, and verify that a queued operation whose budget expired did
   not run. An operation that had already started may still finish after timeout.
   Do not retry a write blindly. COM ownership and shutdown while blocked need
   Windows validation and a separate owner-thread design.

Preflight save and the subsequent destructive operation still use separate
`execute()` calls. They are not an atomic transaction against another MCP call or
manual UI changes. In-place autosave overwrites the current file; it is not a
versioned backup. Session Notes/Excel lists are best effort: originpro 1.1.15
`pages()` enumerates only worksheet, matrix, graph and image pages.

## Record the outcome

Record PASS / FAIL / NOT RUN for portable checks, live workflow, config loading,
failure handling, identity switching, periodic saving, and modal timeout. Attach
`origin-smoke.log`, `coverage.xml`, output files and version information. A green
fake suite cannot replace any NOT RUN live item.
