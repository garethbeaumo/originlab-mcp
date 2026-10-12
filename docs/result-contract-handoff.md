# Result contract follow-up — 2026-10-12

Based on PR #8 head `0ad9a849f53d127caa3d1a1b761a5d6107ba8687`,
with `main` at `09962a16561f204e559ca366090bfe4414991ec8`.
Work was done in an independent Windows checkout. No existing user checkout,
Origin project, credentials, permissions, or dependency lock was changed.

## Contracts and behavior

- **Fitting:** checked installed, locked originpro **1.1.15** `analysis.py`.
  `LinearFit.fix_slope` and `fix_intercept` are methods despite assignment syntax
  in their upstream examples. Both zero-valued constraints now call the methods.
  Linear results retain their nested parameter structure. `NLFit.result()`
  returns the flat nltree: the complete tree is exposed as `data.result`, with
  recognized parameters also normalized under `data.parameters`. Statistics and
  custom fields are retained in the raw tree, not guessed as nested Statistics.
  When a report is requested, `report()` precedes `result()`, as required by
  upstream. Empty results/ranges and report/result exceptions are failures.
- **Export:** `save_fig` returns the actual generated path or an empty string.
  Missing/invalid paths fail; requested paths never substitute for an unknown
  result. Batch output retains confirmed `files`, records `failed`, and sets
  `ok:false` for any failure. Enumeration failures preserve earlier confirmed
  files and mark `enumeration_complete:false`. Exceptions can happen after a
  write, so failure means unconfirmed output, not proof that no file exists.
- **Preview:** both tool and resource/session previews use `to_list2` with
  explicit inclusive endpoints. Zero rows or columns bypass cell reads entirely.
  Maximum rectangle: **100 rows × 32 columns**, default 20 rows; the data tool
  adds optional `max_cols` (default 32). Cell strings are capped at 256 characters,
  display names at 128 plus uniqueness suffixes, with a 20,000-character budget
  shared by display names and cell text. Repeated JSON keys/formatting are outside
  that text-value budget. Metadata is also limited to 32 columns, 256 characters
  per label, with omitted columns and clipped fields identified. COM cell-read
  failures propagate instead of becoming empty successful previews.
  `total_rows` continues to mean the used row extent including gaps (`wks.rows`,
  backed by `wks.maxRows`); `allocated_rows` separately reports `wks.shape[0]`.
  No scan or full-column transfer is needed to count rows. Full strings within
  the requested rectangle still cross COM before response truncation.
- **Partial plots:** graph identity is recorded before adding curves. Failures
  preserve confirmed curves, completed steps, failed step, and creation state
  (`not_attempted`, `unknown`, or `created`). The failing COM call may have had
  side effects; current state is explicitly unverified. No automatic cleanup,
  retry, or COM read occurs in the error handler. Timeout responses use stable
  snapshots even if the worker later completes.

The existing MCP adapter maps these `ok:false` responses to `isError:true`.
All 67 tools retain their annotations. No COM owner/lifecycle redesign or atomic
preflight/mutation refactor is included.

## Verification

Windows 11 / Python 3.12.4, MCP SDK 1.26.0, originpro 1.1.15:

| Check | Result |
| --- | --- |
| Full pytest | **269 passed**, 71% coverage, exit code **0** |
| Real stdio regression file | **18 tests** included in the full suite |
| Ruff (`src tests scripts/verify_mcp.py`) | Passed |
| Mypy | Passed, 28 source files |
| sdist and wheel built from sdist | Passed |
| Separate wheel install, portable stdio and fake workflow | Passed, exit codes **0**; 67 tools / 5 prompts / 3 fixed resources |


- Unit regressions cover X=[1,2,3], Y=[3,5,7]: free slope/intercept **2 / 1**;
  intercept fixed at zero: slope **17/7**; slope fixed at zero: intercept **5**.
- Strict fake rejects overwriting the constraint methods. NLFit fake emits a flat
  tree and enforces report-before-result ordering. Both fitting tools inject
  empty results/ranges and exceptions.
- Large fake worksheet confirms exactly `(r1,r2,c1,c2)=(0,2,0,3)` for a 3×4
  request. Tests cover zero/negative bounds, long values and duplicate names,
  allocated versus used rows, malformed blocks and COM read exceptions.
- Export tests cover empty/None/False returns, changed actual paths, mixed
  success, and interrupted enumeration. Plot tests cover preflight, uncertain
  creation, the second curve, rescale/group, and a continuing timed-out worker.
- `tests/test_protocol.py` uses the real stdio transport and ClientSession with
  a fake Origin backend: numerical fitting, report/result failures, partial plot
  and export, discovery, SDK and business error contracts.
- `scripts/verify_mcp.py --fake-workflow` performs numerical and failure
  acceptance over real stdio. The `--live` path now includes the same numerical
  fitting assertions. Default portable mode only performs discovery/diagnostics.
- Build with `uv build` (wheel built from sdist); install wheel in a separate
  environment and run portable and fake-workflow stdio checks there. Verify the
  imported package path points into that environment's site-packages.

Commands: `uv run --no-sync pytest -q`, `uv run --no-sync ruff check src tests`,
`uv run --no-sync mypy src`, `uv run --no-sync python scripts/verify_mcp.py`,
`uv run --no-sync python scripts/verify_mcp.py --fake-workflow`, `uv build`.

## Live Origin boundary

No Origin application process, conventional OriginLab installation directory,
Origin Application COM registration, or Origin uninstall entry was found on
this Windows machine. Installing the Python bridge is not an Origin license.
**No successful live Origin connection, fitting, rendering, save/reopen, or
license validation was performed.** No current user project was closed or replaced.

An existing validation-only test initially attempted to import the installed
native bridge through an autosave preflight. All assertions passed, but OriginExt
aborted at interpreter exit (`0xC0000409`; isolated traceback in its Exit/Detach
handler). Portable tests now explicitly reject importing the live backend and
continue using fake objects, including fake reconnects. The complete suite then
passed with exit code 0. Production connect/detach behavior was not changed.

Keep PR #8 draft. Run the live script only in a confirmed disposable, licensed
Origin session: its explicit `--allow-project-reset` mode resets the attached
project and must never be used with a user-owned project. Portable and fake
backend results do not establish release readiness.

## Sources

- [LinearFit API](https://docs.originlab.com/originpro/classoriginpro_1_1analysis_1_1LinearFit.html)
  and [NLFit API](https://docs.originlab.com/originpro/classoriginpro_1_1analysis_1_1NLFit.html),
  cross-checked against the locked package source.
- [NLFit tree](https://docs.originlab.com/x-function/ref/nlbegin/).
- [WSheet API](https://docs.originlab.com/originpro/classoriginpro_1_1worksheet_1_1WSheet.html)
  and [used versus allocated rows](https://docs.originlab.com/labtalk/ref/wks-obj/).
- [GPage save_fig](https://docs.originlab.com/originpro/classoriginpro_1_1graph_1_1GPage.html).
- Community design references (not Origin API proof):
  [bounded overview #1062](https://github.com/sbroenne/mcp-server-excel/pull/1062),
  [partial creation #1070](https://github.com/sbroenne/mcp-server-excel/pull/1070).
