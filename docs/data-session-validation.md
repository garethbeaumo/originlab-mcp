# Data/session repair validation — 2026-11-02

Baseline: `41ad0c37d5a49a7bfb3af6b382d7cd6f654909ba` (main already contains
PRs #9 and #10). This repair does not establish licensed Origin acceptance.

## Reproduced contracts

- Locked originpro 1.1.15 `WSheet.sort()` calls `lt_col_index(col)` internally.
  Pass zero-based indices once. Its prose says 1-offset, but its implementation
  and `sort(0)` example disagree with that prose. Portable tests use three
  distinct keys in both directions and verify complete row correspondence.
- `WSheet.sort()` and `clear()` return integer 0 on success, otherwise an error
  code. Nonzero, absent status and exceptions must not become MCP success.
- `clear(c1, c2=...)` uses inclusive endpoints. Omitted start means the first
  column, omitted end means the last. Reject negative/out-of-range/reversed
  explicit endpoints before autosave or mutation; revalidate before clearing.
- Core page enumeration (`w`, `g`, `m`) must finish successfully. Exceptions,
  missing iterables and generator failures propagate as errors, not empty lists.
- The locked `pages()` only supports `w/m/g/i`; it silently returns no matches
  for Notes/Excel. These optional counts/lists are `null`, with `partial=true`
  and explicit warnings, rather than false zero counts. An alternative backend
  can advertise supported optional filters with `session_page_types`; failures
  then have `enumeration.<kind>.status=error`, distinct from `unsupported`.
  A supported successful empty enumeration has count 0 and status `ok`.

The initial focused regression suite produced **27 failed / 7 passed** on
unchanged production code after correcting the fake backend's native contract.
It passed after the fixes. Tests never attach to Origin.

## Transport and lifecycle acceptance

`tests/test_data_session_protocol.py` starts the production server registrations
over real subprocess stdio with a test-only backend. It covers success/nonzero/
absent/exception native outcomes, empty/core-error/partial session responses,
and `isError` flags. Its lifecycle test sends an MCP cancellation notification,
checks a subsequent discovery request, disconnects during a bounded operation,
restarts the MCP process and initializes again. The fake backend records detach,
completion and exit calls. Already-started work may finish after cancellation;
the test does not claim native abort or persistent-session reattachment.

Community references motivating these checks (not evidence of Origin data loss):

- [MATLAB #126](https://github.com/matlab/matlab-mcp-server/issues/126): reports
  client cancellation/restart causing a launched MATLAB process to be killed.
- [Jupyter #498](https://github.com/datalayer/jupyter-mcp-server/pull/498): checks
  readiness explicitly instead of reporting an unsynced model as empty.

## Live checks: NOT RUN

On this Windows host, no running Origin process or conventional OriginLab
installation directory was found. The Origin.Application ProgID exists, but its
CLSID has no LocalServer32 registration. A licensed isolated instance could not
be verified. No connection/reset/close was attempted, and no user project was
touched.

Before live acceptance, use only a disposable licensed instance and temporary
project. Save a three-column dataset with distinct key ordering; sort each key
in both directions and verify all rows. Clear first/middle/last ranges and
single-column/end-only intervals; verify untouched cells after save/reopen.
Exercise native failures, core enumeration failures and optional unavailable
types. Record exact versions, SHA, logs, retained project and reopened values.

For lifecycle acceptance, record Origin PID/project identity before a long call,
cancel it, disconnect stdio and restart MCP. Check whether started work finishes,
whether queued work runs, and whether the original Origin and unsaved temporary
project survive. Verify detach/reconnect without closing a user-owned instance.
Do not kill Origin or automatically retry an uncertain mutation. COM ownership,
native cancellation, persistent identity and modal recovery remain NOT VERIFIED.

The merge restriction in [follow-up.md](follow-up.md) applies: "Do not merge
runtime-dependent changes or publish on the strength of fake tests alone."
Keep this PR draft until the live evidence exists. Portable CI/stdio success
does not waive that restriction. Preflight and mutation are still separate
dispatches and are not an atomic transaction against manual UI changes.
