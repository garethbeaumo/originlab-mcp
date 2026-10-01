# Bounded MCP protocol follow-up

Starting checkpoint: `e0cf819d4fb288101c7a7f661f8c3cdad8f3be88`.
The preceding API/save fixes remain unchanged and independently reviewable.

## Reproduction and fix

A subprocess runs the production server registrations with a controlled fake
Origin backend. The client is the actual SDK `ClientSession` over stdio, not
DummyMCP. Three cases reproduced the same mismatch:

| Case | Existing JSON payload | Before | After |
| --- | --- | --- | --- |
| Empty `execute_labtalk.command` | `ok:false`, `invalid_input`, target `command` | `isError:false` | `isError:true` |
| `list_worksheets` backend raises RuntimeError | `ok:false`, `internal_error` | `isError:false` | `isError:true` |
| Finite blocked backend exceeds dispatch budget | `ok:false`, `timeout`, target `dispatch_timeout` | `isError:false` | `isError:true` |

Malformed parameter types and unknown tools were already reported as protocol
errors by the SDK. They still are. Successful calls retain `isError:false` and the
same JSON shape. All 67 tools retain input schemas, annotations and descriptions;
this round does not change their output schema or add structuredContent.

`OriginFastMCP.call_tool` uses the public FastMCP handler hook. It recognizes the
existing single JSON TextContent envelope with top-level `ok:false` and wraps the
**unchanged content** in `CallToolResult(isError=True)`. It leaves argument
validation and all other SDK handling intact. Python callers and DummyMCP tests
continue receiving ordinary dictionaries. There is no backend/lifecycle redesign.

This follows the [MCP tool error convention](https://modelcontextprotocol.io/specification/2025-11-25/server/tools).
The fake controls only the Origin boundary; these results do not establish
Windows COM correctness or reproduce a real modal dialog.

## SDK compatibility

- Installed, locked and exercised version: **MCP Python SDK 1.26.0**.
- Declaration changed from `mcp[cli]>=1.0.0` to `mcp[cli]>=1.26.0,<2`.
- No package version in the lockfile changed; only the project requirement did.
- 1.26.0's handler supports CallToolResult passthrough and serializes this
  project's unparameterized dict returns as one JSON text block. Both facts were
  verified in the installed SDK source and by real client/server calls.
- Earlier SDKs are not claimed supported. Future 1.x versions are permitted by
  the declaration but require the wire regression suite before updating the lock.
  SDK v2 is deliberately excluded. New structured-output tools will need an
  explicit error-envelope contract before adopting a different return shape.

## Checks and evidence

- Initial wire regression: **3 failed, 4 passed**, failures at `isError` assertion.
- Corrected wire regression: **7 passed**.
- Final full suite: **219 passed, 70% coverage** on macOS/Python 3.12.13.
- Ruff passed; Mypy passed for 28 source files.
- Build and installed-wheel portable smoke are recorded in
  `out/validation/protocol-build.log` and `protocol-wheel-smoke.log`.
- Before/after logs: `out/validation/protocol-before.log`,
  `protocol-after.log`, `protocol-final-pytest.log`.
- Windows/Origin and local Python 3.10/3.11 were **not run** this round.

## Read-only remote verification

GitHub PR and Git-ref reads both returned:

- PR #8: open, draft, head repository `garethbeaumo/originlab-mcp`.
- Head branch: `cursor/ci-annotations-mypy-87ad`.
- Head SHA: `4b42cb5465c69e61a0d92d8fcc5b524a7dd1dda0`.
- `git merge-base --is-ancestor` confirms this head is an ancestor of `e0cf819`
  and the new local commit, so the current history allows a fast-forward update.

No push was attempted. Write authorization, branch protections and a race with
another remote update remain separate from ancestry; re-read the head immediately
before any user-authorized push. Never force-push this work.

Next priority remains Windows acceptance, then a bounded COM-owner/atomic
preflight design, as recorded in `follow-up.md`.
