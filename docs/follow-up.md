# Evidence-driven follow-up

## Audited baseline — 2026-10-29

[PR #8](https://github.com/garethbeaumo/originlab-mcp/pull/8) is merged, with
final head `809307fbb86fa4feb5d086a5bbdea8952f05f001` and merge/main commit
`e157772d77a3bc2e0a2d126508e8da4802b29f57`. These two commits have the same
Git tree. The previous `4b42cb5` instruction referred to an earlier review round;
start new work from the latest `main`, recording its SHA before making changes.

The merge is a completed repository action, not proof of live acceptance.
**Windows + licensed Origin acceptance remains NOT VERIFIED.** See
[validation status](validation-status.md) for confirmed CI evidence, historical
reports and publication status. Do not merge runtime-dependent changes or publish
on the strength of fake tests alone; documentation corrections can be reviewed
independently without claiming runtime acceptance.

| Priority | Work | Evidence / acceptance | Boundary |
| --- | --- | --- | --- |
| P0 — NOT VERIFIED | Run the Windows acceptance checklist | [Checklist](windows-validation.md): retain commit/versions, exit codes, logs, project/PNG/CSV, reopened values, and each manual outcome; latest handoff records no successful live run | Requires licensed interactive Windows + Origin; portable CI is insufficient |
| P0 — OPEN | COM owner and lifecycle design | `run_with_soft_timeout` creates a per-call executor; lifecycle entry points and timer detach also access the backend. Serialization does not establish one COM owner. Add owner-thread tests and verify connect/call/release/reconnect/idle/shutdown, including a blocked modal dialog, on Windows | Keep user-owned Origin alive; no timeout force-kill; already-started COM calls are not cancelled |
| P0 — OPEN | Atomic preflight + destructive dispatch | Preflight and mutation still use separate `execute()` calls. Add deterministic competing-call tests proving the saved project identity is the one mutated; design one owner dispatch for both, and separately detect/handle manual UI project switches | A successful save does not make the pair transactional or protect against external UI changes |
| P1 | Authoritative live project identity | Cached-path reuse reproduced; current fix refuses cache-only saves. Validate a documented live full filename API, readonly behavior and same-directory UI switches | Unknown/ambiguous path skips autosave; strict mode blocks mutation |
| P1 | Actual periodic save semantics | `_reset_autosave_timer` runs after every operation; default interval equals idle timeout | Prove repeated editing and idle behavior with a controlled clock and Windows run |
| P2 | Broaden real API contract coverage | A fake accepted `pages("Book")` / `pages("Graph")`, hiding empty real results | Validate high-use APIs against locked package and Windows outputs before adding tools |

The queue-late-write issue was independently reproduced with threading Events;
the current patch bounds lock acquisition so expired queued work never starts.
This does not prove COM apartment correctness or cancel an already-running COM call.

Community leads supplied by the parent task (reference material, not proof of
this repository's behavior):

- Excel MCP [issue #888](https://github.com/sbroenne/mcp-server-excel/issues/888)
  and [PR #926](https://github.com/sbroenne/mcp-server-excel/pull/926): timeout work.
- [ThepExcel session owner](https://github.com/ThepExcel/ThepExcelMCP/blob/main/src/thepexcel_mcp/session.py)
  and [Origin-Pro-MCP daemon](https://github.com/youngminsw/Origin-Pro-MCP/blob/main/src/origin_pro_mcp/daemon.py): lifecycle design references.
- MATLAB [issue #72](https://github.com/matlab/matlab-mcp-server/issues/72) and
  ThepExcel [PR #10](https://github.com/ThepExcel/ThepExcelMCP/pull/10): identity/cache boundary leads.
- [MCP tool specification](https://modelcontextprotocol.io/specification/2025-11-25/server/tools): error-contract follow-up.

Prioritize one reproduced problem with a clear acceptance test per iteration.
Keep future PRs making runtime-dependent claims draft until those claims have
live evidence. This cannot retroactively change PR #8's merged state or mark
its unperformed acceptance complete.

## Completed protocol follow-up

The bounded round after `e0cf819` reproduced three `ok:false` responses marked
`isError:false` over real stdio. A transport-boundary adapter now preserves the
existing JSON text while setting the protocol error flag; schema and unknown-tool
errors remain SDK-managed. Seven wire regressions pass. The SDK remains locked at
1.26.0, with the declaration now `>=1.26.0,<2`. See
[protocol handoff](protocol-handoff.md) for scope, evidence and compatibility limits.


## Completed bounded result-contract follow-up

The October 12 Windows round corrects fitting method calls and flat NLFit results,
checks actual export paths and partial batches, bounds worksheet previews before
cell transfer, and reports graph objects/steps retained after creation failures.
See [result contract handoff](result-contract-handoff.md) for API sources,
**269 passing portable tests on Windows with a fake Origin backend**, installed
stdio evidence, and live Origin limits. The audited main CI separately confirms
269 tests and 71% combined coverage on Linux; neither is a licensed Origin run.
COM owner/lifecycle and atomic preflight/mutation work remain separate.
