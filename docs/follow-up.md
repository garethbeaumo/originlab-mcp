# Evidence-driven follow-up

Work from the verified PR #8 head `4b42cb5465c69e61a0d92d8fcc5b524a7dd1dda0`.
Do not merge or publish on the strength of fake tests alone.

| Priority | Work | Evidence / acceptance | Boundary |
| --- | --- | --- | --- |
| P0 | Run the Windows acceptance checklist | `windows-validation.md`, retained files and logs | Requires licensed interactive Windows + Origin |
| P0 | COM owner and lifecycle design | Locally reproduced distinct worker thread IDs across serialized calls; verify connect/call/detach/idle/shutdown share the required owner on Windows | Keep user-owned Origin alive; no timeout force-kill |
| P0 | Atomic preflight + destructive dispatch | Current tools perform two separate `execute()` calls; add deterministic competing-call tests before changing API | Do not claim current save/mutation pair is transactional |
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

Prioritize one reproduced problem with a clear acceptance test per iteration;
retain a draft until runtime-dependent claims are supported by live evidence.

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
**269 passing Windows tests**, installed stdio evidence, and live Origin limits.
COM owner/lifecycle and atomic preflight/mutation work remain separate.
