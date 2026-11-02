# PR #8 local review handoff

Historical local review record: the refs, remote state and local-only actions
below describe that review round. PR #8 has since been merged; see
[current validation status](validation-status.md). No live acceptance is implied.

## Verified starting point

- Repository: https://github.com/garethbeaumo/originlab-mcp
- PR: https://github.com/garethbeaumo/originlab-mcp/pull/8
- Base for this local work: `4b42cb5465c69e61a0d92d8fcc5b524a7dd1dda0`.
- At that starting point PR #8 was open/draft; GitHub reported its
  Python 3.10/3.11/3.12 jobs successful.
  Those are historical checks of the old head, not checks of this patch.
- Independent clone, local branch `codex/pr8-usability`. No pre-existing local
  changes were overwritten. No repository AGENTS.md or `.agents/skills` exists.
  Available local memory summary contained no OriginLab project history.

## Reproduced and corrected

1. `pages("Book")` / `pages("Graph")` returned no real pages. The locked
   originpro 1.1.15 wheel uses `w` / `g` / `m` filters. The original tests' fakes
   accepted invalid filters; tightening those fakes produced six failures before
   correcting list/session/export consumers.
2. Save/open boolean failures were ignored. Required preflight could allow a
   destructive operation after a failed save. Nine new negative tests initially
   failed across false returns, required-save gating and implicit-save roots.
   Explicit, implicit, preflight and periodic save paths now share validation and
   require a truthy native save result. Failed open does not update local context.
3. A current live path different from the cache was ignored. Cache-only,
   ambiguous or incomplete metadata is now rejected for implicit saves. Release,
   reconnect and active-object refresh invalidate old context. This is a
   conservative restriction: explicit saves remain available, but implicit saves
   on an Origin version without supported live metadata may be blocked.
4. An Event-controlled blocked call plus a queued mutation demonstrated that the
   mutation ran after both calls reported timeout. Lock acquisition now consumes
   the original budget; a timed-out queued callback cannot begin later. The
   regression test waits for the worker's lock attempt to finish before release.
   Already-started calls are still soft timeouts, not cancelled operations.
5. Codex UI wrote JSON that Codex does not use. It now writes TOML `mcp_servers`,
   preserves other settings/comments, and backs up prior files. Other clients
   retain JSON. Malformed config is preserved and is not reported configured.

API evidence: [OriginLab project API](https://docs.originlab.com/originpro/namespaceoriginpro_1_1project.html),
plus the exact wheel in `uv.lock`, SHA256
`85f3555a5386af6af5975bc61f96194aacd4f18bf984b579f3c51991143772a5`.
Codex format was checked against the installed CLI help and
[official configuration example](https://developers.openai.com/learn/docs-mcp).

## Local validation

Host: macOS Darwin 25.6.0 arm64, Python 3.12.13, locked MCP 1.26.0.

| Check | Outcome |
| --- | --- |
| Original head full pytest | 191 passed, 69% coverage |
| Final full pytest | 212 passed, 70% coverage |
| Ruff (`src tests scripts/verify_mcp.py`) | Passed |
| Mypy (`src`, 27 files) | Passed |
| `git diff --check` | Passed |
| Offline sdist + wheel build | Passed |
| Install wheel and initialize actual MCP stdio | Passed |
| Discovery | 67 tools, 5 prompts, 3 fixed resources |
| `originlab_doctor(ping_origin=False)` over stdio | Passed; correctly reports degraded Mac/no-Origin environment |
| Local Python 3.10 / 3.11 | Not run for this patch |
| Windows/Origin live workflow, COM ownership, config loading in Windows client | Not run; no Windows/Origin host available |

Machine-readable/local logs and output packages are retained under
`out/validation/` and `out/dist/` (ignored by Git). The Mac environment marked
editable `.pth` files hidden, causing Python to skip them. Verification therefore
used a built, normally installed wheel; its manager source was byte-compared with
the checkout. No workaround was added to the product for that local environment.

## Acceptance and next step

Run [Windows acceptance](windows-validation.md) on a disposable licensed Origin
session, retain outputs, and review the local diff before any remote action.
Prioritize the [follow-up backlog](follow-up.md): COM owner lifecycle, atomic
preflight/mutation, authoritative live project identity, then MCP error semantics.
There is no claim that autosave is a backup, a transaction, or fully live-validated.

No push, merge, release, deployment or repository permission changes were made.
