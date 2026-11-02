# Validation and publication status — 2026-10-29

**PR #8 is merged. Windows + licensed Origin acceptance is NOT VERIFIED, and
release readiness is not established.** Portable tests and real MCP stdio tests
with a fake Origin backend establish a different boundary from live COM behavior.

## Audited repository state

- [PR #8](https://github.com/garethbeaumo/originlab-mcp/pull/8) merged into `main`
  at `2026-10-16T16:00:23Z` (2026-10-17 00:00:23 Asia/Shanghai).
- Main at the audit: [`e157772d77a3bc2e0a2d126508e8da4802b29f57`](https://github.com/garethbeaumo/originlab-mcp/commit/e157772d77a3bc2e0a2d126508e8da4802b29f57).
- Final PR head: [`809307fbb86fa4feb5d086a5bbdea8952f05f001`](https://github.com/garethbeaumo/originlab-mcp/commit/809307fbb86fa4feb5d086a5bbdea8952f05f001).
  Both commits have tree `dd798eb7b93b5fb6bfafd4421c50e32d7852ea08`;
  no later main changes were present at this audit.
- The PR description and older handoffs retain pre-merge draft instructions.
  Those instructions describe historical intent, not current PR state. The merge
  did not supply the missing live evidence or resolve the P0 technical risks.

## Completed checks and their limits

| Check | Evidence and outcome | What it establishes |
| --- | --- | --- |
| Final PR head CI | [Run 38213399682](https://github.com/garethbeaumo/originlab-mcp/actions/runs/38213399682): Python 3.10 / 3.11 / 3.12 jobs succeeded | Portable Ruff, Mypy and pytest on `ubuntu-latest`, not Windows/Origin acceptance |
| Merged main CI | [Run 38928483456](https://github.com/garethbeaumo/originlab-mcp/actions/runs/38928483456): all three Python jobs succeeded | Checks the actual merge commit, not just the old PR head |
| Test count and coverage | [Main Python 3.12 job log](https://github.com/garethbeaumo/originlab-mcp/actions/runs/38928483456/job/117066098593): **269 passed**, **71%** combined line/branch coverage; Ruff passed and Mypy found no issues in 28 source files | Confirmed from CI log; coverage is of Python code exercised by the portable suite |
| Retained CI coverage | Main run artifact `coverage-xml`, ID `12057268746`, available at the audit; XML inspected: 2,438/3,272 lines and 567/962 branches covered across 28 source files | Retained test evidence, not a live acceptance bundle; artifact retention can expire |
| Windows portable suite, build and installed wheel | PR #8 and [October 12 handoff](result-contract-handoff.md) report Windows 11 / Python 3.12.4: 269 tests, 71% coverage, build and installed-wheel smoke passed | Historical reported PASS; not independently rerun on Windows in this audit. The handoff explicitly records no live Origin connection |
| Real MCP stdio contracts | Same handoff reports 18 stdio regressions plus portable/fake-workflow checks; [test host](../tests/protocol_server.py) uses `FakeOrigin` with production registrations and actual SDK `ClientSession` | Real transport and error/result contracts with a fake backend, not actual fitting, rendering or persistence in Origin |

The [CI workflow](../.github/workflows/ci.yml) runs only on Ubuntu and skips the
Windows-only `originpro` dependency. The [pytest fixtures](../tests/conftest.py)
explicitly prevent importing the real backend. A Windows OS test result, Python
bridge import, code review, green CI badge or merged PR is not license/COM proof.

## Live checks still missing evidence

The [latest live-boundary report](result-contract-handoff.md#live-origin-boundary)
records no Origin installation/COM registration and no successful live run.
The PR body, conversation, reviews/review comments, tracked handoffs and workflow
jobs/artifacts were reviewed; no later live acceptance bundle was found. The
status below describes available evidence, not a claim that unrecorded work
could not have happened.

| Required acceptance | Status at this audit | Evidence needed |
| --- | --- | --- |
| Licensed attachment and import → fit → plot → export → save/reopen → release | NOT VERIFIED; latest handoff records NOT RUN | Live exit code/log, exact versions and tested SHA, retained `.opju`/PNG/CSV, visual inspection and reopened X=[1,2,3], Y=[3,5,7] |
| Windows client configuration loading | NOT VERIFIED | Preserved settings, client configuration readback and client `originlab_doctor` result |
| Failed save/open and required autosave gating | NOT VERIFIED | Unwritable/corrupt-file outcomes, blocked destructive operations and verified current-project consistency |
| Project A/B identity across release/reconnect and UI switches | NOT VERIFIED | Evidence that no save reuses the wrong cached destination, including unavailable/ambiguous metadata |
| Periodic saving while editing and idling | NOT VERIFIED | Actual save timestamps and timer/idle behavior on a disposable project |
| Modal soft timeout and lifecycle recovery | NOT VERIFIED | Owner/thread trace, expired queued work not executed, behavior of already-started work, detach/shutdown without force-killing Origin |

Follow the full [Windows acceptance checklist](windows-validation.md), using a
disposable licensed Origin session. Record PASS / FAIL / NOT RUN for every check
and attach a versioned evidence bundle here. A successful live smoke script alone
does not complete the manual categories or the P0 design work.

## Remaining P0 technical risks

1. **COM owner and lifecycle — OPEN.**
   [Dispatch](../src/originlab_mcp/utils/dispatch.py) creates a new executor per
   timed call; [manager lifecycle and timer callbacks](../src/originlab_mcp/origin_manager.py)
   also access the backend. The lock serializes operations but does not establish
   one COM owner/apartment across attachment, calls, release/reconnect, idle and
   shutdown. A soft timeout does not cancel work that already started. A bounded
   owner-thread design, deterministic tests and Windows validation are still needed.
2. **Preflight save plus destructive operation — OPEN.**
   `preflight_autosave()` and the mutation use separate `execute()` calls, visible
   in the manager and [worksheet mutations](../src/originlab_mcp/tools/data.py).
   Another server call or manual UI switch can intervene. Add competing-call
   tests and one owner dispatch for the pair; separately detect/handle external
   UI changes. Do not describe the current pair as transactional.
3. **Real Origin acceptance — NOT VERIFIED.**
   All live/manual categories above still need retained evidence. Authoritative
   live project identity and periodic-save behavior also remain P1 in the
   [backlog](follow-up.md); their outstanding Windows checks are part of this gate.

The existing rejection of expired queued work is covered by portable regressions.
It does not close COM ownership, cancel started calls or prove live recovery.

## Publication state

- [GitHub Releases](https://api.github.com/repos/garethbeaumo/originlab-mcp/releases?per_page=100)
  and [tag refs](https://api.github.com/repos/garethbeaumo/originlab-mcp/git/matching-refs/tags/)
  were both empty at the audit. The only tracked workflow is CI; no repository
  publishing workflow was found.
- [Package metadata](../pyproject.toml) says `0.2.1`, and the
  [changelog](../CHANGELOG.md) has both historical version headings and an
  `Unreleased` section. These are source labels, not proof of uploaded releases.
- The [PyPI package named `originlab-mcp`](https://pypi.org/project/originlab-mcp/1.0.0/)
  lists version `1.0.0` (2026-08-22) with publishing provenance for
  **`Yike-Ye/OriginLab-MCP`**, a different repository. It is not release or
  acceptance evidence for **`garethbeaumo/originlab-mcp`**.

No release from the audited main/PR #8 was established by these records. This
documentation update does not publish, tag, bump a version or close acceptance.
Do not publish or declare live acceptance on portable/fake evidence alone.
