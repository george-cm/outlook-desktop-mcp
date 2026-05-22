# Project Status — Resume Point

**Last worked:** 2026-05-22
**Branch:** `main` (clean, all work merged)
**Tests:** 29 passing (`tests/test_constants.py`, `tests/test_com_bridge.py`, `tests/test_formatting.py`)

---

## What is done

### Phase 1 — Bug fixes (COMPLETE)

All pre-implementation review bugs fixed and verified by two code review passes.

| File | What changed |
|---|---|
| `utils/constants.py` | New file replacing `tools/_folder_constants.py`. Adds `olFolderArchive=28`, `OL_CLASS_*`, `OUTLOOK_DATE_FMT`, `MAX_ITEM_COUNT`, `SEND_USING_ACCOUNT_DISPID` |
| `tools/` | Deleted (was a dead stub) |
| `com_bridge.py` | 6 lifecycle bugs fixed: idempotent start, restart, pre-flight checks, `_alive` event, traceback preservation, TOCTOU race, stale queue drain, thread snapshot |
| `utils/formatting.py` | ISO 8601 datetimes, locale-safe sentinel check, cached COM access |
| `server.py` | `Move()` bug fixed, `_get_item` helper, constants extracted, `logging.basicConfig` moved to `main()` |

---

## What is next

### Phase 2 — `outlook` CLI

**Plan file:** `docs/superpowers/plans/2026-05-22-outlook-cli.md`

**Start here — Task 1:**
```bash
git worktree add .worktrees/cli -b feat/outlook-cli
cd .worktrees/cli
uv run --with pytest --no-sync pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py -q
```
Expected: 29 passed. Then proceed task by task per the plan.

**What Plan 2 builds:**
- `src/outlook_desktop_mcp/service.py` - `OutlookService(bridge)` class, 30 methods
- `src/outlook_desktop_mcp/cli.py` - Typer CLI, 7 command groups, 30 commands
- `server.py` refactored to thin `@mcp.tool()` wrappers
- `pyproject.toml` - `outlook` entry point added

**Command tree (when done):**
```
outlook accounts
outlook folders list
outlook email {list,read,send,reply,move,mark-read,mark-unread,search,attachments,save-attachment}
outlook calendar {list,get,create,meeting,update,delete,respond,search}
outlook tasks {list,get,create,complete,delete}
outlook categories {list,set}
outlook rules {list,toggle}
outlook out-of-office
```

**Run method:** subagent-driven-development (one subagent per task, spec + quality review after each)

**Note on uv run:** MCP server process locks `.venv/Scripts/outlook-desktop-mcp.exe`.
If `uv run` fails with "process cannot access the file", use:
```bash
uv run --with pytest --no-sync pytest ...
```

---

## Key files

| File | Purpose |
|---|---|
| `docs/superpowers/specs/2026-05-22-outlook-cli-design.md` | Design decisions, architecture, command tree |
| `docs/superpowers/plans/2026-05-22-fix-com-bridge-formatting.md` | Plan 1 (complete) |
| `docs/superpowers/plans/2026-05-22-outlook-cli.md` | Plan 2 (next) |
| `docs/2026-05-22-cli-arch-review.md` | Pre-implementation review findings |
| `REVIEW.md` | Post-implementation code review findings (all resolved) |
