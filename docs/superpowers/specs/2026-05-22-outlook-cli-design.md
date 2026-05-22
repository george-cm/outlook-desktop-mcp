# Design: `outlook` CLI for outlook-desktop-mcp

**Date:** 2026-05-22
**Status:** Awaiting user review
**Author:** George Murga + Devin

---

## Summary

Add an `outlook` CLI command to the `outlook-desktop-mcp` project, exposing all 30 MCP
tools as shell-invocable subcommands. Primary consumers: AI coding agents (Devin, Claude,
Copilot), shell scripts/pipelines, and George directly. JSON output by default,
`--plain` for human-readable output.

---

## Design Decisions

| Decision | Choice | Rationale |
|---|---|---|
| CLI framework | Typer | Already installed via `mcp[cli]`. Type-annotated, auto-generates `--help`. |
| Async bridge | `asyncio.run()` per command | `anyio.run(lambda: coro)` is broken. stdlib, zero deps. |
| Output format | JSON default, `--plain` flag | Agents parse JSON; humans use `--plain`. |
| Command style | Grouped subcommands | `outlook email list`, `outlook calendar get`. Mirrors MCP categories, discoverable. |
| Service layer | `OutlookService(bridge)` class | Bridge injected, helpers private methods. `(outlook, namespace)` as explicit params on 20+ fns is manual `self`. Testable. |
| Platform scope | Windows-only first | macOS AppleScript path has separate schema drift issues. Follow-up work. |
| Bridge lifecycle | CLI owns its own instance | CLI must NOT import `server.py` (pulls in full FastMCP setup, 30 tool registrations). Each entrypoint owns `start()`/`stop()`. |
| Bug scope | Fix first, CLI second | Phase 1 fixes all review bugs (benefits MCP server too). Phase 2 builds on clean base. |

---

## Phases

### Phase 1 — `fix/com-bridge-and-formatting`

Fix all bugs found in the pre-implementation review. No new features. Every fix also
improves the existing MCP server.

#### `com_bridge.py`

| Bug | Fix |
|---|---|
| `start()` no idempotency guard (line 32) | `if self._thread is not None and self._thread.is_alive(): return` |
| `_shutdown` never cleared (line 106) | In `start()`: `self._shutdown.clear(); self._ready.clear(); self._init_error = None` |
| `call()` blind to dead bridge (line 82) | Pre-flight: check thread alive + `_init_error` before queuing |
| COM thread death undetectable | Add `_alive = threading.Event()`. Set before loop, clear in `finally`, check in `call()` |
| Traceback lost on re-raise (line 103) | Store `e.__traceback__`, re-raise with `.with_traceback(tb)` |
| Dead `import sys` (line 14) | Remove |

#### `utils/formatting.py`

| Bug | Fix |
|---|---|
| `str(item.ReceivedTime)` not ISO 8601 (line 32) | `.isoformat()` on the `PyTime` object |
| `str(item.DueDate) != "01/01/4501"` locale comparison (lines 93-94) | `item.DueDate.year == 4501` |
| `str(item.DueDate)` called twice (lines 93-94) | Cache: `d = item.DueDate` |
| `item.Attachments.Count` accessed twice (lines 34-35) | Cache: `att_count = item.Attachments.Count` |

#### `server.py`

| Issue | Fix |
|---|---|
| `create_event`/`create_task` discard `Move()` return (lines 944-947, 1407-1411) | `appt = appt.Move(cal)` |
| `olFolderArchive` missing from `FOLDER_NAME_TO_ENUM` | Add `"archive": 28` |
| `_OL_CLASS_MAIL/APPOINTMENT/TASK` in wrong file (lines 72-74) | Move to `utils/constants.py` with clear doc separating from `olItemType` constants |
| `OUTLOOK_DATE_FMT` hardcoded 10+ places | `OUTLOOK_DATE_FMT = "%m/%d/%Y %H:%M"` constant |
| `MAX_ITEM_COUNT = 200` hardcoded 5 places | `MAX_ITEM_COUNT = 200` constant |
| `_get_item` block copy-pasted 13 times | Extract `_get_item(namespace, entry_id, account="")` helper |
| Bare `except: continue` with no logging | `except Exception as e: logger.debug("Skipping item %d: %s", i+1, e)` |
| Magic number `64209` (lines 287, 1025) | `SEND_USING_ACCOUNT_DISPID = 64209  # Outlook SendUsingAccount DISPID` |

#### Structure cleanup

- Move `tools/_folder_constants.py` → `utils/constants.py`
- Delete `tools/` directory (empty `__init__.py` + now-moved constants file)
- Update 3 import sites: `formatting.py`, `server.py`, `server_mac.py`

---

### Phase 2 — `feat/outlook-cli` (rebased onto Phase 1)

#### File changes

| File | Change |
|---|---|
| `src/outlook_desktop_mcp/service.py` | **New** — `OutlookService` class |
| `src/outlook_desktop_mcp/cli.py` | **New** — Typer app, all commands |
| `src/outlook_desktop_mcp/server.py` | **Refactor** — thin `@mcp.tool()` wrappers calling `OutlookService` |
| `pyproject.toml` | Add `outlook` entry point, declare `typer` dep explicitly |

---

## Architecture

```
com_bridge.py  (OutlookBridge - fixed)
      |
service.py     (OutlookService(bridge) class - business logic)
      |                    |
server.py              cli.py
(@mcp.tool             (Typer grouped subcommands,
4-6 line wrappers)      asyncio.run() per command,
                        owns bridge lifecycle)
```

### `OutlookService` class

```python
class OutlookService:
    def __init__(self, bridge: OutlookBridge) -> None:
        self._bridge = bridge

    # Email
    async def list_emails(self, folder="inbox", count=10, ...) -> str: ...
    async def read_email(self, entry_id="", ...) -> str: ...
    async def send_email(self, to, subject, body, ...) -> str: ...
    async def reply_email(self, entry_id, body, ...) -> str: ...
    async def move_email(self, entry_id, target_folder="archive", ...) -> str: ...
    async def mark_as_read(self, entry_id, ...) -> str: ...
    async def mark_as_unread(self, entry_id, ...) -> str: ...
    async def search_emails(self, query, ...) -> str: ...
    async def list_attachments(self, entry_id, ...) -> str: ...
    async def save_attachment(self, entry_id, index, output_path, ...) -> str: ...

    # Calendar
    async def list_events(self, ...) -> str: ...
    async def get_event(self, entry_id, ...) -> str: ...
    async def create_event(self, subject, start, end, ...) -> str: ...
    async def create_meeting(self, subject, start, end, required_attendees, ...) -> str: ...
    async def update_event(self, entry_id, ...) -> str: ...
    async def delete_event(self, entry_id, ...) -> str: ...
    async def respond_to_meeting(self, entry_id, response, ...) -> str: ...
    async def search_events(self, query, ...) -> str: ...

    # Tasks
    async def list_tasks(self, ...) -> str: ...
    async def get_task(self, entry_id, ...) -> str: ...
    async def create_task(self, subject, ...) -> str: ...
    async def complete_task(self, entry_id, ...) -> str: ...
    async def delete_task(self, entry_id, ...) -> str: ...

    # Misc
    async def list_accounts(self) -> str: ...
    async def list_folders(self, ...) -> str: ...
    async def list_categories(self, ...) -> str: ...
    async def set_category(self, entry_id, category, ...) -> str: ...
    async def list_rules(self, ...) -> str: ...
    async def toggle_rule(self, name, enabled, ...) -> str: ...
    async def get_out_of_office(self, ...) -> str: ...

    # Private helpers (moved from server.py module level)
    def _get_item(self, namespace, entry_id, account=""): ...
    def _resolve_store(self, namespace, account=""): ...
    def _require_store(self, namespace, account=""): ...
    def _resolve_folder(self, namespace, folder_name, store=None): ...
    def _walk_folders(self, parent, name_lower): ...
    def _safe_dasl(self, query): ...
    def _check_item_class(self, item, expected_class, label): ...
```

Return type is always `str` (JSON) - matches existing MCP tool contract. CLI prints
directly to stdout. Errors go to stderr as JSON `{"error": "..."}` with non-zero exit
code.

### `server.py` after refactor (example)

```python
from outlook_desktop_mcp.service import OutlookService

bridge = OutlookBridge()
svc = OutlookService(bridge)

@mcp.tool()
async def list_emails(folder="inbox", count=10, ...) -> str:
    return await svc.list_emails(folder, count, ...)
```

### CLI bridge lifecycle

```python
# cli.py
app = typer.Typer()
_svc: OutlookService | None = None

@app.callback()
def main(ctx: typer.Context) -> None:
    bridge = OutlookBridge()
    bridge.start()
    ctx.obj = OutlookService(bridge)
    ctx.call_on_close(bridge.stop)
```

---

## Command Tree

```
outlook accounts
outlook folders list [--folder PATH] [--depth N] [--account A]

outlook email list    [--folder F] [--count N] [--unread] [--start DATE] [--end DATE] [--account A]
outlook email read    <entry-id> | --subject QUERY [--folder F] [--account A]
outlook email send    --to TO --subject SUBJ --body BODY [--cc CC] [--bcc BCC] [--html BODY] [--account A]
outlook email reply   <entry-id> --body BODY [--reply-all] [--account A]
outlook email move    <entry-id> [--folder archive] [--account A]
outlook email mark-read   <entry-id> [--account A]
outlook email mark-unread <entry-id> [--account A]
outlook email search  <query> [--folder F] [--count N] [--start DATE] [--end DATE] [--account A]
outlook email attachments <entry-id> [--account A]
outlook email save-attachment <entry-id> <index> --output PATH [--account A]

outlook calendar list     [--start DATE] [--end DATE] [--count N] [--account A]
outlook calendar get      <entry-id> [--account A]
outlook calendar create   --subject S --start DT --end DT [--location L] [--body B] [--all-day] [--reminder N] [--account A]
outlook calendar meeting  --subject S --start DT --end DT --required ADDRS [--optional ADDRS] [--location L] [--body B] [--account A]
outlook calendar update   <entry-id> [--subject S] [--start DT] [--end DT] [--location L] [--body B] [--account A]
outlook calendar delete   <entry-id> [--account A]
outlook calendar respond  <entry-id> --response accept|decline|tentative [--account A]
outlook calendar search   <query> [--start DATE] [--end DATE] [--count N] [--account A]

outlook tasks list     [--status STATUS] [--account A]
outlook tasks get      <entry-id> [--account A]
outlook tasks create   --subject S [--body B] [--due DATE] [--priority N] [--account A]
outlook tasks complete <entry-id> [--account A]
outlook tasks delete   <entry-id> [--account A]

outlook categories list [--account A]
outlook categories set  <entry-id> <category> [--account A]

outlook rules list   [--account A]
outlook rules toggle <name> --enable/--disable [--account A]

outlook out-of-office [--account A]
```

---

## Output Contract

### Success (stdout)

JSON string as returned by `OutlookService` methods (already JSON from existing logic):

```json
[{"entry_id": "...", "subject": "...", "sender": "...", ...}]
```

With `--plain`: human-readable table or summary (Typer/Rich rendering).

### Error (stderr + exit code 1)

```json
{"error": "Folder 'xyz' not found. Use `outlook folders list` to see available folders."}
```

Stdout is empty on error - safe to pipe.

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Success |
| `1` | Outlook/COM error, not found, validation failure |
| `2` | Bridge not running (Outlook not open) |

---

## Error Handling

CLI boundary catches all exceptions from `OutlookService`, writes JSON to stderr,
exits with appropriate code. Agents can check exit code without parsing stderr.

```python
try:
    result = asyncio.run(svc.list_emails(...))
    typer.echo(result)
except RuntimeError as e:
    typer.echo(json.dumps({"error": str(e)}), err=True)
    raise typer.Exit(2)
except Exception as e:
    typer.echo(json.dumps({"error": str(e)}), err=True)
    raise typer.Exit(1)
```

---

## Testing

### Phase 1 tests

- `tests/test_com_bridge.py` - mock the COM thread, test all 3 lifecycle fixes:
  - `start()` idempotency
  - `stop()` + `start()` restart
  - `call()` on dead bridge raises immediately, not after 60s
- `tests/test_formatting.py` - mock COM `PyTime` objects, verify ISO output and
  sentinel `None` on non-US locale strings

### Phase 2 tests

- `tests/test_service.py` - mock `OutlookBridge.call()`, test each `OutlookService`
  method returns expected JSON structure
- `tests/test_cli.py` - Typer `CliRunner`, test each command group:
  - Correct args passed to service
  - JSON written to stdout
  - Errors written to stderr with correct exit code

TDD: failing tests first, then implementation.

---

## `pyproject.toml` changes

```toml
[project.scripts]
outlook-desktop-mcp = "outlook_desktop_mcp.entrypoint:main"
outlook = "outlook_desktop_mcp.cli:main"

[project.dependencies]
# existing...
typer >= 0.9.0
```

---

## Out of Scope

- macOS AppleScript CLI support (separate follow-up)
- `--json` / `--plain` toggle per-field granularity
- Shell autocomplete setup (Typer supports it but not in scope for v1)
- `server_mac.py` schema drift fix
