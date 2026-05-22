# CLI Architecture Review — 2026-05-22

Pre-implementation audit conducted via 3 parallel reviewer subagents before adding the
`outlook` CLI to this project. Findings are ordered by severity.

---

## `com_bridge.py` — 3 Blocking Bugs

### 1. `start()` has no idempotency guard (line 32)

No `if self._thread is not None and self._thread.is_alive(): return` check.
Calling `start()` twice spawns two threads both consuming from the same `queue.Queue`.
COM STA objects are thread-bound. Objects created on thread A used on thread B =
undefined COM behaviour - silent corruption or access violations.

**Fix:** Add at the top of `start()`:
```python
if self._thread is not None and self._thread.is_alive():
    return
```

### 2. `_shutdown` never cleared after `stop()` (line 106)

`stop()` sets `self._shutdown`. It is never cleared. Calling `start()` after `stop()`
creates a new thread that initialises successfully, sets `_ready`, then immediately
exits the `while not self._shutdown.is_set()` loop. `start()` returns without error.
Every subsequent `call()` queues work that will never be consumed, hanging 60 seconds
before raising a `TimeoutError` with a misleading message about Outlook dialogs.

**Fix:** In `start()`, before launching the thread:
```python
self._shutdown.clear()
self._ready.clear()
self._init_error = None
```

### 3. `call()` is blind to a dead or unstarted bridge (line 82)

`call()` queues work unconditionally. If `start()` was never called, if the COM thread
crashed, or if `_init_error` is set, requests sit in the queue forever. The only signal
is a 60-second `TimeoutError` with a misleading message.

**Fix:** At the top of `call()`:
```python
if self._thread is None or not self._thread.is_alive():
    raise RuntimeError("COM bridge is not running. Call start() first.")
if self._init_error:
    raise self._init_error
```

### Additional issues

- **COM thread death mid-session undetectable.** If an exception escapes the `while`
  loop, the thread exits silently. `call()` still queues work. Add a `_alive`
  `threading.Event` - set before the loop, cleared in `finally`, checked in `call()`.
- **Traceback lost on re-raise (line 103).** `raise result_holder["error"]` discards
  the COM thread's original stack. Fix:
  ```python
  # In _com_thread_main except block:
  result_holder["tb"] = e.__traceback__
  # In call():
  raise result_holder["error"].with_traceback(result_holder["tb"])
  ```
- **`run_in_executor(None, ...)` exhausts the default thread pool under load.** Each
  pending `call()` occupies one worker for up to 60s. Default pool is ~12 workers
  (8-core machine). At 12 concurrent calls the pool saturates and new calls don't even
  dispatch their COM work yet.
- **`import sys` (line 14) is dead.** `sys` is never used in this file.

---

## `server.py` — Bugs and Anti-patterns

### Bugs

**`create_event` and `create_task` discard `Move()` return value (lines 944-947, 1407-1411)**

```python
appt.Move(cal)                               # return value DISCARDED
appt = namespace.GetItemFromID(appt.EntryID) # OLD, now-invalid EntryID
```

`Move()` returns the moved item with its **new** EntryID. The old EntryID is invalidated.
The code discards the return value and looks up by the old EntryID. This works by
accident (undocumented Outlook cross-reference) but will fail unpredictably.

**Fix:** `appt = appt.Move(cal)`

**`olFolderArchive` (28) missing from `FOLDER_NAME_TO_ENUM`**

`move_email` defaults to `target_folder="archive"`. This falls through the enum lookup
and triggers a full recursive DFS of the entire store. On localised Outlook installs the
folder name differs - silent failure with "Target folder 'archive' not found."

**Fix:** Add `"archive": 28` to `FOLDER_NAME_TO_ENUM` in `_folder_constants.py`.

**`save_attachment` does file I/O on the COM thread (lines 1565-1601)**

`os.makedirs()` and `att.SaveAsFile()` run on the single serial COM thread. Large
attachments on slow disks or network shares block all Outlook operations for the
duration.

**Fix:** Extract the file write to the async side after the COM call returns.

### Anti-patterns

**`_get_item` block copy-pasted 13 times**

```python
if account:
    store = _require_store(namespace, account)
    item = namespace.GetItemFromID(entry_id, store.StoreID)
else:
    item = namespace.GetItemFromID(entry_id)
```

Appears verbatim in: `mark_as_read`, `mark_as_unread`, `reply_email`, `get_event`,
`update_event`, `delete_event`, `respond_to_meeting`, `get_task`, `complete_task`,
`delete_task`, `list_attachments`, `save_attachment`, `set_category`.

**Fix:** Extract `_get_item(namespace, entry_id, account="") -> item`.

**`OUTLOOK_DATE_FMT` hardcoded in 10+ places (lines 362, 365, 368, 752, 756, 761, 841, 842, 1278, 1279)**

`'%m/%d/%Y %H:%M'` inline everywhere.

**Fix:** `OUTLOOK_DATE_FMT = "%m/%d/%Y %H:%M"`

**`MAX_ITEM_COUNT = 200` hardcoded in 5 places (lines 347, 738, 828, 1267, 1329)**

**Magic number `64209` (lines 287, 1025)**

`mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))` - the DISPID for `SendUsingAccount`.
Completely opaque.

**Fix:** `SEND_USING_ACCOUNT_DISPID = 64209  # Outlook SendUsingAccount property DISPID`

**`_check_item_class` returns a string instead of raising (lines 77-81)**

Callers use the walrus operator to check the string. Mixes data and control flow.
The error string is indistinguishable from a success result to any upstream caller.

**Fix:** Raise `TypeError` or a custom `WrongItemTypeError`.

**`_parse_date` is a pointless one-liner (line 790)**

```python
def _parse_date(date_str: str) -> datetime:
    return datetime.fromisoformat(date_str)
```

No error handling, no additional formats, zero added value.

**Fix:** Delete it and call `datetime.fromisoformat()` directly, or add real value
(try/except, multiple format support, user-friendly error message).

**Bare `except Exception: continue` with zero logging everywhere**

Lines 154, 187, 207, 380, 683, 693 (and more). Corrupt or unsupported items are silently
dropped. Users get fewer results than requested with no explanation.

**Fix:** At minimum: `except Exception as e: logger.debug("Skipping item %d: %s", i+1, e)`

**`_walk_folders` has no recursion depth limit (line 144)**

Unbounded recursive DFS. A deep Outlook folder tree hits Python's 1000-frame recursion
limit. The `RecursionError` propagates through `bridge.call()` and gets swallowed by
`format_com_error` as "An unexpected error occurred."

**`_OL_CLASS_*` constants live in `server.py` instead of `_folder_constants.py` (lines 72-74)**

```python
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48
```

Identical in nature to the constants already in `_folder_constants.py`. Note: these are
`olObjectClass` values (used for type-checking items), distinct from `OL_MAIL_ITEM = 0`
etc. which are `olItemType` values (used in `CreateItem()`). Both sets need to coexist
in `_folder_constants.py` with clear documentation.

---

## `utils/formatting.py` — Data Corruption Bugs

### Bug 1 — Locale-dependent datetime (line 32)

```python
"received_time": str(item.ReceivedTime),
```

`str()` on a COM `PyTime` object produces a locale-specific string:
- en-US: `"3/22/2026 2:00:00 PM"`
- de-DE: `"22.03.2026 14:00:00"`
- en-GB: `"22/03/2026 14:00:00"`

Not ISO 8601. Any downstream consumer parsing this string breaks on non-US Windows.
Same issue on `item.Start` (line 56) and `item.End` (line 57).

**Fix:** `item.ReceivedTime.isoformat()` - the COM `PyTime` exposes a standard Python
`datetime`.

### Bug 2 — Locale-dependent sentinel comparison (lines 93-94) — silent data corruption

```python
"due_date": str(item.DueDate) if str(item.DueDate) != "01/01/4501" else None,
"start_date": str(item.StartDate) if str(item.StartDate) != "01/01/4501" else None,
```

The `"01/01/4501"` sentinel is Outlook's "no date set" magic date. The comparison is
against the string representation, which is locale-dependent. On European machines the
sentinel is `"01.01.4501"`. The comparison is always `False`. Every task without a due
date emits the sentinel string as the value instead of `None`. **Silent data
corruption on any non-US Windows machine.**

**Fix:** Compare the COM object directly: `item.DueDate.year == 4501`

### Bug 3 — Double COM round-trip per task field (lines 93-94)

`str(item.DueDate)` evaluated twice per field - once for the comparison, once for the
value. Same on line 94. 50 tasks = 200 unnecessary COM calls from this one pattern.

**Fix:**
```python
d = item.DueDate
"due_date": None if d.year == 4501 else d.isoformat()
```

### Bug 4 — `item.Attachments` accessed twice (lines 34-35)

```python
"has_attachments": bool(item.Attachments.Count > 0),
"attachment_count": item.Attachments.Count,
```

Two COM round-trips for the same value.

**Fix:** `att_count = item.Attachments.Count`

### Structural issue — formatters are COM-coupled, not reusable

Every formatter directly dereferences COM dispatch attributes. `server_mac.py` cannot
use them. It reimplements the same dict shapes in 3 separate places (lines 201-211,
351-360, 500-513) - schema drift is already present.

`_truncate` in `server_mac.py` (line 69) is character-for-character identical to
`truncate` in `utils/formatting.py` (line 13).

---

## `tools/` Directory — Dead Stub

`tools/__init__.py` is zero bytes. `tools/_folder_constants.py` contains only constants.
The directory exists because a per-category tool module architecture was planned but
never built - everything ended up in a 1814-line `server.py` instead.

**Worse:** `utils/formatting.py` imports from `tools/_folder_constants.py`. `utils/` is
conceptually lower-level infrastructure. `tools/` is higher-level application layer.
This dependency direction is inverted.

**Fix:** Move `_folder_constants.py` to `utils/constants.py`. Update the 3 import sites
in `formatting.py`, `server.py`, and `server_mac.py`. Delete `tools/`.

---

## Architectural Implication for CLI

The original plan - "CLI calls `server.py` tool functions directly" - is wrong.
Importing `server.py` triggers `FastMCP(...)` setup, 30 `@mcp.tool()` registrations,
and `bridge = OutlookBridge()` instantiation. None of that belongs in a CLI invocation.

**Approved architecture:**

```
com_bridge.py  (fix 3 blocking bugs above)
      |
service.py     (extracted inner functions, all helpers at top, bugs fixed)
      |                    |
server.py              cli.py
(@mcp.tool             (typer commands,
4-6 line wrappers)      asyncio.run(),
                        owns bridge lifecycle)
```

`service.py` functions have signature `(outlook, namespace, ...) -> str`. No `bridge`,
no `async`, no `mcp` imports. Importable on any machine including macOS for testing.

`bridge` stays in `server.py`/`cli.py` as infrastructure - service.py has no business
knowing how COM dispatch works.

CLI bridge lifecycle (explicit, not imported global):
```python
def main():
    bridge = OutlookBridge()
    bridge.start()
    try:
        app()  # typer app
    finally:
        bridge.stop()
```

**Alternative worth evaluating:** `OutlookService(bridge: OutlookBridge)` class where
helpers are private methods and bridge is injected. Makes the dependency explicit and
testable. The current `(outlook, namespace)` parameter pattern on every function is
manual `self`.
