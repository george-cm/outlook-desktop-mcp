# Code Review — fix/com-bridge-and-formatting

**Reviewed:** commits `23e7033`–`38db46d` on `main` (merged from `fix/com-bridge-and-formatting`)
**Date:** 2026-05-22
**Scope:** `src/outlook_desktop_mcp/com_bridge.py`, `utils/constants.py`, `utils/formatting.py`, `tests/test_com_bridge.py`, `tests/test_formatting.py`, `tests/test_constants.py`
**Reviewers:** 4 parallel analysis subagents + 3 validation subagents

---

## Summary

| Severity | Count |
|---|---|
| Blocking (AGENTS.md) | 9 |
| Bug | 3 |
| Total | 12 |

Server.py changes: clean. No findings.

---

## Blocking — AGENTS.md Violations

All violations: em dashes (U+2014) banned by AGENTS.md. Replace `—` with ` - ` everywhere.

### B1 — `utils/constants.py` lines 4, 6, 7, 22, 39, 44 (6 violations)

```
Line 4:  "Two distinct constant sets — do NOT mix them up:"
Line 6:  "olItemType — used with Outlook.CreateItem(n)"
Line 7:  "olObjectClass — used to verify item.Class"
Line 22: "# olFolderArchive — was missing, causing ..."
Line 39: "# --- olItemType — used with Outlook.CreateItem() ---"
Line 44: "# --- olObjectClass — used to verify item.Class ---"
```

Fix: replace all 6 `—` with ` - `.

### B2 — `com_bridge.py` lines 8, 33 (2 violations)

```
Line 8:  "as its first two arguments — the live COM objects"
Line 33: "Idempotent — safe to call multiple times."
```

Fix: replace both `—` with ` - `.

### B3 — `tests/test_com_bridge.py` line 3 (1 violation)

```
Line 3: "No Outlook or COM installation required — _com_thread_main is patched."
```

Fix: replace `—` with ` - `.

---

## Bugs

### BUG1 — `com_bridge.py` lines 111-118: TOCTOU race in `call()`

**Confidence: 92 (validated)**

`_alive.is_set()` is checked at line 111, then `queue.put()` happens at line 118. The COM thread can die in the gap — via crash or concurrent `stop()` call. If it does, the request sits in the dead queue forever → 60-second `TimeoutError` with a misleading "Outlook may be waiting for user input" message. The `call()` docstring explicitly promises "Raises RuntimeError immediately if bridge is not running or has crashed" — that promise is violated in this window.

Most exploitable scenario: `stop()` called concurrently with `call()` during service shutdown.

```python
# call() current (broken):
if not self._alive.is_set():           # line 111 — checked here
    raise RuntimeError(...)
result_event = threading.Event()       # line 116
result_holder: dict = {}               # line 117
self._request_queue.put(...)           # line 118 — COM thread can be dead by now
await loop.run_in_executor(...)        # line 121 — waits 60s
```

Fix options (pick one):
- Re-check `_alive.is_set()` after `queue.put()` and drain the item if dead
- Poll `_alive` during the wait instead of waiting the full 60s
- Use an `asyncio.Event` that the COM thread clears on exit, enabling cancellation

### BUG2 — `com_bridge.py` lines 37-40: stale request queue not drained on restart

**Confidence: 95 (validated)**

`start()` resets `_shutdown`, `_ready`, `_alive`, `_init_error` but never drains `self._request_queue` (created once in `__init__` line 24). If a `call()` timed out while the COM thread was blocked, the tuple `(func, args, kwargs, result_event, result_holder)` remains in the queue. On `stop()`+`start()`, the new COM thread immediately dequeues and executes it — silently re-sending an email, re-creating a calendar event, or any other mutation. The original caller's `result_event` is set but no one is listening.

This is a supported code path: `start()` docstring says "Idempotent — safe to call multiple times" and `test_restart_after_stop_call_works` explicitly tests the restart path.

```python
# start() missing drain (lines 37-40):
self._shutdown.clear()
self._ready.clear()
self._alive.clear()
self._init_error = None
# self._request_queue NOT cleared — stale requests survive restart
```

Fix:
```python
# Drain any leftover requests from a previous lifecycle.
import queue as _q
while True:
    try:
        self._request_queue.get_nowait()
    except _q.Empty:
        break
```
Add this block after the 4 reset lines in `start()`.

### BUG3 — `tests/test_formatting.py` line 164: `test_due_date_accessed_once` is a no-op

**Confidence: 97 (validated)**

The test name and docstring claim to verify `DueDate` is accessed only once (no double COM round-trip). The body only asserts `result["due_date"] is not None` — which passes whether `DueDate` is accessed once or ten times. Removing the `due = item.DueDate` cache line and writing two direct accesses (`item.DueDate.year`, `item.DueDate.isoformat()`) would still pass this test. The comment in the test admits as much: "We can't easily count property accesses on MagicMock, but we verify the result is correct."

Compare with `test_attachment_count_cached` (line 90) which correctly uses `PropertyMock` — the same pattern applies here.

Fix:
```python
def test_due_date_accessed_once(self):
    """DueDate must be cached - no double COM round-trip."""
    from unittest.mock import PropertyMock
    item = _mock_task(due_year=2026)
    due_mock = PropertyMock(return_value=_pyTime(datetime(2026, 6, 1)))
    type(item).DueDate = due_mock
    format_task_summary(item)
    assert due_mock.call_count == 1, (
        f"DueDate accessed {due_mock.call_count} times - expected 1 (cached)"
    )
```

---

## Not flagged (confirmed clean)

- `server.py` changes: `_get_item` helper, 13 call sites, `create_event`/`create_task` Move() fixes, `OUTLOOK_DATE_FMT`/`MAX_ITEM_COUNT` constant extraction, `logging.basicConfig` move — all correct
- `utils/formatting.py`: all `str()` → `.isoformat()` migrations complete, sentinel check `due.year == 4501` correct, `format_task_full.date_completed` guard correct
- `test_com_bridge.py`: `_fake_com_main` correctly simulates real thread (event ordering, exit cleanup), `PropertyMock` usage for attachment count correct, double-stop guard in fixture correct
- `test_constants.py`: all 5 assertions correct
- `SEND_USING_ACCOUNT_DISPID = 64209`: verified correct (`0xFAD1` in Outlook type library)
- Pre-existing em dashes in untouched `server.py` docstrings (lines 5, 105, 109, 174, 544, 631, 912, 1077): out of scope for this review
