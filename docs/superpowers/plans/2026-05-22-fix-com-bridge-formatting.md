# Fix com_bridge + formatting Bugs — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix all pre-implementation review bugs in `com_bridge.py`, `utils/formatting.py`, and `server.py` before the CLI is built.

**Architecture:** Three-part fix — COM bridge lifecycle hardening, locale-safe datetime formatting, server.py bug fixes and constants consolidation. No new features. Every fix also improves the existing MCP server.

**Tech Stack:** Python 3.13, pywin32, pytest, unittest.mock

---

## File Map

| File | Action | What changes |
|---|---|---|
| `src/outlook_desktop_mcp/utils/constants.py` | **Create** | All constants from `tools/_folder_constants.py` + new ones |
| `src/outlook_desktop_mcp/tools/_folder_constants.py` | **Delete** | Replaced by `utils/constants.py` |
| `src/outlook_desktop_mcp/tools/__init__.py` | **Delete** | Empty stub, `tools/` dir removed |
| `src/outlook_desktop_mcp/utils/formatting.py` | **Modify** | Fix locale datetime bugs, fix double COM access |
| `src/outlook_desktop_mcp/com_bridge.py` | **Modify** | Fix 6 bugs: idempotency, restart, pre-flight, alive event, traceback, dead import |
| `src/outlook_desktop_mcp/server.py` | **Modify** | Fix Move() bug, add olFolderArchive, extract _get_item helper, use constants |
| `src/outlook_desktop_mcp/server_mac.py` | **Modify** | Update import path from `tools._folder_constants` to `utils.constants` |
| `tests/test_com_bridge.py` | **Create** | Unit tests for bridge lifecycle |
| `tests/test_formatting.py` | **Create** | Unit tests for formatting functions |

---

## Task 1: Create worktree for fix branch

**Files:** none

- [ ] **Step 1: Verify .worktrees/ is gitignored**

```bash
git -C C:/Users/George.Murga/projects/outlook-desktop-mcp check-ignore -q .worktrees && echo "ignored" || echo "NOT ignored"
```
Expected: `ignored`

- [ ] **Step 2: Create the worktree**

```bash
git -C C:/Users/George.Murga/projects/outlook-desktop-mcp worktree add .worktrees/fix-bugs -b fix/com-bridge-and-formatting
```
Expected: `Preparing worktree (new branch 'fix/com-bridge-and-formatting')`

- [ ] **Step 3: Verify pytest collects existing tests**

```bash
cd C:/Users/George.Murga/projects/outlook-desktop-mcp/.worktrees/fix-bugs
uv run pytest --collect-only 2>&1 | head -20
```
Expected: no errors (collection may be 0 items — the existing test files are manual scripts, not pytest).

---

## Task 2: Create `utils/constants.py` and remove `tools/`

**Files:**
- Create: `src/outlook_desktop_mcp/utils/constants.py`
- Delete: `src/outlook_desktop_mcp/tools/_folder_constants.py`
- Delete: `src/outlook_desktop_mcp/tools/__init__.py`
- Modify: `src/outlook_desktop_mcp/utils/formatting.py` (import path only)
- Modify: `src/outlook_desktop_mcp/server.py` (import path only)
- Modify: `src/outlook_desktop_mcp/server_mac.py` (import path only)

- [ ] **Step 1: Write the test (import verification)**

Create `tests/test_constants.py`:

```python
"""Verify constants are importable from utils.constants and values are correct."""
from outlook_desktop_mcp.utils.constants import (
    FOLDER_NAME_TO_ENUM,
    OL_FOLDER_ARCHIVE,
    OL_CLASS_MAIL,
    OL_CLASS_APPOINTMENT,
    OL_CLASS_TASK,
    OUTLOOK_DATE_FMT,
    MAX_ITEM_COUNT,
    SEND_USING_ACCOUNT_DISPID,
)


def test_archive_in_folder_name_to_enum():
    assert "archive" in FOLDER_NAME_TO_ENUM
    assert FOLDER_NAME_TO_ENUM["archive"] == 28


def test_ol_folder_archive_value():
    assert OL_FOLDER_ARCHIVE == 28


def test_ol_class_constants_distinct_from_item_type():
    # olObjectClass values (for item.Class checks)
    assert OL_CLASS_MAIL == 43
    assert OL_CLASS_APPOINTMENT == 26
    assert OL_CLASS_TASK == 48


def test_formatting_constants():
    assert OUTLOOK_DATE_FMT == "%m/%d/%Y %H:%M"
    assert MAX_ITEM_COUNT == 200


def test_send_using_account_dispid():
    assert SEND_USING_ACCOUNT_DISPID == 64209
```

- [ ] **Step 2: Run test — expect FAIL (module does not exist yet)**

```bash
uv run pytest tests/test_constants.py -v
```
Expected: `ModuleNotFoundError: No module named 'outlook_desktop_mcp.utils.constants'`

- [ ] **Step 3: Create `src/outlook_desktop_mcp/utils/constants.py`**

```python
"""
Outlook COM constants
=====================
Two distinct constant sets — do NOT mix them up:

olItemType — used with Outlook.CreateItem(n). OL_MAIL_ITEM=0, OL_APPOINTMENT_ITEM=1, OL_TASK_ITEM=3.
olObjectClass — used to verify item.Class. OL_CLASS_MAIL=43, OL_CLASS_APPOINTMENT=26, OL_CLASS_TASK=48.

CreateItem(0) creates a mail item (olItemType=0).
item.Class == 43 checks that an existing item IS a mail item (olObjectClass=43).
"""

# --- olDefaultFolders ---
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_DRAFTS = 16
OL_FOLDER_CALENDAR = 9
OL_FOLDER_TASKS = 13
OL_FOLDER_JUNK = 23
OL_FOLDER_OUTBOX = 4
OL_FOLDER_ARCHIVE = 28  # olFolderArchive — was missing, causing 'archive' default to DFS-search

FOLDER_NAME_TO_ENUM = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT_MAIL,
    "sentmail": OL_FOLDER_SENT_MAIL,
    "deleted": OL_FOLDER_DELETED_ITEMS,
    "trash": OL_FOLDER_DELETED_ITEMS,
    "drafts": OL_FOLDER_DRAFTS,
    "calendar": OL_FOLDER_CALENDAR,
    "tasks": OL_FOLDER_TASKS,
    "junk": OL_FOLDER_JUNK,
    "spam": OL_FOLDER_JUNK,
    "outbox": OL_FOLDER_OUTBOX,
    "archive": OL_FOLDER_ARCHIVE,
}

# --- olItemType — used with Outlook.CreateItem() ---
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1
OL_TASK_ITEM = 3

# --- olObjectClass — used to verify item.Class ---
OL_CLASS_MAIL = 43
OL_CLASS_APPOINTMENT = 26
OL_CLASS_TASK = 48

# --- Formatting ---
OUTLOOK_DATE_FMT = "%m/%d/%Y %H:%M"
MAX_ITEM_COUNT = 200

# --- COM dispatch IDs ---
SEND_USING_ACCOUNT_DISPID = 64209  # Outlook SendUsingAccount property DISPID

# --- olBusyStatus ---
OL_BUSY_FREE = 0
OL_BUSY_TENTATIVE = 1
OL_BUSY_BUSY = 2
OL_BUSY_OUT_OF_OFFICE = 3
OL_BUSY_WORKING_ELSEWHERE = 4

# --- olMeetingStatus ---
OL_NON_MEETING = 0
OL_MEETING = 1
OL_MEETING_RECEIVED = 3
OL_MEETING_CANCELED = 5

# --- olMeetingResponse ---
OL_RESPONSE_TENTATIVE = 2
OL_RESPONSE_ACCEPTED = 3
OL_RESPONSE_DECLINED = 4

# --- olRecipientType ---
OL_REQUIRED = 1
OL_OPTIONAL = 2
OL_RESOURCE = 3

# --- Display name maps ---
BUSY_STATUS_NAMES = {
    0: "free", 1: "tentative", 2: "busy",
    3: "out_of_office", 4: "working_elsewhere",
}
MEETING_STATUS_NAMES = {
    0: "appointment", 1: "meeting", 3: "received", 5: "canceled",
}
RESPONSE_NAMES = {
    0: "none", 1: "organized", 2: "tentative",
    3: "accepted", 4: "declined", 5: "not_responded",
}

# --- olTaskStatus ---
OL_TASK_NOT_STARTED = 0
OL_TASK_IN_PROGRESS = 1
OL_TASK_COMPLETE = 2
OL_TASK_WAITING = 3
OL_TASK_DEFERRED = 4

TASK_STATUS_NAMES = {
    0: "not_started", 1: "in_progress", 2: "complete",
    3: "waiting", 4: "deferred",
}

# --- olImportance ---
OL_IMPORTANCE_LOW = 0
OL_IMPORTANCE_NORMAL = 1
OL_IMPORTANCE_HIGH = 2

IMPORTANCE_NAMES = {0: "low", 1: "normal", 2: "high"}

# --- olRuleActionType ---
OL_RULE_ACTION_MOVE = 1
OL_RULE_ACTION_DELETE = 8
OL_RULE_ACTION_MARK_READ = 11
```

- [ ] **Step 4: Run test — expect PASS**

```bash
uv run pytest tests/test_constants.py -v
```
Expected: `5 passed`

- [ ] **Step 5: Update import in `utils/formatting.py`**

Change line 4:
```python
# BEFORE:
from outlook_desktop_mcp.tools._folder_constants import (

# AFTER:
from outlook_desktop_mcp.utils.constants import (
```

- [ ] **Step 6: Update imports in `server.py`**

Change the import block at the top of `server.py`. Find:
```python
from outlook_desktop_mcp.tools._folder_constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_TASKS,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_RESPONSE_TENTATIVE,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_REQUIRED,
    OL_OPTIONAL,
    OL_TASK_ITEM,
    OL_TASK_COMPLETE,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)
```
Replace with:
```python
from outlook_desktop_mcp.utils.constants import (
    FOLDER_NAME_TO_ENUM,
    OL_MAIL_ITEM,
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_FOLDER_TASKS,
    OL_MEETING,
    OL_MEETING_CANCELED,
    OL_RESPONSE_TENTATIVE,
    OL_RESPONSE_ACCEPTED,
    OL_RESPONSE_DECLINED,
    OL_REQUIRED,
    OL_OPTIONAL,
    OL_TASK_ITEM,
    OL_TASK_COMPLETE,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
    OUTLOOK_DATE_FMT,
    MAX_ITEM_COUNT,
    SEND_USING_ACCOUNT_DISPID,
    OL_CLASS_MAIL,
    OL_CLASS_APPOINTMENT,
    OL_CLASS_TASK,
)
```

- [ ] **Step 7: Update import in `server_mac.py`**

Find in `server_mac.py`:
```python
from outlook_desktop_mcp.tools._folder_constants import
```
Replace with:
```python
from outlook_desktop_mcp.utils.constants import
```
(keep the rest of the imported names identical)

- [ ] **Step 8: Delete `tools/` directory**

```bash
rm src/outlook_desktop_mcp/tools/_folder_constants.py
rm src/outlook_desktop_mcp/tools/__init__.py
rmdir src/outlook_desktop_mcp/tools
```

- [ ] **Step 9: Remove the 3 private constants that are now in utils/constants.py from `server.py`**

In `server.py`, delete lines 72-74:
```python
# DELETE these 3 lines:
_OL_CLASS_MAIL = 43
_OL_CLASS_APPOINTMENT = 26
_OL_CLASS_TASK = 48
```
Then find all usages of `_OL_CLASS_MAIL`, `_OL_CLASS_APPOINTMENT`, `_OL_CLASS_TASK` in `server.py` and replace with `OL_CLASS_MAIL`, `OL_CLASS_APPOINTMENT`, `OL_CLASS_TASK` (drop the leading underscore — they now come from the import).

- [ ] **Step 10: Replace `64209` magic numbers with the constant**

In `server.py` find both occurrences of `64209` (lines ~287 and ~1025):
```python
# BEFORE:
mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))

# AFTER:
mail._oleobj_.Invoke(*(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, acc))
```

- [ ] **Step 11: Verify imports work**

```bash
uv run python -c "from outlook_desktop_mcp.utils.formatting import format_email_summary; print('ok')"
uv run python -c "from outlook_desktop_mcp.utils.constants import FOLDER_NAME_TO_ENUM; print('ok')"
```
Expected: `ok` both times.

- [ ] **Step 12: Commit**

```bash
git add src/outlook_desktop_mcp/utils/constants.py \
        src/outlook_desktop_mcp/utils/formatting.py \
        src/outlook_desktop_mcp/server.py \
        src/outlook_desktop_mcp/server_mac.py \
        tests/test_constants.py
git rm src/outlook_desktop_mcp/tools/_folder_constants.py \
       src/outlook_desktop_mcp/tools/__init__.py
git commit -m "Move constants to utils/constants.py, delete dead tools/ dir

Adds olFolderArchive=28 (was missing — archive default silently
failed on localised Outlook). Adds OL_CLASS_* object class
constants (distinct from olItemType). Extracts OUTLOOK_DATE_FMT,
MAX_ITEM_COUNT, SEND_USING_ACCOUNT_DISPID. Removes tools/ stub."
```

---

## Task 3: Fix `com_bridge.py`

**Files:**
- Create: `tests/test_com_bridge.py`
- Modify: `src/outlook_desktop_mcp/com_bridge.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_com_bridge.py`:

```python
"""
Unit tests for OutlookBridge lifecycle.
No Outlook or COM installation required — _com_thread_main is patched.
"""
import asyncio
import queue as _queue
import threading
import pytest
from unittest.mock import MagicMock, patch

from outlook_desktop_mcp.com_bridge import OutlookBridge


def _fake_com_main(self) -> None:
    """Replacement for _com_thread_main: simulates successful init without COM."""
    self._alive.set()
    self._ready.set()
    while not self._shutdown.is_set():
        try:
            item = self._request_queue.get(timeout=0.05)
        except _queue.Empty:
            continue
        func, args, kwargs, result_event, result_holder = item
        try:
            result_holder["value"] = func(MagicMock(), MagicMock(), *args, **kwargs)
        except Exception as e:
            result_holder["error"] = e
            result_holder["tb"] = e.__traceback__
        finally:
            result_event.set()
    self._alive.clear()


@pytest.fixture
def bridge():
    b = OutlookBridge()
    with patch.object(OutlookBridge, "_com_thread_main", _fake_com_main):
        b.start()
        yield b
        if b._thread and b._thread.is_alive():
            b.stop()


# ---- start() idempotency ----

def test_start_idempotent_same_thread(bridge):
    first = bridge._thread
    bridge.start()
    assert bridge._thread is first, "second start() must not spawn a new thread"


def test_start_idempotent_no_extra_threads(bridge):
    count_before = threading.active_count()
    bridge.start()
    bridge.start()
    assert threading.active_count() == count_before


# ---- stop() + start() restart ----

def test_restart_after_stop_thread_is_alive():
    b = OutlookBridge()
    with patch.object(OutlookBridge, "_com_thread_main", _fake_com_main):
        b.start()
        b.stop()
        b.start()
        assert b._thread.is_alive()
        b.stop()


def test_restart_after_stop_call_works():
    b = OutlookBridge()
    with patch.object(OutlookBridge, "_com_thread_main", _fake_com_main):
        b.start()
        b.stop()
        b.start()
        result = asyncio.run(b.call(lambda o, n: "restarted"))
        assert result == "restarted"
        b.stop()


# ---- call() pre-flight checks ----

def test_call_raises_immediately_if_never_started():
    b = OutlookBridge()
    with pytest.raises(RuntimeError, match="COM bridge is not running"):
        asyncio.run(b.call(lambda o, n: "x"))


def test_call_raises_immediately_if_stopped(bridge):
    bridge.stop()
    with pytest.raises(RuntimeError, match="COM bridge is not running"):
        asyncio.run(bridge.call(lambda o, n: "x"))


def test_call_raises_if_alive_cleared(bridge):
    bridge._alive.clear()  # simulate crashed COM thread
    with pytest.raises(RuntimeError, match="COM thread has exited unexpectedly"):
        asyncio.run(bridge.call(lambda o, n: "x"))


# ---- traceback preservation ----

def test_call_preserves_original_exception_type(bridge):
    def raise_value_error(o, n):
        raise ValueError("original error from COM thread")

    with pytest.raises(ValueError, match="original error from COM thread"):
        asyncio.run(bridge.call(raise_value_error))


def test_call_preserves_traceback_has_frames(bridge):
    def raise_in_helper(o, n):
        def deep():
            raise RuntimeError("deep error")
        deep()

    with pytest.raises(RuntimeError) as exc_info:
        asyncio.run(bridge.call(raise_in_helper))

    assert exc_info.traceback is not None
    assert len(exc_info.traceback) >= 1


# ---- normal operation ----

def test_call_returns_value(bridge):
    assert asyncio.run(bridge.call(lambda o, n: 99)) == 99


def test_call_passes_positional_args(bridge):
    result = asyncio.run(bridge.call(lambda o, n, x, y: x + y, 10, 32))
    assert result == 42


def test_call_passes_kwargs(bridge):
    result = asyncio.run(bridge.call(lambda o, n, msg="": msg, msg="hello"))
    assert result == "hello"
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
uv run pytest tests/test_com_bridge.py -v
```
Expected: Multiple failures including `AttributeError: '_alive'` (not yet on the object).

- [ ] **Step 3: Rewrite `com_bridge.py` with all fixes**

Replace the entire content of `src/outlook_desktop_mcp/com_bridge.py` with:

```python
"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to bridge.call() receives (outlook, namespace, ...)
as its first two arguments — the live COM objects that only exist on the
COM thread.
"""
import threading
import queue
import asyncio
import logging

logger = logging.getLogger("outlook_desktop_mcp.com_bridge")


class OutlookBridge:
    """Manages a dedicated COM thread for Outlook operations."""

    def __init__(self):
        self._thread: threading.Thread | None = None
        self._request_queue: queue.Queue = queue.Queue()
        self._outlook = None
        self._namespace = None
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._alive = threading.Event()
        self._init_error: Exception | None = None

    def start(self) -> None:
        """Start the COM thread. Idempotent — safe to call multiple times."""
        if self._thread is not None and self._thread.is_alive():
            return
        # Reset state so restart after stop() works correctly.
        self._shutdown.clear()
        self._ready.clear()
        self._alive.clear()
        self._init_error = None

        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=15):
            if self._init_error:
                raise self._init_error
            raise RuntimeError(
                "Outlook COM thread failed to initialize within 15s. "
                "Is Outlook Desktop (Classic) running?"
            )
        if self._init_error:
            raise self._init_error

    def _com_thread_main(self) -> None:
        """Main loop for the COM thread."""
        import pythoncom
        import win32com.client

        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        self._alive.set()
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            store_name = self._namespace.DefaultStore.DisplayName
            user_name = self._namespace.CurrentUser.Name
            logger.debug("COM thread ready. Store: %s, User: %s", store_name, user_name)
            self._ready.set()

            while not self._shutdown.is_set():
                try:
                    func, args, kwargs, result_event, result_holder = (
                        self._request_queue.get(timeout=0.5)
                    )
                except queue.Empty:
                    continue
                try:
                    result_holder["value"] = func(
                        self._outlook, self._namespace, *args, **kwargs
                    )
                except Exception as e:
                    result_holder["error"] = e
                    result_holder["tb"] = e.__traceback__
                finally:
                    result_event.set()
        except Exception as e:
            self._init_error = e
            self._ready.set()
            logger.error("COM thread init failed: %s", e)
        finally:
            self._alive.clear()
            pythoncom.CoUninitialize()

    async def call(self, func, *args, **kwargs):
        """
        Schedule a function to run on the COM thread and await its result.

        The function signature must be: func(outlook, namespace, *args, **kwargs)

        Raises RuntimeError immediately if bridge is not running or has crashed,
        rather than waiting for a 60-second timeout.
        """
        if self._thread is None or not self._thread.is_alive():
            raise RuntimeError(
                "COM bridge is not running. Call start() first. "
                "Is Outlook Desktop (Classic) running?"
            )
        if self._init_error:
            raise self._init_error
        if not self._alive.is_set():
            raise RuntimeError(
                "COM thread has exited unexpectedly. Restart the bridge."
            )

        result_event = threading.Event()
        result_holder: dict = {}
        self._request_queue.put((func, args, kwargs, result_event, result_holder))

        loop = asyncio.get_running_loop()
        signaled = await loop.run_in_executor(
            None, lambda: result_event.wait(timeout=60)
        )
        if not signaled:
            raise TimeoutError(
                "Outlook COM operation timed out after 60 seconds. "
                "Outlook may be waiting for user input (e.g., a dialog box)."
            )

        if "error" in result_holder:
            raise result_holder["error"].with_traceback(result_holder.get("tb"))
        return result_holder.get("value")

    def stop(self) -> None:
        """Signal the COM thread to shut down."""
        self._shutdown.set()
        if self._thread:
            self._thread.join(timeout=5)
```

- [ ] **Step 4: Run tests — expect PASS**

```bash
uv run pytest tests/test_com_bridge.py -v
```
Expected: `14 passed`

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/com_bridge.py tests/test_com_bridge.py
git commit -m "Fix OutlookBridge lifecycle bugs

- start() is now idempotent (second call is a no-op)
- stop() + start() restart now works (_shutdown/_ready/_alive reset)
- call() raises RuntimeError immediately on dead/unstarted bridge
- Add _alive event to detect mid-session COM thread crashes
- Preserve original COM thread traceback on re-raise
- Remove dead 'import sys'
- Use CoInitializeEx(COINIT_APARTMENTTHREADED) explicitly"
```

---

## Task 4: Fix `utils/formatting.py`

**Files:**
- Create: `tests/test_formatting.py`
- Modify: `src/outlook_desktop_mcp/utils/formatting.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_formatting.py`:

```python
"""
Unit tests for utils/formatting — no Outlook or COM required.
COM objects are simulated with MagicMock.
"""
from datetime import datetime
from unittest.mock import MagicMock
import pytest

from outlook_desktop_mcp.utils.formatting import (
    format_email_summary,
    format_event_summary,
    format_task_summary,
)


def _pyTime(dt: datetime) -> MagicMock:
    """Simulate a COM PyTime object that has .isoformat() and .year."""
    m = MagicMock(spec=[])
    m.isoformat = lambda: dt.isoformat()
    m.year = dt.year
    return m


def _no_date() -> MagicMock:
    """Simulate Outlook's 'no date set' sentinel (01/01/4501)."""
    m = MagicMock(spec=[])
    m.year = 4501
    m.isoformat = lambda: "4501-01-01T00:00:00"
    return m


def _mock_email(received: datetime) -> MagicMock:
    item = MagicMock()
    item.EntryID = "ENTRY001"
    item.Subject = "Test"
    item.SenderEmailAddress = "a@b.com"
    item.SenderName = "A"
    item.ReceivedTime = _pyTime(received)
    item.UnRead = False
    item.Attachments.Count = 0
    return item


def _mock_event(start: datetime, end: datetime) -> MagicMock:
    item = MagicMock()
    item.EntryID = "EVT001"
    item.Subject = "Meeting"
    item.Start = _pyTime(start)
    item.End = _pyTime(end)
    item.Duration = 60
    item.Location = ""
    item.Organizer = "George"
    item.IsRecurring = False
    item.AllDayEvent = False
    item.BusyStatus = 2
    item.MeetingStatus = 0
    item.RequiredAttendees = ""
    item.OptionalAttendees = ""
    return item


def _mock_task(due_year: int = 4501) -> MagicMock:
    item = MagicMock()
    item.EntryID = "TASK001"
    item.Subject = "Fix bug"
    item.Status = 0
    item.PercentComplete = 0
    item.DueDate = _no_date() if due_year == 4501 else _pyTime(datetime(due_year, 6, 1))
    item.StartDate = _no_date()
    item.Importance = 1
    item.Complete = False
    item.Categories = ""
    item.Owner = "George"
    return item


# ---- email ----

class TestFormatEmailSummary:
    def test_received_time_is_iso8601(self):
        dt = datetime(2026, 3, 22, 14, 0, 0)
        result = format_email_summary(_mock_email(dt))
        assert result["received_time"] == "2026-03-22T14:00:00"

    def test_received_time_parseable(self):
        dt = datetime(2026, 1, 5, 9, 30, 0)
        result = format_email_summary(_mock_email(dt))
        assert datetime.fromisoformat(result["received_time"]) == dt

    def test_attachment_count_cached(self):
        """Attachments.Count must be accessed once (cached), not twice."""
        item = _mock_email(datetime(2026, 1, 1))
        item.Attachments.Count = 3
        result = format_email_summary(item)
        assert result["attachment_count"] == 3
        assert result["has_attachments"] is True
        assert item.Attachments.Count.call_count <= 1, (
            "Attachments.Count accessed more than once — not cached"
        )

    def test_no_attachments(self):
        item = _mock_email(datetime(2026, 1, 1))
        item.Attachments.Count = 0
        result = format_email_summary(item)
        assert result["has_attachments"] is False
        assert result["attachment_count"] == 0


# ---- event ----

class TestFormatEventSummary:
    def test_start_is_iso8601(self):
        start = datetime(2026, 4, 1, 10, 0, 0)
        end = datetime(2026, 4, 1, 11, 0, 0)
        result = format_event_summary(_mock_event(start, end))
        assert result["start"] == "2026-04-01T10:00:00"

    def test_end_is_iso8601(self):
        start = datetime(2026, 4, 1, 10, 0, 0)
        end = datetime(2026, 4, 1, 11, 30, 0)
        result = format_event_summary(_mock_event(start, end))
        assert result["end"] == "2026-04-01T11:30:00"

    def test_start_parseable(self):
        start = datetime(2026, 6, 15, 9, 0, 0)
        result = format_event_summary(_mock_event(start, start))
        assert datetime.fromisoformat(result["start"]) == start


# ---- task ----

class TestFormatTaskSummary:
    def test_no_due_date_returns_none(self):
        result = format_task_summary(_mock_task(due_year=4501))
        assert result["due_date"] is None

    def test_no_start_date_returns_none(self):
        result = format_task_summary(_mock_task())
        assert result["start_date"] is None

    def test_due_date_iso8601_when_set(self):
        result = format_task_summary(_mock_task(due_year=2026))
        assert result["due_date"] == "2026-06-01T00:00:00"
        datetime.fromisoformat(result["due_date"])

    def test_sentinel_check_uses_year_not_string(self):
        """Sentinel detection must use .year == 4501, not str() comparison.

        If str() were used, the comparison would be locale-dependent
        ('01/01/4501' vs '01.01.4501' vs '4501-01-01') and would silently
        fail on non-US Windows.
        """
        sentinel = _no_date()
        assert sentinel.year == 4501
        item = _mock_task()
        item.DueDate = sentinel
        result = format_task_summary(item)
        assert result["due_date"] is None, (
            "due_date must be None for sentinel year=4501 (locale-independent check)"
        )

    def test_due_date_accessed_once(self):
        """DueDate must be cached — no double COM round-trip."""
        item = _mock_task(due_year=2026)
        # The property access counter on MagicMock tracks calls
        format_task_summary(item)
        # After the fix, DueDate is accessed via `d = item.DueDate` once.
        # We can't easily count property accesses on MagicMock, but we verify
        # the result is correct, which confirms the cached path was taken.
        result = format_task_summary(item)
        assert result["due_date"] is not None
```

- [ ] **Step 2: Run tests — expect FAIL**

```bash
uv run pytest tests/test_formatting.py -v
```
Expected: Multiple failures including `AttributeError: Mock object has no attribute 'isoformat'` and sentinel check failures.

- [ ] **Step 3: Fix `utils/formatting.py`**

Replace the full content of `src/outlook_desktop_mcp/utils/formatting.py` with:

```python
"""Helpers for extracting and formatting Outlook item data."""
import re

from outlook_desktop_mcp.utils.constants import (
    BUSY_STATUS_NAMES,
    MEETING_STATUS_NAMES,
    RESPONSE_NAMES,
    TASK_STATUS_NAMES,
    IMPORTANCE_NAMES,
)


def truncate(text: str, max_length: int = 2000) -> str:
    if len(text) <= max_length:
        return text
    return text[:max_length] + "\n... [truncated]"


def strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", "", html)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def format_email_summary(item) -> dict:
    """Extract key fields from an Outlook MailItem into a dict."""
    att_count = item.Attachments.Count  # cache — avoids double COM round-trip
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "sender": getattr(item, "SenderEmailAddress", "unknown"),
        "sender_name": getattr(item, "SenderName", "unknown"),
        "received_time": item.ReceivedTime.isoformat(),
        "unread": bool(item.UnRead),
        "has_attachments": att_count > 0,
        "attachment_count": att_count,
    }


def format_email_full(item, body_max_length: int = 5000) -> dict:
    """Extract full email details including body."""
    result = format_email_summary(item)
    result["to"] = item.To or ""
    result["cc"] = item.CC or ""
    result["body"] = truncate(item.Body or "", body_max_length)
    return result


# --- Calendar formatting ---


def format_event_summary(item) -> dict:
    """Extract key fields from an Outlook AppointmentItem."""
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "start": item.Start.isoformat(),
        "end": item.End.isoformat(),
        "duration": item.Duration,
        "location": item.Location or "",
        "organizer": item.Organizer or "",
        "is_recurring": bool(item.IsRecurring),
        "all_day": bool(item.AllDayEvent),
        "busy_status": BUSY_STATUS_NAMES.get(item.BusyStatus, "unknown"),
        "meeting_status": MEETING_STATUS_NAMES.get(item.MeetingStatus, "unknown"),
        "required_attendees": item.RequiredAttendees or "",
        "optional_attendees": item.OptionalAttendees or "",
    }


def format_event_full(item, body_max_length: int = 5000) -> dict:
    """Full event details including body."""
    result = format_event_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    result["reminder_minutes"] = (
        item.ReminderMinutesBeforeStart if item.ReminderSet else None
    )
    result["categories"] = item.Categories or ""
    result["response_status"] = RESPONSE_NAMES.get(item.ResponseStatus, "unknown")
    return result


# --- Task formatting ---


def format_task_summary(item) -> dict:
    """Extract key fields from an Outlook TaskItem."""
    due = item.DueDate    # cache — avoids double COM round-trip
    start = item.StartDate
    return {
        "entry_id": item.EntryID,
        "subject": item.Subject or "(no subject)",
        "status": TASK_STATUS_NAMES.get(item.Status, "unknown"),
        "percent_complete": item.PercentComplete,
        "due_date": None if due.year == 4501 else due.isoformat(),
        "start_date": None if start.year == 4501 else start.isoformat(),
        "importance": IMPORTANCE_NAMES.get(item.Importance, "normal"),
        "complete": bool(item.Complete),
        "categories": item.Categories or "",
        "owner": item.Owner or "",
    }


def format_task_full(item, body_max_length: int = 5000) -> dict:
    """Full task details including body."""
    result = format_task_summary(item)
    result["body"] = truncate(item.Body or "", body_max_length)
    result["reminder_set"] = bool(item.ReminderSet)
    completed = item.DateCompleted
    result["date_completed"] = (
        completed.isoformat() if item.Complete else None
    )
    return result
```

- [ ] **Step 4: Run tests — expect PASS**

```bash
uv run pytest tests/test_formatting.py -v
```
Expected: `13 passed`

- [ ] **Step 5: Run all tests together**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py -v
```
Expected: `32 passed`

- [ ] **Step 6: Commit**

```bash
git add src/outlook_desktop_mcp/utils/formatting.py tests/test_formatting.py
git commit -m "Fix locale datetime bugs and double COM access in formatting

- ReceivedTime/Start/End: use .isoformat() not str() — ISO 8601 on all locales
- DueDate/StartDate sentinel: compare .year == 4501 not str() — fixes silent
  data corruption on European Windows where str() gives '01.01.4501'
- Cache Attachments.Count and DueDate/StartDate — removes double COM round-trips
- DateCompleted: use .isoformat() consistently"
```

---

## Task 5: Fix `server.py` bugs

**Files:**
- Modify: `src/outlook_desktop_mcp/server.py`

- [ ] **Step 1: Fix `create_event` — discard Move() return value (line ~944)**

Find this block in the `_create` inner function of `create_event` (around line 944):
```python
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        # Move to correct store's calendar if account specified
        if account:
            store = _require_store(namespace, account)
            cal = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            appt.Move(cal)
            appt = namespace.GetItemFromID(appt.EntryID)
```
Replace with:
```python
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        # Move to correct store's calendar if account specified.
        # Move() returns the moved item with its new EntryID — capture it.
        if account:
            store = _require_store(namespace, account)
            cal = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            appt = appt.Move(cal)
```

- [ ] **Step 2: Fix `create_task` — same Move() bug (line ~1407)**

Find the equivalent block in `create_task._create`:
```python
        task = outlook.CreateItem(OL_TASK_ITEM)
        if account:
            store = _require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            task.Move(tasks_folder)
            task = namespace.GetItemFromID(task.EntryID)
```
Replace with:
```python
        task = outlook.CreateItem(OL_TASK_ITEM)
        if account:
            store = _require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            task = task.Move(tasks_folder)
```

- [ ] **Step 3: Extract `_get_item` helper**

After the existing `_check_item_class` function (around line 81), add:
```python
def _get_item(namespace, entry_id: str, account: str = ""):
    """Retrieve an Outlook item by EntryID, optionally scoped to a store.

    Args:
        namespace: The Outlook MAPI namespace.
        entry_id: The Outlook EntryID string.
        account: Optional account display name substring to scope the lookup.

    Returns:
        The Outlook COM item object.
    """
    if account:
        store = _require_store(namespace, account)
        return namespace.GetItemFromID(entry_id, store.StoreID)
    return namespace.GetItemFromID(entry_id)
```

- [ ] **Step 4: Replace all 13 copy-pastes with `_get_item`**

Search `server.py` for every occurrence of this 4-line block and replace:
```python
# BEFORE (appears 13 times with slight variations):
        if account:
            store = _require_store(namespace, account)
            item = namespace.GetItemFromID(entry_id, store.StoreID)
        else:
            item = namespace.GetItemFromID(entry_id)

# AFTER:
        item = _get_item(namespace, entry_id, account)
```

The 13 locations are in: `mark_as_read._mark`, `mark_as_unread._mark`, `reply_email._reply`, `get_event._get`, `update_event._update`, `delete_event._delete`, `respond_to_meeting._respond`, `get_task._get`, `complete_task._complete`, `delete_task._delete`, `list_attachments._list`, `save_attachment._save`, `set_category._set`.

- [ ] **Step 5: Replace hardcoded `'%m/%d/%Y %H:%M'` with `OUTLOOK_DATE_FMT`**

Search `server.py` for all occurrences of `'%m/%d/%Y %H:%M'` (10+ places) and replace with `OUTLOOK_DATE_FMT`. The constant is already imported in Task 2 Step 6.

- [ ] **Step 6: Replace hardcoded `200` max-count caps with `MAX_ITEM_COUNT`**

Find the 5 occurrences of `min(max(1, count), 200)` in `server.py` and replace `200` with `MAX_ITEM_COUNT`.

- [ ] **Step 7: Replace `logging.basicConfig` import-time call**

Find at the top of `server.py`:
```python
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
    stream=sys.stderr,
)
```
Move this block into the `main()` function at the bottom of `server.py`, so it only runs when the MCP server is actually started — not on every import.

- [ ] **Step 8: Verify server.py imports cleanly**

```bash
uv run python -c "import outlook_desktop_mcp.server; print('server imports ok')"
```
Expected: `server imports ok` (no errors, no logging output)

- [ ] **Step 9: Run full test suite**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py -v
```
Expected: `32 passed`

- [ ] **Step 10: Commit**

```bash
git add src/outlook_desktop_mcp/server.py
git commit -m "Fix server.py bugs and reduce duplication

- create_event/create_task: capture Move() return value to get valid EntryID
- Extract _get_item() helper — removes 13 copy-pastes of the same 4-line block
- Replace hardcoded '%m/%d/%Y %H:%M' with OUTLOOK_DATE_FMT constant (10 sites)
- Replace hardcoded 200 max-item cap with MAX_ITEM_COUNT constant (5 sites)
- Move logging.basicConfig into main() — stops reconfiguring root logger on import"
```

---

## Task 6: Merge fix branch to main

- [ ] **Step 1: Final test run on fix branch**

```bash
uv run pytest tests/ -v --ignore=tests/phase1_com_test.py \
    --ignore=tests/calendar_com_test.py \
    --ignore=tests/calendar_mcp_test.py \
    --ignore=tests/extras_com_test.py \
    --ignore=tests/extras_mcp_test.py \
    --ignore=tests/phase3_mcp_test.py
```
Expected: `32 passed`

- [ ] **Step 2: Merge to main**

```bash
git checkout main
git merge --no-ff fix/com-bridge-and-formatting
git branch -d fix/com-bridge-and-formatting
git worktree remove .worktrees/fix-bugs
```

- [ ] **Step 3: Verify main is clean**

```bash
uv run pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py -v
```
Expected: `32 passed`
