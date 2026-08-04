# Draft-safe mode + COM reconnect resilience Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to
> implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for
> tracking.

**Goal:** Make `send_email`/`reply_email`/`create_meeting` draft-only (opt-in via
env var) instead of transmitting immediately, and make the COM bridge survive
"Outlook not running at startup" and "Outlook restarted mid-session" instead of
permanently breaking.

**Architecture:** `com_bridge.py` gets lazy Outlook connection (attempted on
first queued call, not at thread startup) plus disconnect-HRESULT detection
with a single reconnect-and-retry. `server.py` gets a frozen `DRAFT_MODE` flag
read from `OUTLOOK_MCP_DRAFT_MODE` at import time, and a `confirm_send: bool =
False` parameter on the 3 gated tools that branches `.Save()` vs `.Send()`.

**Tech Stack:** Python 3.12, pywin32 (`pythoncom`, `win32com.client`), pytest,
`unittest.mock`.

**Spec:** `docs/superpowers/specs/2026-08-04-draft-mode-and-reconnect-design.md`

**Working directory for every command below:**
`C:\Users\George.Murga\projects\outlook-desktop-mcp\.worktrees\draft-mode`

---

## Part 1 — COM reconnect resilience (`com_bridge.py`)

### Task 1: `_connect()` — dispatch/reconnect helper

**Files:**
- Modify: `src/outlook_desktop_mcp/com_bridge.py`
- Test: `tests/test_com_bridge.py`

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_com_bridge.py`, after the existing imports:

```python
from unittest.mock import patch
```

Add these tests at the end of the file:

```python
# ---- _connect() ----

def test_connect_dispatches_outlook_and_namespace():
    b = OutlookBridge()
    fake_outlook = MagicMock()
    fake_namespace = MagicMock()
    fake_outlook.GetNamespace.return_value = fake_namespace
    with patch("win32com.client.Dispatch", return_value=fake_outlook) as mock_dispatch:
        b._connect()
    mock_dispatch.assert_called_once_with("Outlook.Application")
    assert b._outlook is fake_outlook
    assert b._namespace is fake_namespace


def test_connect_raises_clear_error_when_outlook_not_running():
    b = OutlookBridge()
    with patch("win32com.client.Dispatch", side_effect=RuntimeError("boom")):
        with pytest.raises(RuntimeError, match="Outlook Desktop is not running"):
            b._connect()
    assert b._outlook is None
    assert b._namespace is None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_connect -v`

Expected: both `FAIL` with `AttributeError: 'OutlookBridge' object has no attribute '_connect'`.

- [ ] **Step 3: Write minimal implementation**

In `src/outlook_desktop_mcp/com_bridge.py`, add this method to `OutlookBridge`,
directly after `_com_thread_main` (i.e. after its closing `pythoncom.CoUninitialize()`
line, before `async def call`):

```python
    def _connect(self) -> None:
        """(Re)connect to a running Outlook instance. Must run on the COM thread.

        Raises RuntimeError with an actionable message if Outlook isn't
        running or the connection otherwise fails - never leaves a half-set
        _outlook/_namespace pair behind.
        """
        import win32com.client

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as e:
            self._outlook = None
            self._namespace = None
            raise RuntimeError(
                "Outlook Desktop is not running. Open it and retry."
            ) from e
        self._outlook = outlook
        self._namespace = namespace
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_connect -v`

Expected: `2 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/com_bridge.py tests/test_com_bridge.py
git commit -m "Add OutlookBridge._connect() reconnect helper"
```

---

### Task 2: `_is_disconnect_error()` — HRESULT classifier

**Files:**
- Modify: `src/outlook_desktop_mcp/com_bridge.py`
- Test: `tests/test_com_bridge.py`

- [ ] **Step 1: Write the failing tests**

Add to the top of `tests/test_com_bridge.py` (with the other imports):

```python
import pythoncom
```

Add these tests:

```python
# ---- _is_disconnect_error() ----

def test_is_disconnect_error_true_for_rpc_unavailable():
    b = OutlookBridge()
    e = pythoncom.com_error(-2147023174, "RPC server is unavailable.", None, None)  # 0x800706BA
    assert b._is_disconnect_error(e) is True


def test_is_disconnect_error_false_for_unrelated_com_error():
    b = OutlookBridge()
    e = pythoncom.com_error(-2147467259, "Unspecified error.", None, None)  # 0x80004005 E_FAIL
    assert b._is_disconnect_error(e) is False


def test_is_disconnect_error_false_for_non_com_error():
    b = OutlookBridge()
    assert b._is_disconnect_error(ValueError("x")) is False
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_is_disconnect_error -v`

Expected: all 3 `FAIL` with `AttributeError: 'OutlookBridge' object has no attribute '_is_disconnect_error'`.

- [ ] **Step 3: Write minimal implementation**

Add to `OutlookBridge`, directly after `_connect()`:

```python
    # HRESULTs indicating the COM connection itself is dead (Outlook was
    # closed/restarted), as opposed to an ordinary operation error.
    _DISCONNECT_HRESULTS = {
        0x800706BA,  # RPC server unavailable
        0x800706BE,  # RPC call failed
        0x80010108,  # RPC_E_DISCONNECTED
        0x800401FD,  # CO_E_OBJNOTCONNECTED
    }

    def _is_disconnect_error(self, e: Exception) -> bool:
        """True if e is a COM error indicating the Outlook connection died."""
        import pythoncom

        if not isinstance(e, pythoncom.com_error):
            return False
        hr = e.args[0] & 0xFFFFFFFF
        return hr in self._DISCONNECT_HRESULTS
```

Note: `_DISCONNECT_HRESULTS` is a class attribute (shared, immutable set), not
set in `__init__`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_is_disconnect_error -v`

Expected: `3 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/com_bridge.py tests/test_com_bridge.py
git commit -m "Add OutlookBridge._is_disconnect_error() HRESULT classifier"
```

---

### Task 3: `_run_with_reconnect()` — lazy connect + retry orchestration

**Files:**
- Modify: `src/outlook_desktop_mcp/com_bridge.py`
- Test: `tests/test_com_bridge.py`

- [ ] **Step 1: Write the failing tests (lazy connect + reuse)**

Add these tests:

```python
# ---- _run_with_reconnect() ----

def test_run_with_reconnect_connects_lazily_on_first_call():
    b = OutlookBridge()
    fake_outlook = MagicMock()
    fake_outlook.GetNamespace.return_value = MagicMock()
    with patch("win32com.client.Dispatch", return_value=fake_outlook) as mock_dispatch:
        result = b._run_with_reconnect(lambda o, n: "ok", (), {})
    assert result == "ok"
    mock_dispatch.assert_called_once()


def test_run_with_reconnect_reuses_existing_connection():
    b = OutlookBridge()
    b._outlook, b._namespace = MagicMock(), MagicMock()
    with patch("win32com.client.Dispatch") as mock_dispatch:
        result = b._run_with_reconnect(lambda o, n: "ok", (), {})
    assert result == "ok"
    mock_dispatch.assert_not_called()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_run_with_reconnect -v`

Expected: both `FAIL` with `AttributeError: 'OutlookBridge' object has no attribute '_run_with_reconnect'`.

- [ ] **Step 3: Write minimal implementation (lazy connect only)**

Add to `OutlookBridge`, directly after `_is_disconnect_error()`:

```python
    def _run_with_reconnect(self, func, args, kwargs):
        """Run func(outlook, namespace, *args, **kwargs) on the COM thread.

        Connects lazily if not yet connected. If the call fails with a
        disconnect-class COM error (Outlook was restarted), reconnects once
        and retries the same call. Any other error - including a failed
        reconnect - propagates to the caller unchanged.
        """
        if self._outlook is None:
            self._connect()
        return func(self._outlook, self._namespace, *args, **kwargs)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_run_with_reconnect -v`

Expected: `2 passed`.

- [ ] **Step 5: Write the failing tests (reconnect-and-retry)**

Add these tests:

```python
def test_run_with_reconnect_reconnects_and_retries_on_disconnect_error():
    b = OutlookBridge()
    b._outlook, b._namespace = MagicMock(), MagicMock()  # already "connected"
    disconnect_error = pythoncom.com_error(-2147023174, "RPC server is unavailable.", None, None)
    calls = {"n": 0}

    def flaky(outlook, namespace):
        calls["n"] += 1
        if calls["n"] == 1:
            raise disconnect_error
        return "recovered"

    fresh_outlook = MagicMock()
    fresh_outlook.GetNamespace.return_value = MagicMock()
    with patch("win32com.client.Dispatch", return_value=fresh_outlook) as mock_dispatch:
        result = b._run_with_reconnect(flaky, (), {})
    assert result == "recovered"
    assert calls["n"] == 2
    mock_dispatch.assert_called_once()
    assert b._outlook is fresh_outlook


def test_run_with_reconnect_surfaces_error_when_reconnect_fails():
    b = OutlookBridge()
    b._outlook, b._namespace = MagicMock(), MagicMock()
    disconnect_error = pythoncom.com_error(-2147023174, "RPC server is unavailable.", None, None)

    def always_disconnected(outlook, namespace):
        raise disconnect_error

    with patch("win32com.client.Dispatch", side_effect=RuntimeError("Outlook still down")):
        with pytest.raises(RuntimeError, match="Outlook Desktop is not running"):
            b._run_with_reconnect(always_disconnected, (), {})
    assert b._outlook is None


def test_run_with_reconnect_does_not_reconnect_on_non_disconnect_error():
    b = OutlookBridge()
    b._outlook, b._namespace = MagicMock(), MagicMock()
    original_outlook = b._outlook

    def raise_value_error(outlook, namespace):
        raise ValueError("bad entry_id")

    with patch("win32com.client.Dispatch") as mock_dispatch:
        with pytest.raises(ValueError, match="bad entry_id"):
            b._run_with_reconnect(raise_value_error, (), {})
    mock_dispatch.assert_not_called()
    assert b._outlook is original_outlook
```

- [ ] **Step 6: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_run_with_reconnect -v`

Expected: `test_run_with_reconnect_reconnects_and_retries_on_disconnect_error` and
`test_run_with_reconnect_surfaces_error_when_reconnect_fails` `FAIL` (the
disconnect_error propagates immediately instead of triggering a reconnect;
`test_does_not_reconnect_on_non_disconnect_error` already `PASS`es since the
minimal implementation never reconnects at all yet).

- [ ] **Step 7: Write the full implementation**

Replace the `_run_with_reconnect` body written in Step 3 with:

```python
    def _run_with_reconnect(self, func, args, kwargs):
        """Run func(outlook, namespace, *args, **kwargs) on the COM thread.

        Connects lazily if not yet connected. If the call fails with a
        disconnect-class COM error (Outlook was restarted), reconnects once
        and retries the same call. Any other error - including a failed
        reconnect - propagates to the caller unchanged.
        """
        if self._outlook is None:
            self._connect()
        try:
            return func(self._outlook, self._namespace, *args, **kwargs)
        except Exception as e:
            if not self._is_disconnect_error(e):
                raise
            logger.warning("Outlook COM connection lost, reconnecting: %s", e)
            self._outlook = None
            self._namespace = None
            self._connect()
            return func(self._outlook, self._namespace, *args, **kwargs)
```

- [ ] **Step 8: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k test_run_with_reconnect -v`

Expected: `5 passed`.

- [ ] **Step 9: Commit**

```bash
git add src/outlook_desktop_mcp/com_bridge.py tests/test_com_bridge.py
git commit -m "Add OutlookBridge._run_with_reconnect() lazy connect + retry"
```

---

### Task 4: Wire into `_com_thread_main()` and `start()`

**Files:**
- Modify: `src/outlook_desktop_mcp/com_bridge.py:32-100` (see current line numbers below - re-check with `grep -n "def start\|def _com_thread_main" src/outlook_desktop_mcp/com_bridge.py` before editing since Tasks 1-3 shifted line numbers)
- Test: `tests/test_com_bridge.py`

This is the integration step: it proves `start()` no longer requires Outlook
to be running, and that `bridge.call()` connects lazily end-to-end through a
real (not mocked) `_com_thread_main` loop.

- [ ] **Step 1: Write the failing tests**

Add these tests (these use the REAL `_com_thread_main`, not the `_fake_com_main`
fixture - do not use the `bridge` fixture for these):

```python
# ---- start() no longer requires Outlook running; call() connects lazily ----

def test_start_succeeds_without_outlook_running():
    """start() must not eagerly dispatch Outlook.Application."""
    b = OutlookBridge()
    with patch("win32com.client.Dispatch", side_effect=RuntimeError("Outlook not running")) as mock_dispatch:
        b.start()  # must not raise
    try:
        assert b._thread.is_alive()
        mock_dispatch.assert_not_called()
    finally:
        b.stop()


def test_call_connects_lazily_then_succeeds():
    b = OutlookBridge()
    fake_outlook = MagicMock()
    fake_outlook.GetNamespace.return_value = MagicMock()
    with patch("win32com.client.Dispatch", return_value=fake_outlook) as mock_dispatch:
        b.start()
        try:
            result = asyncio.run(b.call(lambda o, n: "connected"))
        finally:
            b.stop()
    assert result == "connected"
    mock_dispatch.assert_called_once()


def test_call_raises_clear_error_when_outlook_not_running_at_call_time():
    b = OutlookBridge()
    with patch("win32com.client.Dispatch", side_effect=RuntimeError("boom")):
        b.start()
        try:
            with pytest.raises(RuntimeError, match="Outlook Desktop is not running"):
                asyncio.run(b.call(lambda o, n: "x"))
        finally:
            b.stop()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -k "test_start_succeeds_without_outlook_running or test_call_connects_lazily_then_succeeds or test_call_raises_clear_error_when_outlook_not_running_at_call_time" -v`

Expected: `test_start_succeeds_without_outlook_running` `FAIL`s (current `start()`
calls `Dispatch` eagerly inside `_com_thread_main` and raises before
`b.start()` returns). The other two currently `PASS` against the old code by
coincidence of the old eager-connect design - re-verify after Step 3 that they
still pass for the *new* reason (lazy connect), not the old one.

- [ ] **Step 3: Write the implementation**

In `src/outlook_desktop_mcp/com_bridge.py`, replace the `start()` method body with:

```python
    def start(self) -> None:
        """Start the COM thread. Idempotent - safe to call multiple times.

        Only waits for the thread itself to come alive - NOT for a
        connection to Outlook. Connecting to Outlook happens lazily on the
        first queued call (see _run_with_reconnect), so start() succeeds
        even if Outlook isn't running yet.
        """
        if self._thread is not None and self._thread.is_alive():
            return
        # Reset state so restart after stop() works correctly.
        self._shutdown.clear()
        self._ready.clear()
        self._alive.clear()
        self._init_error = None
        # A fresh thread means a fresh COM apartment - any cached Outlook
        # proxy from a previous thread is invalid, force a reconnect.
        self._outlook = None
        self._namespace = None
        # Drain any stale requests left over from the previous lifecycle.
        # Without this, timed-out requests would re-execute on the new thread.
        while True:
            try:
                self._request_queue.get_nowait()
            except queue.Empty:
                break

        self._thread = threading.Thread(
            target=self._com_thread_main, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=15):
            if self._init_error:
                raise self._init_error
            raise RuntimeError("Outlook COM thread failed to start within 15s.")
        if self._init_error:
            raise self._init_error
```

Replace the `_com_thread_main()` method body with:

```python
    def _com_thread_main(self) -> None:
        """Main loop for the COM thread.

        Does NOT connect to Outlook here. Connecting is lazy - attempted by
        _run_with_reconnect() on the first queued call, and re-attempted if
        the connection dies mid-session (e.g. Outlook restarts). This lets
        the thread - and the MCP server around it - come up even if Outlook
        isn't running yet.
        """
        import pythoncom

        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception as e:
            self._init_error = e
            self._ready.set()
            return

        self._alive.set()
        self._ready.set()
        try:
            while not self._shutdown.is_set():
                try:
                    func, args, kwargs, result_event, result_holder = (
                        self._request_queue.get(timeout=0.5)
                    )
                except queue.Empty:
                    continue
                try:
                    result_holder["value"] = self._run_with_reconnect(func, args, kwargs)
                except Exception as e:
                    result_holder["error"] = e
                    result_holder["tb"] = e.__traceback__
                finally:
                    result_event.set()
        finally:
            self._alive.clear()
            pythoncom.CoUninitialize()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_com_bridge.py -v`

Expected: every test in the file passes (existing 29 + new ones from Tasks 1-4).

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/com_bridge.py tests/test_com_bridge.py
git commit -m "Make COM connection lazy: start() no longer requires Outlook running"
```

---

## Part 2 — Draft-safe mode (`server.py`)

### Task 5: `DRAFT_MODE` flag + `send_email` gating

**Files:**
- Modify: `src/outlook_desktop_mcp/server.py:122` (insert flag), `src/outlook_desktop_mcp/server.py:261-315` (send_email)
- Test: `tests/test_send_gating.py` (new)

- [ ] **Step 1: Write the failing tests**

Create `tests/test_send_gating.py`:

```python
"""
Unit tests for draft-safe mode gating on send_email, reply_email, create_meeting.
No Outlook or COM installation required - bridge.call() is replaced with a
synchronous passthrough over MagicMock outlook/namespace objects.
"""
import asyncio
from unittest.mock import MagicMock

import pytest

from outlook_desktop_mcp import server
from outlook_desktop_mcp.utils.constants import OL_CLASS_MAIL


@pytest.fixture
def fake_com(monkeypatch):
    """Replace bridge.call with a synchronous passthrough over fake COM objects."""
    fake_outlook = MagicMock()
    fake_namespace = MagicMock()
    fake_outlook.Session.Accounts = []

    async def _fake_call(func, *args, **kwargs):
        return func(fake_outlook, fake_namespace, *args, **kwargs)

    monkeypatch.setattr(server.bridge, "call", _fake_call)
    return fake_outlook, fake_namespace


# ---- send_email ----

def test_send_email_sends_immediately_when_draft_mode_off(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", False)
    fake_outlook, _ = fake_com
    asyncio.run(server.send_email(to="alice@example.com", subject="s", body="b"))
    mail = fake_outlook.CreateItem.return_value
    mail.Send.assert_called_once()
    mail.Save.assert_not_called()


def test_send_email_drafts_when_draft_mode_on(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    result = asyncio.run(server.send_email(to="alice@example.com", subject="s", body="b"))
    mail = fake_outlook.CreateItem.return_value
    mail.Save.assert_called_once()
    mail.Send.assert_not_called()
    assert "Draft saved" in result


def test_send_email_sends_when_draft_mode_on_and_confirm_send_true(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    asyncio.run(server.send_email(
        to="alice@example.com", subject="s", body="b", confirm_send=True,
    ))
    mail = fake_outlook.CreateItem.return_value
    mail.Send.assert_called_once()
    mail.Save.assert_not_called()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_send_gating.py -v`

Expected: `test_send_email_sends_immediately_when_draft_mode_off` currently
`PASS`es by coincidence (today's code always sends). The other two `FAIL`:
`AttributeError: <module 'outlook_desktop_mcp.server'> does not have the
attribute 'DRAFT_MODE'`.

- [ ] **Step 3: Write minimal implementation**

In `src/outlook_desktop_mcp/server.py`, after the line `bridge = OutlookBridge()`
(currently line 122), add:

```python

# Draft-safe mode: when enabled, send_email/reply_email/create_meeting create
# a reviewable Outlook artifact (Drafts item / unsent meeting) instead of
# transmitting immediately, unless the caller passes confirm_send=True.
# Off by default so behavior is unchanged for existing consumers.
DRAFT_MODE = os.getenv("OUTLOOK_MCP_DRAFT_MODE", "").strip().lower() in ("1", "true", "yes")
```

Replace the entire `send_email` tool (currently lines 261-315) with:

```python
@mcp.tool()
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    bcc: str = "",
    html_body: str = "",
    account: str = "",
    confirm_send: bool = False,
) -> str:
    """Send an email using the user's Outlook account.

    Creates and sends an email immediately through the default Outlook profile.
    The email will appear in the user's Sent Items folder after sending.

    When the server is running in draft-safe mode (OUTLOOK_MCP_DRAFT_MODE=1),
    this creates a Drafts item instead of sending. Pass confirm_send=true
    ONLY if the user has explicitly asked, in this turn, to send/dispatch
    immediately rather than draft it. Ignored (has no effect) when the
    server is not running in draft-safe mode.

    Args:
        to: One or more recipient email addresses, separated by semicolons.
            Example: "alice@example.com" or "alice@example.com; bob@example.com"
        subject: The email subject line.
        body: The plain-text body of the email. If html_body is also provided,
            both are set and Outlook will prefer the HTML version.
        cc: Optional. CC recipients, separated by semicolons.
        bcc: Optional. BCC recipients, separated by semicolons.
        html_body: Optional. HTML-formatted body. When provided, Outlook renders
            the email as HTML. The plain-text body serves as fallback.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.
        confirm_send: Optional. Only meaningful in draft-safe mode - set true
            to actually send instead of drafting. Default false.

    Returns:
        A confirmation message with subject and recipients, or an error.
    """
    def _send(outlook, namespace, to, subject, body, cc, bcc, html_body, account, confirm_send):
        store = _require_store(namespace, account)
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        # Set the sending account
        for acc in outlook.Session.Accounts:
            if acc.DeliveryStore.StoreID == store.StoreID:
                mail._oleobj_.Invoke(*(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, acc))  # SendUsingAccount
                break
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        if html_body:
            mail.HTMLBody = html_body
        if DRAFT_MODE and not confirm_send:
            mail.Save()
            return f"Draft saved: '{subject}' to {to} (entry_id={mail.EntryID})"
        mail.Send()
        return f"Email sent: '{subject}' to {to}"

    try:
        return await bridge.call(_send, to, subject, body, cc, bcc, html_body, account, confirm_send)
    except Exception as e:
        return f"Error sending email: {format_com_error(e)}"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_send_gating.py -v`

Expected: `3 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/server.py tests/test_send_gating.py
git commit -m "Add OUTLOOK_MCP_DRAFT_MODE gating to send_email"
```

---

### Task 6: `reply_email` gating (threaded reply draft)

**Files:**
- Modify: `src/outlook_desktop_mcp/server.py:582-619`
- Test: `tests/test_send_gating.py`

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_send_gating.py`:

```python
# ---- reply_email ----

def test_reply_email_sends_immediately_when_draft_mode_off(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", False)
    fake_outlook, fake_namespace = fake_com
    source = fake_namespace.GetItemFromID.return_value
    source.Class = OL_CLASS_MAIL
    source.Reply.return_value.Body = "original quoted body"
    asyncio.run(server.reply_email(entry_id="AAA", body="my note"))
    reply = source.Reply.return_value
    reply.Send.assert_called_once()
    reply.Save.assert_not_called()
    assert reply.Body == "my note\n\noriginal quoted body"


def test_reply_email_drafts_via_reply_never_createitem(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, fake_namespace = fake_com
    source = fake_namespace.GetItemFromID.return_value
    source.Class = OL_CLASS_MAIL
    source.Reply.return_value.Body = "original quoted body"
    result = asyncio.run(server.reply_email(entry_id="AAA", body="my note"))
    reply = source.Reply.return_value
    reply.Save.assert_called_once()
    reply.Send.assert_not_called()
    fake_outlook.CreateItem.assert_not_called()
    assert "Reply draft saved" in result


def test_reply_email_reply_all_uses_replyall(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, fake_namespace = fake_com
    source = fake_namespace.GetItemFromID.return_value
    source.Class = OL_CLASS_MAIL
    source.ReplyAll.return_value.Body = "original quoted body"
    asyncio.run(server.reply_email(entry_id="AAA", body="note", reply_all=True))
    source.ReplyAll.assert_called_once()
    source.ReplyAll.return_value.Save.assert_called_once()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_send_gating.py -k reply_email -v`

Expected: `test_reply_email_sends_immediately_when_draft_mode_off` `PASS`es by
coincidence (today's code always sends). The other two `FAIL` because
`reply_item.Save()` is never called and the message never contains "Reply draft saved".

- [ ] **Step 3: Write minimal implementation**

Replace the entire `reply_email` tool (currently lines 582-619) with:

```python
@mcp.tool()
async def reply_email(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    account: str = "",
    confirm_send: bool = False,
) -> str:
    """Reply to an email in Outlook.

    Creates a reply, preserving the original message thread. Use
    reply_all=True to reply to all recipients (sender + CC list).

    When the server is running in draft-safe mode (OUTLOOK_MCP_DRAFT_MODE=1),
    this creates a threaded Drafts item instead of sending - it is always
    built via Outlook's own Reply()/ReplyAll() on the source message, so it
    is guaranteed to land in the same conversation thread. Pass
    confirm_send=true ONLY if the user has explicitly asked, in this turn,
    to send/dispatch immediately rather than draft it. Ignored (has no
    effect) when the server is not running in draft-safe mode.

    Args:
        entry_id: The unique Outlook EntryID of the email to reply to.
        body: The reply message text. Prepended above the original message
            in the email thread.
        reply_all: If true, reply to all recipients (sender + all CC/To).
            If false (default), reply only to the sender.
        account: Optional. Account display name (or substring). Only needed
            if entry_id is ambiguous across stores.
        confirm_send: Optional. Only meaningful in draft-safe mode - set true
            to actually send instead of drafting. Default false.

    Returns:
        Confirmation indicating the reply was sent or drafted, or an error.
    """
    def _reply(outlook, namespace, entry_id, body, reply_all, account, confirm_send):
        item = _get_item(namespace, entry_id, account)
        if err := _check_item_class(item, OL_CLASS_MAIL, "mail item"):
            return err
        subject = item.Subject
        reply_item = item.ReplyAll() if reply_all else item.Reply()
        reply_item.Body = body + "\n\n" + reply_item.Body
        if DRAFT_MODE and not confirm_send:
            reply_item.Save()
            return (
                f"Reply draft saved for '{subject}' (reply_all={reply_all}, "
                f"entry_id={reply_item.EntryID})"
            )
        reply_item.Send()
        return f"Reply sent to '{subject}' (reply_all={reply_all})"

    try:
        return await bridge.call(_reply, entry_id, body, reply_all, account, confirm_send)
    except Exception as e:
        return f"Error replying to email: {format_com_error(e)}"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_send_gating.py -k reply_email -v`

Expected: `3 passed`.

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/server.py tests/test_send_gating.py
git commit -m "Add OUTLOOK_MCP_DRAFT_MODE gating to reply_email (threaded draft)"
```

---

### Task 7: `create_meeting` gating (unsent meeting request)

**Files:**
- Modify: `src/outlook_desktop_mcp/server.py:978-1057`
- Test: `tests/test_send_gating.py`

- [ ] **Step 1: Write the failing tests**

Add to `tests/test_send_gating.py`:

```python
# ---- create_meeting ----

def test_create_meeting_sends_invites_when_draft_mode_off(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", False)
    fake_outlook, _ = fake_com
    asyncio.run(server.create_meeting(
        subject="Sync", start="2026-08-05 09:00", end="2026-08-05 09:30",
        required_attendees="bob@example.com",
    ))
    appt = fake_outlook.CreateItem.return_value
    appt.Send.assert_called_once()
    appt.Save.assert_not_called()


def test_create_meeting_saves_unsent_when_draft_mode_on(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    result = asyncio.run(server.create_meeting(
        subject="Sync", start="2026-08-05 09:00", end="2026-08-05 09:30",
        required_attendees="bob@example.com",
    ))
    appt = fake_outlook.CreateItem.return_value
    appt.Recipients.ResolveAll.assert_called_once()
    appt.Save.assert_called_once()
    appt.Send.assert_not_called()
    assert "NOT sent" in result


def test_create_meeting_sends_when_draft_mode_on_and_confirm_send_true(fake_com, monkeypatch):
    monkeypatch.setattr(server, "DRAFT_MODE", True)
    fake_outlook, _ = fake_com
    asyncio.run(server.create_meeting(
        subject="Sync", start="2026-08-05 09:00", end="2026-08-05 09:30",
        required_attendees="bob@example.com", confirm_send=True,
    ))
    appt = fake_outlook.CreateItem.return_value
    appt.Send.assert_called_once()
    appt.Save.assert_not_called()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `uv run --no-sync pytest tests/test_send_gating.py -k create_meeting -v`

Expected: `test_create_meeting_sends_invites_when_draft_mode_off` `PASS`es by
coincidence (today's code always sends). The other two `FAIL`: `appt.Save`
is never called and `confirm_send` is not a valid keyword argument yet.

- [ ] **Step 3: Write minimal implementation**

Replace the entire `create_meeting` tool (currently lines 978-1057) with:

```python
@mcp.tool()
async def create_meeting(
    subject: str,
    start: str,
    end: str,
    required_attendees: str,
    location: str = "",
    body: str = "",
    optional_attendees: str = "",
    account: str = "",
    confirm_send: bool = False,
) -> str:
    """Create a meeting and send invitations to attendees.

    Creates a calendar meeting and sends meeting requests to all specified
    attendees. The meeting will appear on the organizer's calendar and
    attendees will receive an invitation they can accept, decline, or
    tentatively accept.

    When the server is running in draft-safe mode (OUTLOOK_MCP_DRAFT_MODE=1),
    this saves the meeting to the organizer's calendar WITHOUT sending
    invitations - it sits there unsent until opened in Outlook and sent
    manually. Pass confirm_send=true ONLY if the user has explicitly asked,
    in this turn, to send invitations immediately rather than draft the
    meeting. Ignored (has no effect) when the server is not running in
    draft-safe mode.

    Args:
        subject: The meeting title.
        start: Start time in ISO 8601 format (e.g. "2026-02-25 14:00").
        end: End time in ISO 8601 format (e.g. "2026-02-25 15:00").
        required_attendees: Required attendee email addresses, separated by
            semicolons. Example: "alice@example.com; bob@example.com"
        location: Optional. Meeting location (e.g. "Teams", "Room 301").
        body: Optional. Meeting description or agenda.
        optional_attendees: Optional. Optional attendee emails, separated
            by semicolons.
        account: Optional. Account display name (or substring) to send from.
            Default: primary account. Use list_accounts to see available accounts.
        confirm_send: Optional. Only meaningful in draft-safe mode - set true
            to actually send invitations instead of drafting. Default false.

    Returns:
        Confirmation that the meeting was created and invitations sent (or
        saved unsent in draft-safe mode).
    """
    def _create(outlook, namespace, subject, start, end, required_attendees,
                location, body, optional_attendees, account, confirm_send):
        appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
        # Set sending account
        if account:
            store = _require_store(namespace, account)
            for acc in outlook.Session.Accounts:
                if acc.DeliveryStore.StoreID == store.StoreID:
                    appt._oleobj_.Invoke(*(SEND_USING_ACCOUNT_DISPID, 0, 8, 0, acc))
                    break
        appt.Subject = subject
        appt.Start = start
        appt.End = end
        appt.MeetingStatus = OL_MEETING
        if location:
            appt.Location = location
        if body:
            appt.Body = body

        for addr in required_attendees.split(";"):
            addr = addr.strip()
            if addr:
                recip = appt.Recipients.Add(addr)
                recip.Type = OL_REQUIRED

        if optional_attendees:
            for addr in optional_attendees.split(";"):
                addr = addr.strip()
                if addr:
                    recip = appt.Recipients.Add(addr)
                    recip.Type = OL_OPTIONAL

        appt.Recipients.ResolveAll()
        if DRAFT_MODE and not confirm_send:
            appt.Save()
            return (
                f"Meeting '{subject}' saved to your calendar but invitations "
                f"were NOT sent (entry_id={appt.EntryID}). Open it in Outlook "
                f"and click Send to invite {required_attendees}."
            )
        appt.Send()
        return (
            f"Meeting '{subject}' created and invitations sent to "
            f"{required_attendees}"
        )

    try:
        return await bridge.call(
            _create, subject, start, end, required_attendees, location, body,
            optional_attendees, account, confirm_send,
        )
    except Exception as e:
        return f"Error creating meeting: {format_com_error(e)}"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `uv run --no-sync pytest tests/test_send_gating.py -v`

Expected: all tests in the file pass (9 total across Tasks 5-7).

- [ ] **Step 5: Commit**

```bash
git add src/outlook_desktop_mcp/server.py tests/test_send_gating.py
git commit -m "Add OUTLOOK_MCP_DRAFT_MODE gating to create_meeting"
```

---

### Task 8: Document `OUTLOOK_MCP_DRAFT_MODE` in README

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Add a new section**

In `README.md`, insert a new section directly after "## Available Tools by
Platform" and before "## Architecture Details" (i.e. after the "Total: 29
tools..." line):

```markdown
## Draft-Safe Mode

By default `send_email`, `reply_email`, and `create_meeting` transmit
immediately, same as always. Set `OUTLOOK_MCP_DRAFT_MODE=1` in the server's
environment to flip that default: the three tools create a reviewable
Outlook artifact instead (a Drafts item for email, an unsent meeting on the
calendar) and never transmit unless the caller explicitly passes
`confirm_send=true`.

| Tool | Draft-safe mode behavior |
|---|---|
| `send_email` | Saves to Drafts instead of sending |
| `reply_email` | Saves a threaded reply draft (via Outlook's own `Reply()`/`ReplyAll()`) instead of sending |
| `create_meeting` | Saves the meeting to your calendar without sending invitations - open it in Outlook and click Send when ready |

`respond_to_meeting` and cancelling an organized meeting via `delete_event`
are unaffected - both still act immediately regardless of this setting
(Outlook has no equivalent "draft" artifact for either).

Enable it when registering the server, e.g.:

```bash
claude mcp add outlook-desktop --env OUTLOOK_MCP_DRAFT_MODE=1 -- outlook-desktop-mcp
```
```

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "Document OUTLOOK_MCP_DRAFT_MODE in README"
```

---

### Task 9: Full regression pass + live smoke check

**Files:** none (verification only)

- [ ] **Step 1: Run the full automated test suite**

Run: `uv run --no-sync pytest tests/test_constants.py tests/test_com_bridge.py tests/test_formatting.py tests/test_send_gating.py -q`

(Do NOT run a bare `pytest -q` - it collects `calendar_com_test.py`,
`extras_com_test.py`, `phase1_com_test.py`, which are standalone manual
scripts requiring live Outlook, not pytest tests. They error under plain
collection regardless of this change - pre-existing, unrelated.)

Expected: all tests pass (29 pre-existing + new tests from Tasks 1-7), zero
failures, zero warnings.

- [ ] **Step 2: Live smoke check against real Outlook (safe operations only)**

With Outlook Desktop (Classic) running, from the worktree root:

```bash
uv run --no-sync python -c "
import asyncio
import os
os.environ['OUTLOOK_MCP_DRAFT_MODE'] = '1'
from outlook_desktop_mcp.com_bridge import OutlookBridge
from outlook_desktop_mcp import server
server.bridge = OutlookBridge()
server.DRAFT_MODE = True
server.bridge.start()
result = asyncio.run(server.send_email(
    to='devin-smoke-test@example.invalid',
    subject='draft-mode smoke test - safe to delete',
    body='If you see this in Drafts (not Sent), draft-safe mode works.',
))
print(result)
server.bridge.stop()
"
```

Expected output: `Draft saved: 'draft-mode smoke test - safe to delete' to
devin-smoke-test@example.invalid (entry_id=...)`.

Then manually check Outlook: the item must be in **Drafts**, not Sent Items.
Delete the test draft afterward.

**Do not** run this with `confirm_send=True` or with `OUTLOOK_MCP_DRAFT_MODE`
unset against a real recipient address - that would actually send mail.

- [ ] **Step 3: Update `docs/superpowers/STATUS.md`**

Add a new "Resume Point" entry reflecting this work is complete, following
the existing format in that file (last worked date, branch, tests passing
count, what's next).

- [ ] **Step 4: Final commit**

```bash
git add docs/superpowers/STATUS.md
git commit -m "Update STATUS.md: draft-mode and reconnect work complete"
```

---

## Out of Scope (see spec for rationale)

- macOS (`server_mac.py`) draft-mode parity
- `respond_to_meeting`, `delete_event` cancel-path gating
- Attachments on drafted emails/replies
- Any change to `docs/superpowers/plans/2026-05-22-outlook-cli.md` (Phase 2)
