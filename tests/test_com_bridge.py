"""
Unit tests for OutlookBridge lifecycle.
No Outlook or COM installation required - _com_thread_main is patched.
"""
import asyncio
import queue as _queue
import threading
import pytest
import pythoncom
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
