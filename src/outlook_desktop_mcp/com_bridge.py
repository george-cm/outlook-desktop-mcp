"""
COM Threading Bridge
====================
Runs all Outlook COM calls on a dedicated STA (Single-Threaded Apartment)
thread so the async MCP event loop never touches COM objects directly.

Every COM function passed to bridge.call() receives (outlook, namespace, ...)
as its first two arguments - the live COM objects that only exist on the
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
        # Snapshot the thread before enqueuing. A concurrent start() could
        # swap self._thread after queue.put(); checking owning_thread ensures
        # we detect death of the thread that actually owns this request.
        owning_thread = self._thread
        self._request_queue.put((func, args, kwargs, result_event, result_holder))

        # Poll in short intervals so we detect COM thread death promptly
        # rather than waiting the full 60s for a misleading TimeoutError.
        loop = asyncio.get_running_loop()
        deadline = 60.0
        interval = 0.5
        elapsed = 0.0
        while elapsed < deadline:
            signaled = await loop.run_in_executor(
                None, lambda: result_event.wait(timeout=interval)
            )
            if signaled:
                break
            # Re-check bridge health after each poll interval.
            if not self._alive.is_set() or not owning_thread.is_alive():
                raise RuntimeError(
                    "COM bridge died while processing request. "
                    "Restart the bridge to continue."
                )
            elapsed += interval
        else:
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
