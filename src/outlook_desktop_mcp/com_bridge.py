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
        """Start the COM thread. Idempotent - safe to call multiple times."""
        if self._thread is not None and self._thread.is_alive():
            return
        # Reset state so restart after stop() works correctly.
        self._shutdown.clear()
        self._ready.clear()
        self._alive.clear()
        self._init_error = None
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
