import asyncio
import os
import threading
import queue
from collections import defaultdict
from hashlib import md5
from typing import Any, Awaitable, Callable, Dict, Optional
import os.path

# Global serialization for Office COM actions (safest default)
GLOBAL_OFFICE_LOCK = asyncio.Lock()

# Per-document serialization when a stable key (e.g., absolute path) is known
_doc_locks: Dict[str, asyncio.Lock] = defaultdict(asyncio.Lock)


def get_doc_lock(key: str) -> asyncio.Lock:
    """Return (and create if needed) an asyncio.Lock for a document key.

    Use a stable, unique key such as an absolute file path. If you don't know the
    path (e.g., operating on the current active, unsaved document), prefer the
    GLOBAL_OFFICE_LOCK instead.
    """
    return _doc_locks[key]


async def with_global_lock(coro: Callable[[], Awaitable]):
    """Run the given coroutine under the global Office lock."""
    async with GLOBAL_OFFICE_LOCK:
        return await coro()


async def with_doc_lock(key: str, coro: Callable[[], Awaitable]):
    """Run the given coroutine under a per-document lock.

    NOTE: This does not guarantee parallelism across documents unless the
    underlying automation layer uses separate Word instances. It does ensure
    requests targeting the same document are serialized.
    """
    async with get_doc_lock(key):
        return await coro()


# ------------------------
# Word worker pool (Windows only)
# ------------------------

class _WordWorker(threading.Thread):
    def __init__(self) -> None:
        super().__init__(daemon=True)
        self._tasks: "queue.Queue[tuple[Callable[[Any], Any], asyncio.Future]]" = queue.Queue()
        self._stopped = threading.Event()
        self._word: Optional[Any] = None

    def run(self) -> None:
        try:
            import pythoncom  # type: ignore
            pythoncom.CoInitialize()
            try:
                from win32com.client import DispatchEx  # type: ignore
            except Exception as e:
                # Surface error to any waiting tasks
                while not self._tasks.empty():
                    fn, fut = self._tasks.get_nowait()
                    if not fut.done():
                        fut.set_exception(e)
                return
            # Each worker creates its own isolated Word instance
            self._word = DispatchEx("Word.Application")
            try:
                if self._word is not None:
                    # Minimize UI/modal interruptions (can be overridden via env)
                    try:
                        visible_env = os.getenv("MCP_WORD_VISIBLE", "0").strip().lower()
                        self._word.Visible = visible_env in ("1", "true", "yes")
                    except Exception:
                        self._word.Visible = False
                    self._word.DisplayAlerts = 0  # wdAlertsNone
                    # Disable macro prompts and other security popups during automation
                    # msoAutomationSecurityForceDisable = 3
                    try:
                        self._word.AutomationSecurity = 3
                    except Exception:
                        pass
                    try:
                        opts = getattr(self._word, "Options", None)
                        if opts is not None:
                            # Avoid conversion prompts and link updates on open
                            try:
                                opts.ConfirmConversions = False
                            except Exception:
                                pass
                            try:
                                opts.UpdateLinksAtOpen = False
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass

            while not self._stopped.is_set():
                try:
                    fn, fut = self._tasks.get(timeout=0.05)
                except queue.Empty:
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:
                        pass
                    continue
                try:
                    result = fn(self._word)
                    if not fut.done():
                        fut.set_result(result)
                except Exception as e:
                    if not fut.done():
                        fut.set_exception(e)
                finally:
                    try:
                        pythoncom.PumpWaitingMessages()
                    except Exception:
                        pass
        finally:
            try:
                if self._word is not None:
                    try:
                        self._word.Quit()
                    except Exception:
                        pass
            finally:
                try:
                    import pythoncom  # type: ignore
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def submit(self, fn: Callable[[Any], Any], loop: asyncio.AbstractEventLoop) -> asyncio.Future:
        fut: asyncio.Future = loop.create_future()
        self._tasks.put((fn, fut))
        return fut

    def stop(self) -> None:
        self._stopped.set()


_POOL_SIZE = int(os.getenv("MCP_WORD_POOL_SIZE", "2"))
_WORD_WORKERS: list[_WordWorker] = []


def _ensure_pool_started() -> None:
    global _WORD_WORKERS
    if _WORD_WORKERS:
        return
    if os.name != "nt":
        # Non-Windows: pool not applicable
        _WORD_WORKERS = []
        return
    size = max(1, _POOL_SIZE)
    _WORD_WORKERS = [_WordWorker() for _ in range(size)]
    for w in _WORD_WORKERS:
        w.start()


def _pick_worker(key: str) -> Optional[_WordWorker]:
    _ensure_pool_started()
    if not _WORD_WORKERS:
        return None
    i = int(md5(key.encode("utf-8")).hexdigest(), 16) % len(_WORD_WORKERS)
    return _WORD_WORKERS[i]


async def run_in_word_worker(key: str, fn: Callable[[Any], Any]):
    """Run a synchronous COM function against a dedicated Word instance.

    The callable receives a Word Application object bound to that worker's STA thread.
    All COM operations must be performed synchronously inside the callable.
    """
    worker = _pick_worker(key)
    if worker is None:
        # Fallback: run on a generic background thread (will bind to default Word instance)
        return await asyncio.to_thread(lambda: fn(None))
    loop = asyncio.get_running_loop()
    fut = worker.submit(fn, loop)
    return await fut


# ------------------------
# Path canonicalization helpers
# ------------------------

def canonical_doc_key(path_str: str) -> str:
    """Return a stable, canonical key for a file path used to route work to a worker/lock.

    - Absolute path
    - Normalized separators
    - Case-normalized on Windows (so C:\\ and c:\\ map to the same key)
    """
    try:
        # Accept file:// URIs and plain paths
        if path_str.lower().startswith("file:"):
            from .path_utils import _from_file_uri
            fs_path = str(_from_file_uri(path_str))
        else:
            fs_path = path_str
        abs_path = os.path.abspath(fs_path)
        norm = os.path.normpath(abs_path)
        return os.path.normcase(norm)
    except Exception:
        # Best-effort fallback
        return path_str

