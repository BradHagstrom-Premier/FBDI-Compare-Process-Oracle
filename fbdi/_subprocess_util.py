"""Shared subprocess runner for per-file isolation.

Both catalog.py and compare.py run per-file work in a fresh subprocess so that
openpyxl resource leaks don't accumulate across sequential loads. This module
centralises the result-handling pattern.

Drain-before-join is required. multiprocessing.Queue.put() hands the payload
to a background feeder thread which writes to an OS pipe (~64 KB buffer on
Windows). If the payload exceeds the pipe buffer and the parent joins first,
the feeder blocks, the child cannot exit, and join() times out. Reading from
the queue before joining unblocks the feeder.
"""

import multiprocessing
import queue as queue_module
import time
from dataclasses import dataclass
from typing import Any, Callable


@dataclass
class WorkerOutcome:
    """Outcome of one subprocess run.

    status:
      "completed" — worker put a message; payload holds it.
      "timeout"   — deadline reached while worker still alive.
      "crashed"   — worker exited without putting anything on the queue.
    """
    status: str
    payload: Any
    exitcode: int | None


def run_worker(
    target: Callable,
    args: tuple,
    timeout: int,
    poll_interval: float = 0.5,
) -> WorkerOutcome:
    """Run target(*args, queue) in a subprocess; drain queue before joining.

    The target must accept `queue` as its final positional argument and put
    exactly one message on it then return. The helper polls the queue and
    process liveness together so that a crash before put() is reported as
    "crashed" rather than "timeout", and a large payload cannot deadlock
    the feeder thread behind a join().
    """
    mp_queue: multiprocessing.Queue = multiprocessing.Queue()
    proc = multiprocessing.Process(target=target, args=args + (mp_queue,))
    proc.start()

    deadline = time.monotonic() + timeout
    payload: Any = None
    status: str | None = None

    while True:
        try:
            payload = mp_queue.get(timeout=poll_interval)
            status = "completed"
            break
        except queue_module.Empty:
            pass

        if not proc.is_alive():
            try:
                payload = mp_queue.get_nowait()
                status = "completed"
            except queue_module.Empty:
                status = "crashed"
            break

        if time.monotonic() >= deadline:
            proc.terminate()
            proc.join(5)
            return WorkerOutcome("timeout", None, proc.exitcode)

    proc.join(5)
    if proc.is_alive():
        proc.terminate()
        proc.join(5)

    return WorkerOutcome(status, payload, proc.exitcode)
