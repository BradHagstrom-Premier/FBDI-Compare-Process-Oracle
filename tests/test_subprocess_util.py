"""Tests for fbdi._subprocess_util.run_worker."""

import time

from fbdi._subprocess_util import run_worker


# Worker targets must be module-level so Windows spawn can pickle them.

def _target_small(queue):
    queue.put(("ok", 42))


def _target_large(queue):
    # ~200 KB serialized — comfortably above the Windows pipe buffer (~64 KB)
    # and above the macOS pipe buffer (~16-64 KB depending on kernel).
    queue.put(("ok", ["x" * 100] * 2000))


def _target_slow(queue):
    time.sleep(10)
    queue.put(("ok", "too late"))


def _target_crash(queue):
    raise RuntimeError("worker died before put")


def _target_error_sentinel(queue):
    queue.put("ERROR: synthetic")


class TestRunWorker:
    def test_small_payload_returns_completed(self):
        out = run_worker(_target_small, args=(), timeout=30)
        assert out.status == "completed"
        assert out.payload == ("ok", 42)
        assert out.exitcode == 0

    def test_large_payload_no_deadlock(self):
        # Regression: pre-fix this hung for the full timeout because the
        # feeder thread could not drain the pipe while parent was joining.
        t0 = time.perf_counter()
        out = run_worker(_target_large, args=(), timeout=30)
        elapsed = time.perf_counter() - t0
        assert out.status == "completed"
        assert out.payload[0] == "ok"
        assert len(out.payload[1]) == 2000
        assert elapsed < 15, f"returned in {elapsed:.1f}s (possible deadlock)"

    def test_timeout_when_worker_hangs(self):
        out = run_worker(_target_slow, args=(), timeout=2)
        assert out.status == "timeout"
        assert out.payload is None

    def test_crash_before_put_reported_as_crashed(self):
        out = run_worker(_target_crash, args=(), timeout=30)
        assert out.status == "crashed"
        assert out.payload is None
        assert out.exitcode not in (0, None)

    def test_error_sentinel_is_passed_through(self):
        # Helper does not interpret "ERROR:" — caller owns that convention.
        out = run_worker(_target_error_sentinel, args=(), timeout=30)
        assert out.status == "completed"
        assert out.payload == "ERROR: synthetic"
