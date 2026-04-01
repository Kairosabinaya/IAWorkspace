"""Formatting queue with a multiprocessing-based worker.

Each file is formatted in a **separate OS process** so that heavy openpyxl
operations (load_workbook, cell iteration, wb.save) never block the GUI
process's GIL.  Progress flows back through a ``multiprocessing.Queue``.
"""

import multiprocessing
import queue as _queue  # stdlib thread-safe queue — only for Empty exception
import threading
import traceback
from collections import deque
from dataclasses import dataclass
from typing import Optional

from app.modules.excel_formatter.models.file_config import FileConfig


# ======================================================================
# Child-process entry point (runs with its OWN Python interpreter + GIL)
# ======================================================================

def _run_in_process(
    config: FileConfig,
    output_folder: str,
    progress_q: multiprocessing.Queue,
):
    """Called inside a spawned child process.  Completely independent GIL."""
    # Import here so the child process loads openpyxl in its own memory space.
    from app.modules.excel_formatter.engine.processor import process_file

    def progress_cb(file_name, pct, text):
        try:
            progress_q.put_nowait(("progress", file_name, pct, text))
        except Exception:
            pass  # queue full / broken — non-fatal

    try:
        success = process_file(config, output_folder, progress_cb)
        progress_q.put((
            "done", success,
            config.status, config.error_message, config.progress,
        ))
    except Exception as exc:
        traceback.print_exc()
        progress_q.put(("done", False, "Error", str(exc), 0.0))


# ======================================================================
# Job data
# ======================================================================

@dataclass
class FormatJob:
    """Represents a single formatting job in the queue."""

    job_id: str  # file_path as unique ID
    config: FileConfig
    output_folder: str
    status: str = "queued"  # queued | processing | done | error | cancelled
    progress: float = 0.0
    status_text: str = "Waiting"
    error_message: str = ""


# ======================================================================
# Queue manager
# ======================================================================

class FormattingQueue:
    """FIFO queue.  A coordinator thread picks jobs and spawns child
    processes; progress is relayed to the GUI via ``progress_state`` dict."""

    def __init__(self):
        self._queue: deque[FormatJob] = deque()
        self._current_job: Optional[FormatJob] = None
        self._completed: list[FormatJob] = []
        self._lock = threading.Lock()
        self._event = threading.Event()
        self._shutdown = False

        # GUI polling reads this dict (same pattern as before)
        self.progress_state: dict[str, tuple[float, str]] = {}

        self._worker = threading.Thread(target=self._worker_loop, daemon=True)
        self._worker.start()

    # ------------------------------------------------------------------
    # Public API (called from GUI / main thread)
    # ------------------------------------------------------------------

    def enqueue(self, config: FileConfig, output_folder: str) -> FormatJob:
        job = FormatJob(job_id=config.file_path, config=config,
                        output_folder=output_folder)
        with self._lock:
            self._queue.append(job)
        config.status = "Queued"
        self._event.set()
        return job

    def cancel(self, job_id: str) -> bool:
        with self._lock:
            for job in self._queue:
                if job.job_id == job_id:
                    self._queue.remove(job)
                    job.status = "cancelled"
                    job.config.status = "Ready"
                    return True
            return False

    def is_job_active(self, job_id: str) -> bool:
        with self._lock:
            if self._current_job and self._current_job.job_id == job_id:
                return True
            return any(j.job_id == job_id for j in self._queue)

    def is_job_processing(self, job_id: str) -> bool:
        with self._lock:
            return (self._current_job is not None
                    and self._current_job.job_id == job_id)

    def get_all_jobs(self) -> list[FormatJob]:
        with self._lock:
            result = list(self._completed)
            if self._current_job:
                result.append(self._current_job)
            result.extend(self._queue)
            return result

    def get_queue_position(self, job_id: str) -> int:
        with self._lock:
            if self._current_job and self._current_job.job_id == job_id:
                return 0
            for i, job in enumerate(self._queue):
                if job.job_id == job_id:
                    return i + 1
            return -1

    def is_idle(self) -> bool:
        with self._lock:
            return self._current_job is None and len(self._queue) == 0

    def has_completed(self) -> bool:
        with self._lock:
            return len(self._completed) > 0

    def pop_completed(self) -> list[FormatJob]:
        with self._lock:
            jobs = list(self._completed)
            self._completed.clear()
            return jobs

    def clear_completed(self):
        with self._lock:
            self._completed.clear()

    def cancel_all_queued(self):
        with self._lock:
            while self._queue:
                job = self._queue.popleft()
                job.status = "cancelled"
                job.config.status = "Ready"

    def shutdown(self):
        self._shutdown = True
        self._event.set()
        self._worker.join(timeout=5)

    # ------------------------------------------------------------------
    # Coordinator thread  (lightweight — just spawns processes & relays)
    # ------------------------------------------------------------------

    def _worker_loop(self):
        while not self._shutdown:
            self._event.wait()
            if self._shutdown:
                break
            self._event.clear()

            while not self._shutdown:
                with self._lock:
                    if not self._queue:
                        self._current_job = None
                        break
                    job = self._queue.popleft()
                    self._current_job = job

                job.status = "processing"
                job.config.status = "Processing..."
                self.progress_state[job.config.file_name] = (0.0, "Loading workbook...")

                self._run_job_in_process(job)

                with self._lock:
                    self._current_job = None
                    self._completed.append(job)

    def _run_job_in_process(self, job: FormatJob):
        """Spawn a child process for one job, relay progress until done."""
        progress_q: multiprocessing.Queue = multiprocessing.Queue()

        proc = multiprocessing.Process(
            target=_run_in_process,
            args=(job.config, job.output_folder, progress_q),
            daemon=True,
        )
        proc.start()

        # Relay progress messages until child signals "done"
        finished = False
        while not finished:
            # If child died unexpectedly, handle it
            if not proc.is_alive() and progress_q.empty():
                if not finished:
                    job.status = "error"
                    job.error_message = "Formatting process terminated unexpectedly"
                    job.config.status = "Error"
                    job.config.error_message = job.error_message
                break

            try:
                msg = progress_q.get(timeout=0.15)
            except _queue.Empty:
                continue

            kind = msg[0]
            if kind == "progress":
                _, file_name, pct, text = msg
                job.progress = pct
                job.status_text = text
                self.progress_state[file_name] = (pct, text)
            elif kind == "done":
                _, success, status, error_msg, progress = msg
                if success and status != "Error":
                    job.status = "done"
                    job.progress = 1.0
                    job.status_text = "Done"
                    job.config.status = "Done"
                    job.config.progress = 1.0
                else:
                    job.status = "error"
                    job.error_message = error_msg
                    job.config.status = "Error"
                    job.config.error_message = error_msg
                finished = True

        proc.join(timeout=10)
        if proc.is_alive():
            proc.terminate()
        progress_q.close()
