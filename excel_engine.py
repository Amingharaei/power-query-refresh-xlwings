"""The Excel side of the job: find the workbooks, and refresh one workbook in an
isolated Excel instance guarded by a watchdog.

Refresh strategy -- each connection (query) is refreshed INDIVIDUALLY, inside
try/except, so a failure is caught and attributed to the specific query. This is the
only reliable way to detect failures: Excel's "Refresh All", when run from code, does
NOT raise or report an error when a query fails, so a broken query would pass
silently. See the README for the trade-off this implies for queries that depend on
each other.

If any query in a workbook fails, the workbook is NOT saved -- it keeps its last good
version rather than being left half-updated -- and the failure is logged and emailed.

A fresh Excel instance is used per workbook (so one bad file can't sink the batch),
and a watchdog force-closes Excel if a refresh runs past the timeout.
"""

from __future__ import annotations

import os
import signal
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from time import perf_counter

import xlwings as xw

_XL_OLEDB = 1        # Power Query connections are OLEDB.
_XL_ODBC = 2


@dataclass
class ConnectionResult:
    name: str
    status: str                     # SUCCESS | FAILED | SKIPPED
    duration: float = 0.0
    error_type: str = ""
    error_message: str = ""


@dataclass
class WorkbookResult:
    workbook: str
    status: str = "SUCCESS"         # SUCCESS | FAILED
    duration: float = 0.0
    connections: list[ConnectionResult] = field(default_factory=list)
    error_type: str = ""            # set only for a whole-workbook failure
    error_message: str = ""

    @property
    def failed(self) -> list[ConnectionResult]:
        return [c for c in self.connections if c.status == "FAILED"]


def find_workbooks(reports_dir: Path, include: tuple[str, ...],
                   exclude: tuple[str, ...]) -> list[Path]:
    """Report files in reports_dir, sorted, skipping Excel lock files."""
    if not reports_dir.is_dir():
        raise FileNotFoundError(f"Reports directory not found: {reports_dir}")
    matched: set[Path] = set()
    for pattern in include:
        matched.update(reports_dir.glob(pattern))
    result: list[Path] = []
    for path in sorted(matched):
        if path.name.startswith("~$"):             # Excel's open-file lock marker
            continue
        if any(path.match(pat) for pat in exclude):
            continue
        result.append(path)
    return result


@contextmanager
def _excel_app(timeout_seconds: int, logger):
    """A fresh, invisible Excel instance with a watchdog that kills it on timeout."""
    app = xw.App(visible=False, add_book=False)
    pid = app.pid
    done = threading.Event()

    def watchdog() -> None:
        if not done.wait(timeout_seconds):          # False => timed out
            logger.error("Timeout after %ss; killing Excel PID %s.", timeout_seconds, pid)
            try:
                os.kill(pid, signal.SIGTERM)        # TerminateProcess on Windows
            except OSError as exc:
                logger.error("Could not kill Excel PID %s: %s", pid, exc)

    watcher = threading.Thread(target=watchdog, daemon=True)
    watcher.start()
    try:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.api.AskToUpdateLinks = False
        except Exception:
            pass
        yield app
    finally:
        done.set()                                  # stop the watchdog
        try:
            app.quit()
        except Exception:                           # may already be killed
            pass
        watcher.join(timeout=5)


def _background_off(conn) -> bool:
    """Turn off background refresh so the refresh call blocks until done.
    Returns True if this is a refreshable data connection (OLEDB/ODBC)."""
    ctype = conn.Type
    if ctype == _XL_OLEDB:
        conn.OLEDBConnection.BackgroundQuery = False
        return True
    if ctype == _XL_ODBC:
        conn.ODBCConnection.BackgroundQuery = False
        return True
    return False


def _refresh_with_retry(conn, logger) -> None:
    """Refresh one connection, retrying once. Raises if it fails both times."""
    try:
        conn.Refresh()
    except Exception as first:
        # The first code-driven refresh right after opening a file sometimes throws a
        # spurious 'initialization of the data source failed'. A second attempt
        # usually succeeds; a genuinely broken query fails again and is caught below.
        logger.info("Retrying connection '%s' after: %s", conn.Name, first)
        conn.Refresh()


def refresh_workbook(path: Path, timeout_seconds: int, events, logger) -> WorkbookResult:
    """Open, refresh each query (capturing per-query outcome), save only if all
    succeeded, then close -- all in this workbook's own Excel instance."""
    result = WorkbookResult(workbook=path.name)
    start = perf_counter()
    try:
        with _excel_app(timeout_seconds, logger) as app:
            wb = app.books.open(str(path), update_links=False)
            try:
                for conn in wb.api.Connections:
                    name = conn.Name
                    try:
                        refreshable = _background_off(conn)
                    except Exception:
                        refreshable = False
                    if not refreshable:
                        result.connections.append(ConnectionResult(name, "SKIPPED"))
                        events.emit(scope="connection", target=name, event="REFRESH",
                                    status="SKIPPED", detail="not an OLEDB/ODBC connection")
                        continue

                    events.emit(scope="connection", target=name, event="REFRESH", status="STARTED")
                    c_start = perf_counter()
                    try:
                        _refresh_with_retry(conn, logger)
                    except Exception as exc:
                        dur = perf_counter() - c_start
                        cr = ConnectionResult(name, "FAILED", dur,
                                              type(exc).__name__, str(exc)[:500])
                        result.connections.append(cr)
                        events.emit(scope="connection", target=name, event="REFRESH",
                                    status="FAILED", duration_seconds=dur,
                                    error_type=cr.error_type, error_message=cr.error_message)
                        logger.error("Query '%s' failed: %s", name, exc)
                    else:
                        dur = perf_counter() - c_start
                        result.connections.append(ConnectionResult(name, "SUCCESS", dur))
                        events.emit(scope="connection", target=name, event="REFRESH",
                                    status="SUCCESS", duration_seconds=dur)

                # Save only if every query succeeded; otherwise keep the last good file.
                if result.failed:
                    result.status = "FAILED"
                    logger.error("%s: %d query(ies) failed; not saving.",
                                 path.name, len(result.failed))
                else:
                    app.calculate()
                    wb.save()
            finally:
                try:
                    wb.close()
                except Exception:
                    pass                            # don't mask the real error
    except Exception as exc:
        # Whole-workbook failure: open failed, watchdog kill, Excel crash, save error.
        result.status = "FAILED"
        result.error_type = type(exc).__name__
        result.error_message = str(exc)[:500]
        logger.error("Workbook '%s' failed: %s", path.name, exc)
    result.duration = perf_counter() - start
    return result
