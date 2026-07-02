"""Two logs, both written to the log folder.

  refresh-events.csv  -- one row per event (run start/end, and each workbook's
                         result). Tabular and append-only, so you can build reports
                         on it over time in Power BI, Excel, or SQL. Every row from
                         one run shares a run_id.
  excel_refresh.log    -- a rotating, human-readable log for debugging a single run.
"""

from __future__ import annotations

import csv
import logging
import uuid
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

EVENT_COLUMNS = [
    "run_id", "timestamp", "scope", "target", "event", "status",
    "duration_seconds", "error_type", "error_message", "detail",
]


def _now_iso() -> str:
    # Local time, millisecond precision, with a UTC offset -> unambiguous and sortable.
    return datetime.now().astimezone().isoformat(timespec="milliseconds")


def setup_text_logger(log_dir: Path) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("excel_refresh")
    logger.setLevel(logging.INFO)
    if not logger.handlers:                          # avoid duplicate handlers
        handler = RotatingFileHandler(
            log_dir / "excel_refresh.log", maxBytes=2_000_000,
            backupCount=5, encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s"))
        logger.addHandler(handler)
    return logger


class EventLog:
    """Append-only writer for refresh-events.csv."""

    def __init__(self, log_dir: Path) -> None:
        log_dir.mkdir(parents=True, exist_ok=True)
        self.run_id = uuid.uuid4().hex[:12]
        self._path = log_dir / "refresh-events.csv"
        is_new = (not self._path.exists()) or self._path.stat().st_size == 0
        self._fh = self._path.open("a", newline="", encoding="utf-8")
        self._writer = csv.writer(self._fh)
        if is_new:
            self._writer.writerow(EVENT_COLUMNS)
            self._fh.flush()

    def emit(
        self, *, scope: str, target: str, event: str, status: str,
        duration_seconds: float | None = None, error_type: str = "",
        error_message: str = "", detail: str = "",
    ) -> None:
        self._writer.writerow([
            self.run_id, _now_iso(), scope, target, event, status,
            "" if duration_seconds is None else f"{duration_seconds:.3f}",
            error_type, error_message, detail,
        ])
        self._fh.flush()                             # durable across a later crash

    def close(self) -> None:
        self._fh.close()
