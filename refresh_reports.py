"""Excel Refresh Orchestrator -- the program you run.

    uv run python refresh_reports.py
    (or, without uv:  .venv\\Scripts\\python refresh_reports.py)

It reads config.toml, refreshes every workbook in the reports folder (each in its own
Excel instance), writes the event log, optionally emails an Outlook summary, and exits
0 if all workbooks succeeded or 1 if any failed. Task Scheduler reads that exit code as
"Last Run Result".
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import replace
from pathlib import Path

from config import Config, ConfigError, load_config
from eventlog import EventLog, setup_text_logger
from excel_engine import WorkbookResult, find_workbooks, refresh_workbook


def _console(message: str = "") -> None:
    """Best-effort console feedback for whoever is watching the window.

    This is separate from the logger on purpose: the logger always writes to
    excel_refresh.log, but this window may have no console at all -- if
    run-refresh.cmd is ever switched to pythonw.exe (see the comment in that
    file), sys.stdout is None and a bare print() would crash the run. Catching
    that here means the console message can fail silently without ever
    affecting the refresh itself or the exit code.
    """
    try:
        print(message, flush=True)
    except Exception:
        pass


def run(config: Config) -> int:
    _console("Excel Refresh Orchestrator")
    _console("The reports are being refreshed one at a time. This can take a while")
    _console("depending on how many reports there are and how large they are.")
    _console("This window will close on its own once every report has been refreshed.")
    _console()

    logger = setup_text_logger(config.log_dir)
    events = EventLog(config.log_dir)
    logger.info("Run %s start. Reports folder: %s", events.run_id, config.reports_dir)
    events.emit(scope="run", target=str(config.reports_dir), event="RUN", status="STARTED")

    results: list[WorkbookResult] = []
    try:
        workbooks = find_workbooks(config.reports_dir, config.include, config.exclude)
        logger.info("Found %d workbook(s).", len(workbooks))
        _console(f"Found {len(workbooks)} report(s) in {config.reports_dir}")
        _console()

        for i, path in enumerate(workbooks, start=1):
            logger.info("Refreshing %s", path.name)
            _console(f"[{i}/{len(workbooks)}] Refreshing {path.name} ...")
            events.emit(scope="workbook", target=path.name, event="WORKBOOK", status="STARTED")
            result = refresh_workbook(path, config.timeout_seconds, events, logger)
            results.append(result)
            detail = f"{len(result.failed)} query failure(s)" if result.failed else ""
            events.emit(
                scope="workbook", target=result.workbook, event="WORKBOOK",
                status=result.status, duration_seconds=result.duration,
                error_type=result.error_type, error_message=result.error_message, detail=detail,
            )
            note = f": {detail}" if detail else ""
            _console(f"    {result.status:7} ({result.duration:.1f}s){note}")
    finally:
        any_failed = any(r.status == "FAILED" for r in results)
        events.emit(
            scope="run", target=str(config.reports_dir), event="RUN",
            status="FAILED" if any_failed else "SUCCESS",
            detail=f"{len(results)} workbook(s)",
        )
        events.close()
        _email_summary(config, results, events.run_id, logger)

    any_failed = any(r.status == "FAILED" for r in results)
    logger.info("Run done. %d workbook(s), failures=%s", len(results), any_failed)
    ok = len(results) - sum(1 for r in results if r.status == "FAILED")
    _console()
    _console(f"Done. {ok}/{len(results)} report(s) refreshed successfully.")
    if any_failed:
        _console("Check excel_refresh.log and refresh-events.csv in the log folder for details.")
    return 1 if any_failed else 0


# --- Outlook summary email ---------------------------------------------------

def _email_summary(config: Config, results: list[WorkbookResult], run_id: str, logger) -> None:
    if not config.email.enabled:
        return
    any_failed = any(r.status == "FAILED" for r in results)
    if config.email.send_on == "failure" and not any_failed:
        return
    try:
        subject, body = _build_summary(results, run_id)
        _send_outlook(config.email.recipients, subject, body)
        logger.info("Summary email sent to %d recipient(s).", len(config.email.recipients))
    except Exception as exc:
        # A mail failure must never change the run's exit code or crash the process.
        logger.error("Could not send summary email: %s", exc)


def _build_summary(results: list[WorkbookResult], run_id: str) -> tuple[str, str]:
    total = len(results)
    failed = [r for r in results if r.status == "FAILED"]
    ok = total - len(failed)
    subject = (
        f"[Excel refresh] {ok}/{total} ok"
        + (f", {len(failed)} FAILED" if failed else "")
        + f" (run {run_id})"
    )
    lines = [f"Run {run_id}", f"Workbooks: {total}   Succeeded: {ok}   Failed: {len(failed)}", ""]
    for r in results:
        note = "  (not saved)" if r.failed else ""
        lines.append(f"{r.status:8} {r.workbook}  ({r.duration:.1f}s){note}")
        for c in r.failed:
            lines.append(f"    - query '{c.name}': {c.error_type}: {c.error_message}")
        if r.status == "FAILED" and not r.failed and r.error_message:
            lines.append(f"    - {r.error_type}: {r.error_message}")
    return subject, "\n".join(lines)


def _send_outlook(recipients: tuple[str, ...], subject: str, body: str) -> None:
    import win32com.client            # provided by pywin32 (installed with xlwings)

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)      # 0 = olMailItem
    mail.To = "; ".join(recipients)
    mail.Subject = subject
    mail.Body = body
    mail.Send()


# --- entry point -------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Refresh every Excel report in a folder.")
    parser.add_argument("--config", type=Path, default=Path("config.toml"),
                        help="Path to the config file (default: ./config.toml).")
    parser.add_argument("--reports-dir", type=Path, default=None,
                        help="Override the reports folder for a one-off run.")
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
    except ConfigError as exc:
        try:
            print(f"Configuration error: {exc}", file=sys.stderr)
        except Exception:
            pass
        return 2

    if args.reports_dir is not None:
        config = replace(config, reports_dir=args.reports_dir)

    return run(config)


if __name__ == "__main__":
    sys.exit(main())
