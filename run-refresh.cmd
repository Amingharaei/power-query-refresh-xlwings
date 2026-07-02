@echo off
REM ============================================================
REM  Runs the Excel refresh. Task Scheduler (or a double-click)
REM  points at THIS file. It self-locates via %~dp0, so the
REM  project folder can live anywhere.
REM ============================================================

REM %~dp0 = the folder this file lives in (the project folder).
cd /d "%~dp0"

REM Title + banner print instantly, before Python has even started up, so the
REM window never just sits there with a blinking cursor and nothing else.
title Excel Refresh Orchestrator
echo.
echo  Excel Refresh Orchestrator
echo  Starting up...
echo.

REM Run with the project's own Python environment (created by setup.ps1 / uv sync).
REM Calling the venv's python directly needs no PATH lookup, which is the most
REM reliable option under Task Scheduler.
".venv\Scripts\python.exe" refresh_reports.py --config "config.toml"

REM To hide the console window during scheduled runs, use pythonw instead:
REM ".venv\Scripts\pythonw.exe" refresh_reports.py --config "config.toml"
