@echo off
rem Double-click to build last week's report. Result: output folder.
chcp 65001 > nul
set PYTHONUTF8=1
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" run_report.py %*
) else (
    python run_report.py %*
)
if not defined NO_PAUSE pause
