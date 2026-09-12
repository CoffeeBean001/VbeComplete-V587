@echo off
if not exist "%~dp0VbeComplete.log" (
  echo No log file yet. Run run_debug.bat first, then trigger the completer once.
  pause
  exit /b 1
)
start "" notepad "%~dp0VbeComplete.log"
