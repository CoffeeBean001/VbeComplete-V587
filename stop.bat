@echo off
rem Stop running VbeComplete instances (ASCII-only file on purpose)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0stop.ps1"
pause
