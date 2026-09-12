@echo off
rem VbeComplete launcher with diagnostic logging (troubleshooting only).
rem Default is run.bat (no log file). Use this one only when you need logs.
rem Log file: VbeComplete.log in the same directory as this .bat.

set VBECOMPLETE_LOG=1

where py >nul 2>nul
if %errorlevel%==0 (set PYW=pyw -3) else (set PYW=pythonw)

echo Installing dependencies (first run only)...
%PYW% -m pip install -q -r "%~dp0requirements.txt"
if not %errorlevel%==0 goto :pip_err

if exist "%~dp0VbeComplete.log" del "%~dp0VbeComplete.log"

echo Starting VbeComplete with diagnostic log enabled...
echo Log file: %~dp0VbeComplete.log
echo (run show_log.bat to open it; run.bat for normal use)

start "" %PYW% "%~dp0main.py"

goto :eof

:pip_err
echo Failed to install dependencies. Run install_autostart.bat first or run manually.
pause
exit /b 1
