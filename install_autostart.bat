@echo off
rem Install VbeComplete to run at Windows startup (silent background).
where py >nul 2>nul
if %errorlevel%==0 (set PYW=py -3) else (set PYW=python)
echo Installing dependencies (first run only)...
%PYW% -m pip install -q -r "%~dp0requirements.txt"
if not %errorlevel%==0 (
  echo pip install failed. Please run "python main.py" in a console to see errors.
  pause
  exit /b 1
)
echo Creating startup shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_autostart.ps1"
pause
