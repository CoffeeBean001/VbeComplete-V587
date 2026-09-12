@echo off
rem VbeComplete launcher (silent background, no console window)
where py >nul 2>nul
if %errorlevel%==0 (set PYW=pyw -3) else (set PYW=pythonw)
echo Installing dependencies (first run only)...
%PYW% -m pip install -q -r "%~dp0requirements.txt"
if not %errorlevel%==0 (
  echo pip install failed. Please run "python main.py" in a console to see errors.
  pause
  exit /b 1
)
echo Starting VbeComplete (runs in background)...
start "" %PYW% "%~dp0main.py"
echo Done. You can close this window. The tool stays running.
timeout /t 2 >nul
