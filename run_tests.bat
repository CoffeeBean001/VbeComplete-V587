@echo off
rem Run regression tests (uses real project modules in tests\testdata,
rem does not require Excel).

where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)

%PY% "%~dp0tests\test_real_project.py"
pause
