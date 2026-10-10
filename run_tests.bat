@echo off
rem Run regression tests.
rem   Sample project : tests\testdata\ (14 modules exported from a real .xlsm)
rem   Needs Excel COM only for 3 cases (host type-library enums); the other 871
rem   are pure parsing-layer checks and run anywhere.
rem   Keep this file in the repo — delete the tests and every later change has
rem   to be verified on the live VBE alone.

where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)

%PY% "%~dp0tests\test_real_project.py"
echo.
pause
