@echo off
rem 探测 VBE 自带提示列表的窗口（只读，不碰代码）
where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)
%PY% "%~dp0probe_vbe_list.py" %1
echo.
pause
