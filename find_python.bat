@echo off
rem Sets PY to the Python launcher to use. Set PY yourself before calling the
rem run_*.bat files to force one (e.g. a virtualenv's python.exe).
if defined PY exit /b 0
where py >nul 2>&1 && (set "PY=py -3" & exit /b 0)
where python >nul 2>&1 && (set "PY=python" & exit /b 0)
echo Python not found: install it from python.org or set PY to python.exe 1>&2
if exist data echo Python not found >> data\radar.log
exit /b 1
