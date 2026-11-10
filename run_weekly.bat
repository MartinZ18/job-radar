@echo off
rem Weekly Telegram digest of the last 7 days, then the interview radar (radar.py weekly).
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
call "%~dp0find_python.bat" || exit /b 1
%PY% radar.py weekly --log
exit /b %errorlevel%
