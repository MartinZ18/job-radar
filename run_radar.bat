@echo off
rem Full chain (radar.py run): fetch -> excel -> obsidian notes -> telegram. Output goes to data\radar.log.
rem Scheduled by Task Scheduler (python radar.py setup --schedule); safe to run by hand too.
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
call "%~dp0find_python.bat" || exit /b 1
%PY% radar.py run --log
exit /b %errorlevel%
