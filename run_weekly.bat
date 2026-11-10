@echo off
rem Weekly Telegram digest of everything registered in the last 7 days, then the
rem interview radar (forums -> Groq -> Radar notes, which also follow the tracker).
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
call "%~dp0find_python.bat" || exit /b 1
if not exist data mkdir data
echo ==== %date% %time% weekly ==== >> data\radar.log
%PY% scripts\notify.py --weekly >> data\radar.log 2>&1
call run_interviews.bat
