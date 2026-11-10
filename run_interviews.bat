@echo off
rem Interview radar only (radar.py interviews): forums -> Groq -> Obsidian notes. Log in data\radar.log.
rem If Groq fails the notes are still regenerated: the queue waits, and the tracker section stays current.
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
call "%~dp0find_python.bat" || exit /b 1
%PY% radar.py interviews --log
exit /b %errorlevel%
