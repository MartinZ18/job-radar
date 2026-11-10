@echo off
rem Interview radar: forums -> Groq extraction -> Obsidian notes. Log in data\radar.log.
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
call "%~dp0find_python.bat" || exit /b 1
if not exist data mkdir data
echo ==== %date% %time% interviews ==== >> data\radar.log
%PY% scripts\fetch_interviews.py >> data\radar.log 2>&1
rem If Groq fails the notes are still regenerated: the queue waits, and the tracker section stays current.
%PY% scripts\extract_interviews.py >> data\radar.log 2>&1 || echo extract FAILED, notes rebuilt from the items already saved >> data\radar.log
%PY% scripts\update_obsidian.py >> data\radar.log 2>&1 || echo radar notes FAILED >> data\radar.log
exit /b 0
