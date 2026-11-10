@echo off
rem Full chain: fetch -> excel -> obsidian notes -> telegram. Output goes to data\radar.log.
rem The notes step runs before telegram (notify.py deletes data\last_added.json); if it fails the message still goes out.
rem Scheduled by Task Scheduler (see README); safe to run by hand too.
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
call "%~dp0find_python.bat" || exit /b 1
if not exist data mkdir data
echo ==== %date% %time% ==== >> data\radar.log
%PY% scripts\fetch_jobs.py >> data\radar.log 2>&1 && %PY% scripts\update_excel.py >> data\radar.log 2>&1 || (echo fetch/update FAILED >> data\radar.log & exit /b 1)
%PY% scripts\write_job_notes.py >> data\radar.log 2>&1 || echo job notes FAILED >> data\radar.log
%PY% scripts\notify.py >> data\radar.log 2>&1
exit /b 0
