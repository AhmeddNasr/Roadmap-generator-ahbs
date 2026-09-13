@echo off
echo Running ADO Roadmap Sync...
cd /d "%~dp0"
python ado_roadmap_sync.py
echo.
pause
