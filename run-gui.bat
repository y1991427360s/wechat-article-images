@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "weixin-gui.py"
) else (
    start "" pythonw "weixin-gui.py"
)
