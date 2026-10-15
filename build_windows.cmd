@echo off
cd /d "%~dp0"
py -3 -m pip install -r requirements-dev.txt
py -3 -m PyInstaller --noconfirm --clean --windowed --onedir --name Xuheng main.py
xcopy samples dist\Xuheng\samples /E /I /Y
pause
