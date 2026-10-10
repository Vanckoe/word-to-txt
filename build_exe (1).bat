@echo off
setlocal
cd /d "%~dp0"
if not exist "test.py" goto fail
where py >nul 2>&1
if errorlevel 1 (
    python -m venv .venv_windows
) else (
    py -3 -m venv .venv_windows
)
if errorlevel 1 goto fail
echo Installing packages (internet and Python are required)...
".venv_windows\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto fail
echo Building exe...
".venv_windows\Scripts\python.exe" -m PyInstaller --clean --noconfirm --onefile --noconsole --name WordToTxt test.py
if errorlevel 1 goto fail
echo.
echo DONE. Your file: %~dp0dist\WordToTxt.exe
pause
exit /b 0
:fail
echo.
echo ERROR. Take a screenshot of this window and send it.
pause
exit /b 1
