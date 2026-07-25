@echo off
REM Inventory Manager - run with a visible console (useful for troubleshooting).
REM For everyday use, double-click "Inventory Manager.vbs" instead: no window.

cd /d "%~dp0"

set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY where py >nul 2>nul && set "PY=py -3"
if not defined PY goto nopython

echo Using: %PY%
%PY% -c "import flask" >nul 2>nul
if errorlevel 1 (
    echo Installing Flask...
    %PY% -m pip install flask
)

echo.
echo Starting Inventory Manager - your browser will open shortly.
echo Close this window or press Ctrl+C to stop the app.
echo.
%PY% app.py
pause
exit /b 0

:nopython
echo.
echo Python 3 was not found.
echo Install it from https://www.python.org/downloads/
echo and tick "Add Python to PATH" during setup.
echo.
pause
exit /b 1
