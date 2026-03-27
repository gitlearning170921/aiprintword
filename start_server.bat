@echo off
REM Start Flask web: batch print UI on port 5050
setlocal
cd /d "%~dp0"

set "PYEXE="
if exist ".venv\Scripts\python.exe" set "PYEXE=.venv\Scripts\python.exe"
if "%PYEXE%"=="" if exist "venv\Scripts\python.exe" set "PYEXE=venv\Scripts\python.exe"
if "%PYEXE%"=="" set "PYEXE=python"

if /i "%PYEXE%"=="python" (
    where python >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found. Install Python 3.8+ and add to PATH.
        goto END
    )
) else (
    if not exist "%PYEXE%" (
        echo [ERROR] Not found: %PYEXE%
        goto END
    )
)

echo.
echo ========================================
echo   aiprintword - starting server
echo   Open: http://127.0.0.1:5050
echo   Press Ctrl+C to stop
echo ========================================
echo.

"%PYEXE%" app.py
if errorlevel 1 (
    echo.
    echo [ERROR] Server stopped with an error.
    goto END
)
goto :EOF

:END
echo.
pause
endlocal
