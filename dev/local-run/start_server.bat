@echo off
setlocal
cd /d "%~dp0..\.."

set "PYEXE="
if exist ".venv\Scripts\python.exe" set "PYEXE=.venv\Scripts\python.exe"
if "%PYEXE%"=="" if exist "venv\Scripts\python.exe" set "PYEXE=venv\Scripts\python.exe"
if "%PYEXE%"=="" set "PYEXE=python"

where %PYEXE% >nul 2>&1
if errorlevel 1 ( echo [ERROR] Python not found & pause & exit /b 1 )

echo ========================================
echo   aiprintword  http://127.0.0.1:5050
echo   测试库: .env 中 AIWORD_ENV=test
echo   Ctrl+C 停止
echo ========================================
"%PYEXE%" app.py
pause
endlocal
