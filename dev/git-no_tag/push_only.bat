@echo off
setlocal
cd /d "%~dp0..\.."

where git >nul 2>&1
if errorlevel 1 ( echo [ERROR] Git not found & pause & exit /b 1 )

set "BR="
for /f "tokens=*" %%i in ('git rev-parse --abbrev-ref HEAD 2^>nul') do set "BR=%%i"
if "%BR%"=="" set "BR=main"

git push -u origin "%BR%"
if errorlevel 1 ( pause & exit /b 1 )
echo [OK] 已 push %BR%（未打 tag）
pause
endlocal
