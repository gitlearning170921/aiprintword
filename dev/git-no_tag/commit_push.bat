@echo off
setlocal
cd /d "%~dp0..\.."

if "%~1"=="" (set "COMMIT_MSG=chore: update") else (set "COMMIT_MSG=%*")

where git >nul 2>&1
if errorlevel 1 ( echo [ERROR] Git not found & pause & exit /b 1 )

git add -A
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo [INFO] 无新提交或 commit 失败
    pause & exit /b 1
)
git push -u origin main
if errorlevel 1 ( pause & exit /b 1 )
echo [OK] 已 push（未打 tag）
pause
endlocal
