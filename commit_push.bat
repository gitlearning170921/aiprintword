@echo off
REM UTF-8 BOM required for Chinese echo below under chcp 65001
chcp 65001 >nul
setlocal

cd /d "%~dp0"

if "%~1"=="" (
    set "COMMIT_MSG=chore: update"
) else (
    set "COMMIT_MSG=%*"
)

where git >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found. Install Git and add it to PATH.
    goto END
)

echo === git add ===
git add -A
if errorlevel 1 (
    echo [ERROR] git add failed.
    goto END
)

echo === git status ===
git status -sb

echo === git commit ===
git commit -m "%COMMIT_MSG%"
if errorlevel 1 (
    echo [INFO] Nothing to commit, or commit failed.
    goto END
)

echo === git push origin main ===
git push -u origin main
if errorlevel 1 (
    echo [ERROR] git push failed. Check network, proxy, or GitHub auth.
    goto END
)

echo [OK] Pushed to origin main.

:END
echo.
echo ========================================
echo   Done. Press any key to close.
echo ========================================
pause
endlocal
