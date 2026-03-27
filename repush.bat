@echo off
REM Re-push only: no add/commit. Uses current branch (default main if detection fails).
setlocal
cd /d "%~dp0"

where git >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git not found. Install Git and add it to PATH.
    goto END
)

git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Not a git repository.
    goto END
)

set "BR="
for /f "tokens=*" %%i in ('git rev-parse --abbrev-ref HEAD 2^>nul') do set "BR=%%i"
if "%BR%"=="" set "BR=main"
if "%BR%"=="HEAD" (
    echo [ERROR] Detached HEAD. Checkout a branch first.
    goto END
)

echo === git push -u origin %BR% ===
git push -u origin "%BR%"
if errorlevel 1 (
    echo [ERROR] Push failed. Check network, proxy, or GitHub auth.
    goto END
)

echo [OK] Pushed branch %BR% to origin.

:END
echo.
echo ========================================
echo   Done. Press any key to close this window.
echo ========================================
pause
endlocal
