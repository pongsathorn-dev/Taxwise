@echo off
setlocal EnableExtensions
set "ROOT=%~dp0"
cd /d "%ROOT%"

echo ===========================================
echo   Taxwise - Tax Data Comparison System
echo ===========================================
echo.

where node >nul 2>&1
if errorlevel 1 goto no_node
where npm.cmd >nul 2>&1
if errorlevel 1 goto no_node

echo Checking frontend packages. The first launch can take a few minutes...
pushd "%ROOT%frontend"
call npm.cmd install
if errorlevel 1 (
    popd
    goto setup_error
)
popd

echo Starting the app. Leave the server window open while using Taxwise.
start "Taxwise" /D "%ROOT%frontend" cmd /k call npm.cmd run dev
timeout /t 5 /nobreak >nul
start "" "http://localhost:5173"
echo Taxwise is opening at http://localhost:5173
pause
exit /b 0

:no_node
echo Node.js and npm were not found. Install Node.js 20 or newer, then try again.
goto failed

:setup_error
echo Setup did not complete. Check the error above, then run start.bat again.

:failed
pause
exit /b 1
