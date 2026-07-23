@echo off
chcp 65001 >nul
where node >nul 2>nul
if errorlevel 1 (
 start https://nodejs.org/
 pause
 exit /b 1
)
start http://localhost:3010/join.html
set PORT=3010
node server.js
pause
