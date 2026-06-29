@echo off
chcp 65001 >nul
echo ========================================
echo   财务支出看板 - 正在启动服务...
echo ========================================
cd /d "%~dp0"
node server.js
pause
