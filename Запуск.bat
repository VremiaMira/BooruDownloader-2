@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Booru Downloader - Search Test
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0main.ps1"
echo.
pause
