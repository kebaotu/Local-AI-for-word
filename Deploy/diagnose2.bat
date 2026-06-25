@echo off
echo === Local Word AI - Chan doan sau (Phase 2) ===
echo.

:: Bat VSTO Host log qua environment variable (dang tin cay hon registry)
set VSTOHOST_LOGFILE=%USERPROFILE%\Desktop\VSTO_host_log.txt

powershell -ExecutionPolicy Bypass -File "%~dp0diagnose2.ps1"
