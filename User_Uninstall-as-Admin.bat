@echo off
cd /d %~dp0
powershell -NoLogo -ExecutionPolicy RemoteSigned -File .\Setup.ps1 -ActionType Uninstall
pause
