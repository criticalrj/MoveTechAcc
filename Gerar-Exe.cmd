@echo off
rem Builds dist\MoveTA.exe on this PC with the Windows csc.exe (same build that produced v2.2).
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0build\Build-Exe.ps1"
pause
