@echo off
call "%~dp0PublishScript.cmd" 2>&1 | powershell.exe -NoProfile -Command "$input | Tee-Object -FilePath '%~dp0BuildLog.txt'"
pause