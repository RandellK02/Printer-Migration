@echo off
if defined MYSCRIPT_MINIMIZED (goto min)
set MYSCRIPT_MINIMIZED=yes
start /min cmd.exe /c %~f0 %*" 
exit

:min
start C:\Users\Public\Documents\Printer_Migration\Payload.exe %*
