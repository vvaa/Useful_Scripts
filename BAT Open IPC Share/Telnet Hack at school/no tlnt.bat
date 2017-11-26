@echo off

::close telnet
set oldport=223
sc config TlntSvr start= disabled
net stop TlntSvr 
tlntadmn config port 23
reg delete "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v %oldport%:TCP /f

::del hide user
set myuser=user01
net user %myuser% /delete      
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" /v %myuser% /f

pause