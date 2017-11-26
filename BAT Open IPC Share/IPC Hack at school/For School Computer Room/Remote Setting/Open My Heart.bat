
@echo off

set myfile=my.txt

::添加防火墙例外
>nul reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 137:udP /t reg_sz /d 137:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22001 /f
>nul reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 138:udP /t reg_sz /d 138:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22002 /f
>nul reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 139:TCP /t reg_sz /d 139:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22004 /f
>nul reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 445:TCP /t reg_sz /d 445:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22005 /f

::去除远程登录管理员权限限制
::去除空密码限制
>nul reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v forceguest /t reg_dword /d 00000000 /f
>nul reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v limitblankpassworduse /t reg_dword /d 00000000 /f

for /f "delims=" %%i in ('ipconfig^|find "IP Address"')do (
echo %%i>> %myfile%
)
echo.>>%myfile%