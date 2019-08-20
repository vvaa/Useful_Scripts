@echo off
cd %~dp0

echo 开始 01,71 - EnableFirewall.reg
regedit /s "01,71 - EnableFirewall.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 03 - add shtele and ban guest.administrator.bat
call "03 - add shtele and ban guest.administrator.bat"
ping 127.0.0.1 -n 2 >nul

echo 开始 04 - no AutoShareServer.reg
regedit /s "04 - no AutoShareServer.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 05-07 - protect tcp.reg
regedit /s "05-07 - protect tcp.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 08 - change 3389 to 6666.reg
regedit /s "08 - change 3389 to 6666.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 09 - only administrators can remote shutdown.bat
call "09 - only administrators can remote shutdown.bat"
ping 127.0.0.1 -n 2 >nul

echo 开始 10 - only administrators can shutdown.bat
call "10 - only administrators can shutdown.bat"
ping 127.0.0.1 -n 2 >nul

echo 开始 11 - NoDriveTypeAutoRun.reg
regedit /s "11 - NoDriveTypeAutoRun.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 12-16,73 - password policy.bat
call "12-16,73 - password policy.bat"
ping 127.0.0.1 -n 2 >nul

echo 开始 17 - only administrators TakeOwnership.bat
call "17 - only administrators TakeOwnership.bat"
ping 127.0.0.1 -n 2 >nul

echo 开始 18 - restrictanonymous.reg
regedit /s "18 - restrictanonymous.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 19-26 - audit policy.bat
call "19-26 - audit policy.bat"
ping 127.0.0.1 -n 2 >nul

echo 开始 69 - EnablePMTUDiscovery.reg
regedit /s "69 - EnablePMTUDiscovery.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 72 - DisableIPSourceRouting.reg
regedit /s "72 - DisableIPSourceRouting.reg"
ping 127.0.0.1 -n 2 >nul

echo 开始 75 - no AutoAdminLogon.reg
regedit /s "75 - no AutoAdminLogon.reg"
ping 127.0.0.1 -n 2 >nul

echo 全部完成！
pause
