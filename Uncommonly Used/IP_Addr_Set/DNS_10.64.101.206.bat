@echo off 

set Target=WLAN

set DNS_Srv=10.64.101.206

netsh interface ip set dns "%Target%" static %DNS_Srv%
netsh int ip show config
