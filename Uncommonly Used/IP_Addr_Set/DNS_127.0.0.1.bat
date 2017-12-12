@echo off 

set Target=WLAN

set DNS_Srv=127.0.0.1

netsh interface ip set dns "%Target%" static %DNS_Srv%
netsh int ip show config
