@echo off 

set Target=WLAN

set DNS_Srv=114.114.114.114

netsh interface ip set dns "%Target%" static %DNS_Srv%
netsh int ip show config
