@echo off 

set Target=WLAN

set IP_Addr=192.168.0.99
set D_Gate=192.168.0.1
set Sub_Mask=255.255.255.0
set DNS_Srv=127.0.0.1

netsh interface ip set address "%Target%" static %IP_Addr% %Sub_Mask% %D_Gate% 1 
netsh interface ip set dns "%Target%" static %DNS_Srv%
netsh int ip show config
