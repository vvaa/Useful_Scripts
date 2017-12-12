@echo off 

set Target=WLAN

netsh interface ip set address name = "%Target%" source = dhcp
netsh interface ip set dns "%Target%" source = dhcp

ipconfig /renew
netsh int ip show config
