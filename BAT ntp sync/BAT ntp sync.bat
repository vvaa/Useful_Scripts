
w32tm /config /manualpeerlist:time1.aliyun.com /syncfromflags:manual /reliable:yes /update
net stop w32time
net start w32time
w32tm /resync
w32tm /query /status
pause