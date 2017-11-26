@echo off
set /p mip=input the IP:192.168.1.
set mip=192.168.1.%mip%
net use \\%mip%\ipc$ /del
net use \\%mip%\c$ /del
net use \\%mip%\d$ /del
pause