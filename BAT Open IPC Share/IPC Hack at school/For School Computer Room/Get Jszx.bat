@echo off
set /p mip=input the IP you want to connect:  10.10.
set  mip=10.10.%mip%
echo.
set muser=jszx
echo.
net use \\%mip%\ipc$ "" /user:"%muser%"

explorer \\%mip%\c$

pause>nul