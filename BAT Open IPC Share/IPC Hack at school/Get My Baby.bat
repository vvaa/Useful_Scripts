@echo off
set /p mip=input the IP you want to connect:  192.168.1.
set  mip=192.168.1.%mip%
echo.
set /p muser=input the Username:  
echo.
set /p mpword=input the password:  
echo.
net use \\%mip%\ipc$ "%mpword%" /user:"%muser%"
net use M: \\%mip%\c$
net use N: \\%mip%\d$ 

pause>nul