@echo off
set /p mip=input the IP you want to connect:  10.10.
set  mip=10.10.%mip%
echo.
set /p muser=input the Username:  
echo.
set /p mpword=input the password:  
echo.
net use \\%mip%\ipc$ "%mpword%" /user:"%muser%"
net use M: \\%mip%\c$
net use N: \\%mip%\d$ 

pause>nul