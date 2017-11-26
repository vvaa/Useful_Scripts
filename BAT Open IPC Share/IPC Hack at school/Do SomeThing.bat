@echo off
set /p mip=His Ip:  192.168.1.
set mip=192.168.1.%mip%
set /p mcmd=Give me your command:  

net time \\%mip%>temp.txt
set /p strout=<temp.txt
del /f temp.txt
set mflag=%strout:~-8,2%
set mhou=%strout:~-5,2%
set mmin=%strout:~-2%
if "%mflag%"=="ÏÂÎç" (set mt=pm) else (set mt=am)
set /a mmin+=1
set histime=%mhou%:%mmin%%mt%

at \\%mip% %histime% "%mcmd%"
echo.
at \\%mip%
echo.

set /p answer=will you keep it?(input "y" or "n"):
if "%answer%"=="n" ( at \\%mip% /del /yes
                     echo All task has been quilted.
)else if "%answer%"=="y" (echo Well,It will wait to work!
)else echo Wrong Input! It will be kept!

echo.
at \\%mip%
pause >nul