@echo off
::������ip��dns��������3����������ip.txt�е�ip��������traceroute
del /f result.txt
for /f "delims=," %%i in (ip.txt) do tracert  -d -h 3 %%i >>result.txt

echo Done!
pause