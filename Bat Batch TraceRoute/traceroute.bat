@echo off
::不解析ip的dns，最多进行3跳，将所有ip.txt中的ip进行批量traceroute
del /f result.txt
for /f "delims=," %%i in (ip.txt) do tracert  -d -h 3 %%i >>result.txt

echo Done!
pause