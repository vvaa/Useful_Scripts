@echo off
echo set ws = CreateObject("WScript.Shell") >temp.vbs
echo.  >>temp.vbs
echo s = "AutoStart\Appetizer\Appetizer.exe" >>temp.vbs
echo p = wscript.ScriptFullName >>temp.vbs
echo p = left(p,instrrev(p,"\")) ^& s >>temp.vbs
echo ws.CurrentDirectory = left(p,instrrev(p,"\")-1) >>temp.vbs
echo p = chr(34) ^& p ^& chr(34) >>temp.vbs
echo ws.run p >>temp.vbs
echo.  >>temp.vbs
echo s = "AutoStart\RunAsAdmin\launch.exe" >>temp.vbs
echo p = wscript.ScriptFullName >>temp.vbs
echo p = left(p,instrrev(p,"\")) ^& s >>temp.vbs
echo ws.CurrentDirectory = left(p,instrrev(p,"\")-1) >>temp.vbs
echo p = chr(34) ^& p ^& chr(34) >>temp.vbs
echo ws.run p >>temp.vbs
cscript //nologo temp.vbs
del /f /a /q temp.vbs
rd /s /q temp.vbs