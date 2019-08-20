@echo off

cd %~dp0


(echo [Unicode]
echo Unicode=yes
echo [Version]
echo signature="$CHICAGO$"
echo Revision=1 
echo [Privilege Rights]
echo seremoteshutdownprivilege = Administrators)>>sec27.inf
secedit /configure /db sec27.sdb /cfg sec27.inf /log sec27.log /quiet
del /f /a /q sec27.*
echo ÐÞ¸Ä³É¹¦.
::pause>nul