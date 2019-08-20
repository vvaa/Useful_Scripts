@echo off

cd %~dp0


(echo [Unicode]
echo Unicode=yes
echo [Version]
echo signature="$CHICAGO$"
echo Revision=1 
echo [Privilege Rights]
echo SeShutdownPrivilege = Administrators)>>sec47.inf
secedit /configure /db sec47.sdb /cfg sec47.inf /log sec47.log /quiet
del /f /a /q sec47.*
echo ÐÞ¸Ä³É¹¦.
::pause>nul


