@echo off

cd %~dp0


(echo [Unicode]
echo Unicode=yes
echo [Version]
echo signature="$CHICAGO$"
echo Revision=1 
echo [Privilege Rights]
echo SeTakeOwnershipPrivilege = Administrators)>>sec22.inf
secedit /configure /db sec22.sdb /cfg sec22.inf /log sec22.log /quiet
del /f /a /q sec22.*
echo ÐÞ¸Ä³É¹¦.
::pause>nul