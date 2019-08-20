@echo off

cd %~dp0


(echo [Unicode]
echo Unicode=yes
echo [Version]
echo signature="$CHICAGO$"
echo Revision=1 
echo [System Access]
echo MaximumPasswordAge = 90
echo MinimumPasswordLength = 8
echo PasswordComplexity = 1
echo PasswordHistorySize = 5
echo LockoutBadCount = 5
echo ResetLockoutCount = 30
echo LockoutDuration = 120)>>sec11.inf
secedit /configure /db sec11.sdb /cfg sec11.inf /log sec11.log /quiet
del /f /a /q sec11.*
echo ÐÞ¸Ä³É¹¦.
::pause>nul