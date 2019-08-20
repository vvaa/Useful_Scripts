@echo off

cd %~dp0


(echo [Unicode]
echo Unicode=yes
echo [Version]
echo signature="$CHICAGO$"
echo Revision=1 
echo [Event Audit]
echo AuditSystemEvents = 3
echo AuditLogonEvents = 3
echo AuditObjectAccess = 3
echo AuditPrivilegeUse = 3
echo AuditPolicyChange = 3
echo AuditAccountManage = 3
echo AuditProcessTracking = 3
echo AuditDSAccess = 3
echo AuditAccountLogon = 3)>>sec24.inf
secedit /configure /db sec24.sdb /cfg sec24.inf /log sec24.log /quiet
del /f /a /q sec24.*
echo ÐÞ¸Ä³É¹¦.
::pause>nul