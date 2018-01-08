@echo off
cd /d %~dp0

for /f "tokens=1,2,3,4,5 delims=," %%a in (users.csv) do dsadd user "cn=%%c,ou=XXXX,ou=YYYY,dc=XXXX,dc=YYYY,dc=ZZZZ" -samid %%d -upn %%d -ln %%a -fn %%b -pwd %%e -disabled no -pwdneverexpires yes
pause