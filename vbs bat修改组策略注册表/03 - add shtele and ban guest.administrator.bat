


@echo off
set myuser=shtele
set mypwd=P@ssw0rd






::add  user
net user %myuser% %mypwd% /add
net user %myuser% %mypwd%
net localgroup administrators %myuser% /add


::ban  user
net user administrator /active:no
net user guest /active:no


::pause