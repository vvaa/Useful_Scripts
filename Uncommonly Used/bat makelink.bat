cd /d %~dp0

::删除目录
rd /s /q "C:\Users\Jack\Documents\Tencent Files"
rd /s /q "C:\Users\Jack\Documents\WeChat Files"

::建立链接
mklink /d "C:\Users\Jack\Documents\Tencent Files" "Z:\App Data\Tencent Files"
mklink /d "C:\Users\Jack\Documents\WeChat Files" "Z:\App Data\WeChat Files"

pause