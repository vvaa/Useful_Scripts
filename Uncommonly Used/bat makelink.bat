cd /d %~dp0

::ɾ��Ŀ¼
rd /s /q "C:\Users\Jack\Documents\Tencent Files"
rd /s /q "C:\Users\Jack\Documents\WeChat Files"

::��������
mklink /d "C:\Users\Jack\Documents\Tencent Files" "Z:\App Data\Tencent Files"
mklink /d "C:\Users\Jack\Documents\WeChat Files" "Z:\App Data\WeChat Files"

pause