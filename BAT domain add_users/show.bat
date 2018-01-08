for /f "tokens=1,2,3,4,5 delims=," %%a in (users.csv) do @echo %%a %%b %%c %%d %%e
pause