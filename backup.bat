@REM Script to backup a PC (call backup.vbs)
@echo off
set SDIR=\\192.7.1.8\b
cls
echo "STARTING BACKUP"
copy /y %SDIR%\backup.vbs  %TMP%\
copy /y %SDIR%\shadowspawn.* %TMP%\
copy /y %SDIR%\7z.*  %TMP%\
cd %TMP%
cls 
cscript //nologo %TMP%\backup.vbs 
pause 
exit 0 
