# Developed by Deepak Mourya, Information Security Consltant   

@echo off

echo Enter IP address of server:
echo.
set IPAddrs=""
set /p IPAddrs=

IF exist %IPAddrs%_websites.txt (del %IPAddrs%_websites.txt)
IF exist %IPAddrs%_details.txt (del %IPAddrs%_details.txt)
IF exist listVDsPath.txt (del listVDsPath.txt)
IF exist templistVDsPath.txt (del templistVDsPath.txt)

REM Creating list of VDs and its physical path
ECHO --------------------- > %IPAddrs%_websites.txt 
cscript list_Vdir.vbs >> %IPAddrs%_websites.txt
for /F "skip=1" %%i in ('type "%IPAddrs%_websites.txt"') do (@echo FOR WEBSITE :- "%%i" & cscript fetch-path-details.vbs "%%i" ) | findstr /I /V "Microsoft Copyright" >> %IPAddrs%_details.txt

goto :eof

REM filtering "Inetpub W3SVC PerlEx -" from output
more %IPAddrs%_details.txt | findstr /V /I "Inetpub W3SVC PerlEx -" > templistVDsPath.txt

REM Removing blank lines from templistVDsPath.txt file
For /F "tokens=* delims=" %%A in (templistVDsPath.txt) Do Echo %%A >> listVDsPath.txt

REM Searching "zip sql bat config copy backup html" key values in VD-folders
for /f "tokens=*" %%x in ('type "listVDsPath.txt"') do (
dir /S /B %%x | findstr /I "zip sql bat copy backup exe" >> %IPAddrs%_Unsafe.txt	
)

del templistVDsPath.txt
del %IPAddrs%_websites.txt
del listVDsPath.txt

pause
