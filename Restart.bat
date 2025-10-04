@echo off
echo Closing all services...

taskkill /F /IM "python.exe"

echo All windows closed.

echo Restarting
cd /D C:\MB\
start Start.bat

echo Batch file completed.
exit