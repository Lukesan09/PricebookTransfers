@echo off
SET dir=%~dp0
SET file=zuzascript.py
SET poo="%dir%%file%"
start /wait python %poo%
echo FILE TRANSFER COMPLETE
echo REFRESH THE FOLDER IF YOU DON'T SEE THE NEW FILE
pause