@echo off
SET dir=%~dp0
SET file=cloverscript.py
SET poo="%dir%%file%"
start /wait python %poo%
echo FILE TRANSFER COMPLETE
pause