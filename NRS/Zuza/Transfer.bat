@echo off
cls

set dir=%~dp0
set file=zuzascript.py
set poo="%dir%%file%"

echo ========= [36mZuza Inventory Transfer[0m =========
echo ========= Made by: [33mLucas Anderson[0m =========
echo.

if exist inventory.xlsm (
    if exist nrs.xlsx (
        echo [31mERROR[0m: nrs.xlsx FILE FOUND
	echo PLEASE REMOVE THE FILE AND TRY AGAIN
	echo.
    ) else (
	echo [32mGOOD[0m
	echo BEGINNING TRANSFER
	start /wait python %poo%
	echo.
	echo [32mTRANSFER COMPLETE[0m
    )
) else (
    echo [31mERROR[0m: FILE NOT FOUND
    echo MAKE SURE THE FILE IS NAMED [36minventory.xlsm[0m
    echo.
)

pause