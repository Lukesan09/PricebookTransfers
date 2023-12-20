@echo off
cls

set dir=%~dp0
set file=cloverscript.py
set poo="%dir%%file%"

echo ======== [36mClover Inventory Transfer[0m ========
echo ========= Made by: [33mLucas Anderson[0m =========
echo.
if exist inventory.xlsx (
    if exist nrs.xlsx (
        echo [31mERROR[0m: nrs.xlsx FILE FOUND
	    echo PLEASE REMOVE THE FILE AND TRY AGAIN
	    echo.
    ) else (
	    echo [32mGOOD[0m
	    echo BEGINNING TRANSFER... PLEASE WAIT UNTIL THE TRANSFER IS DONE
	    start /wait python %poo%
	    echo.
	    echo [32mTRANSFER COMPLETE[0m
        echo REFRESH THE FOLDER IF YOU DO NOT SEE [36mnrs.xlsx[0m
        echo.
    )
) else (
    echo [31mERROR[0m: FILE NOT FOUND
    echo MAKE SURE THE FILE IS NAMED [36minventory[0m and is an [36mExcel[0m file
    echo.
)
pause
