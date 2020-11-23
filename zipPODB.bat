:: Batch to execute a powershell file with the same name as the batch file
:: Simply copy this to the same folder as the ps1 file and give it the same name.
:: For Example:
::    fixprinter.ps1
::    fixprinter.bat
:: Running fixprinter.bat will execute the script without admin, if you need admin just swap the commented lines below

:: Admin
:: @ECHO OFF
::PowerShell.exe -NoProfile -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dpn0.ps1""' -Verb RunAs}"

:: Non-Admin
@ECHO OFF
IF "%~n0"=="runPSFile" (
    ECHO RTFM Rename this batch to match your powershell script name
	pause
	exit
)
PowerShell.exe -NoProfile -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dpn0.ps1""'}"
