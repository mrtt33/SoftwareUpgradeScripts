:::::::::::::::::::::::::::::::::::::::::
:: Automatically check & get admin rights
:::::::::::::::::::::::::::::::::::::::::

@echo off
CLS
ECHO.
ECHO =============================
ECHO Running Admin shell
ECHO =============================

:checkPrivileges
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges )

:getPrivileges
if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges)
ECHO.
ECHO **************************************
ECHO Invoking UAC for Privilege Escalation
ECHO **************************************

setlocal DisableDelayedExpansion
set "batchPath=%~0"
setlocal EnableDelayedExpansion
ECHO Set UAC = CreateObject^("Shell.Application"^) > "%temp%\OEgetPrivileges.vbs"
ECHO args = "ELEV " >> "%temp%\OEgetPrivileges.vbs"
ECHO For Each strArg in WScript.Arguments >> "%temp%\OEgetPrivileges.vbs"
ECHO args = args ^& strArg ^& " "  >> "%temp%\OEgetPrivileges.vbs"
ECHO Next >> "%temp%\OEgetPrivileges.vbs"
ECHO UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%temp%\OEgetPrivileges.vbs"
"%SystemRoot%\System32\WScript.exe" "%temp%\OEgetPrivileges.vbs" %*
exit /B

:gotPrivileges
if '%1'=='ELEV' shift /1
setlocal & pushd .
cd /d %~dp0

:::::::::::::::::::::::::::::::
::Find drive letter of USB
:::::::::::::::::::::::::::::::

::change directory to the script's directory's drive
pushd %~d0

::navigate from the drive to the relevant path(s)
cd Scripts

:::::::::::::::::::::::::::::::
::Run Powershell Uninstaller
:::::::::::::::::::::::::::::::

REM Changing execution to bypass
@echo. Changing POSH Execution Policy...
@echo.
@powershell -command "set-executionpolicy -ExecutionPolicy Bypass" 

REM Running script from drive found above
@powershell -file ".\Uninstall-Applications.ps1"

REM Resetting execution policy to restricted
@echo. Resetting POSH Execution Policy...
@echo.
@powershell -command "Set-ExecutionPolicy -ExecutionPolicy Restricted"

REM Copy over necessary files from USB
cd ..

@robocopy ".\Office2016" "C:\Office2016" /E /J /ETA /s /njh /njs /ndl /nc /ns
@robocopy ".\FRSCC2016" "C:\FRSCC2016" /E /J /ETA /s /njh /njs /ndl /nc /ns
@robocopy ".\Scripts" "%userprofile%\Desktop" /s /njh /njs /ndl /nc /ns /xf Uninstall-Applications.ps1 /E /J /ETA

::not required
popd