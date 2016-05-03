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

:::::::::::::::::::::::::::::::::::::::
::Run Powershell installation script
:::::::::::::::::::::::::::::::::::::::

REM Change execution policy to allow powershell to run
@powershell "Set-ExecutionPolicy -ExecutionPolicy Bypass

REM Run installation script
@powershell -file ".\install-applications.ps1"

REM Redo execution policy
@powershell "Set-ExecutionPolicy -ExecutionPolicy Restricted"

@del %userprofile%\desktop\install-applications.ps1
@del %userprofile%\desktop\install.bat