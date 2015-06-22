@ECHO OFF
ECHO.
ECHO.
ECHO ============================================
ECHO The UserModGUI Utility requires the Active Directory Module for Windows PowerShell to be installed.
ECHO Dism.exe /Online /Enable-Feature /FeatureName:RemoteServerAdministrationTools-Roles-AD-Powershell
ECHO ============================================
ECHO.
ECHO.
ECHO.Launching script, please wait.

%~dp0Common\ShellRunas.exe /reg /quiet /accepteula

%~dp0Common\ShellRunas powershell.exe %~dp0Common\Ver5\UserModGUIv5_02.ps1 /quiet -noprofile

%~dp0Common\ShellRunas.exe /unreg /quiet 



