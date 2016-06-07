@ECHO OFF

PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "%~dp0PS2EXE.ps1 -inputfile .\Source\Send-Message.ps1 -outputFile \Source\send-message.exe -iconFile .\Source\send-message.ico -noConsole -runtime20 -verbose"

pause