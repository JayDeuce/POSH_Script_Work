<#
.SYNOPSIS
    Pushes and runs script on listed machines useing PSExec remote administration tool.

.DESCRIPTION
    Using passed parameters for the Computer name and and script Name; this script
    will copy the script to the designated Computer or Computers, run the script,
    and return the exit/error code of the powershell.exe command.

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  computers
    (Required, No Default)

    This parameter is the active directory/workstation name or IP Address of the Computer that will receive the script. This designation
    can be a single Computer name, IP Address, or a function to read in a listing of Computers names, such as

    EX: "Workstation1"  
        "192.168.1.24"  
        "(Get-Content "Computers.txt")" //File is located in same directory as the script//
        "(Get-Content "C:\Temp\List\Computers.txt")"

.PARAMETER  scriptName
    (Required, No Default)

    This parameter is the Name of the script file you wish to push, including the file extension. To ensure
    filename is read correctly due to spaces in the names, wrap the name in quotation marks.

    EX: "ScriptToRun.ps1"

.PARAMETER scriptPath
    (Not Required, Defaulted)

    This parameter is the directory location of the script file, exlcuding the filename and extension. (Just the file path
    up to the folder the file is located in.) This parameter is defaulted to the standard DAHC sys admin folder for
    scripts to run. No entry is required unless you are loading an update from a seperate folder location
    than the default.

    Default Path: "C:\Scripts"

.PARAMETER psexecFilePath
    (Not Required, Defaulted)

    This parameter is the directory location of the PSExec.exe file, exlcuding the filename and extension. (Just the file path
    up to the folder the file is located in.) This parameter is defaulted to the standard DAHC sys admin folder for
    the PSExec.exe file. No entry is required unless you are loading the PSExec.exe file from a seperate folder location
    than the default.

    Default Path: "C:\Scripts\PSExec.exe"

.EXAMPLE
    .\Run-PS1Script.ps1 -computers Workstation1 -scriptName "ScriptToRun.ps1"

    Description:

        Runs the script "ScriptToRun.ps1" on the computer named "Workstation1"

.EXAMPLE
    .\Run-PS1Script.ps1 -computers 192.168.112.125 -scriptName "ScriptToRun.ps1"

    Description:

        Runs the script "ScriptToRun.ps1" on the computer with IP "192.168.112.125"

.EXAMPLE
    .\Run-PS1Script.ps1 -computers (Get-Content "Computers.txt") -scriptName "ScriptToRun.ps1"

    Description:

        Using the powershell cmdlet "Get-Content" this example is reading a list of computers
        from the "computers.txt" file located in the current folder and loading them into a
        Powershell Object Array for processing. The script will then cycle through each object
        and run the "ScriptToRun.ps1" script on each one.

.EXAMPLE
    .\Run-PS1Script.ps1 -computers (Get-Content "c:\temp\computers.txt") -scriptName "ScriptToRun.ps1" -scriptPath "c:\scripts\ms\windows\version\7" -psexecFilePath "c:\psexecfolder"

    Description:
        Using the powershell cmdlet "Get-Content" this example is reading a list of computers
        from the "computers.txt" file located in the c:\temp folder and loading them into a
        Powershell Object Array for processing. This will use the script located in the 
        c:\scripts\ms\windows\version\7 folder, and use the PSExec.exe
        file located in the c:\psexecfolder Folder.

.NOTES
    Name: Run-PS1Script.ps1
    Author: Jonathan Durant
    Version: 1.0
    DateUpdated: 2016-04-20

.INPUTS
    Computer names, Script Names, file paths

.OUTPUTS
    Error/Exit codes to show completion status.
#>

[cmdletbinding()]

Param (
    [Parameter(mandatory=$true,Position=0)]
    [array]$computers = "",
    [Parameter(mandatory=$true,Position=1)]
    [string]$scriptName = "",
    [Parameter(mandatory=$false,Position=2)]
    [string]$scriptPath = "C:\ScriptsToRun",
    [Parameter(mandatory=$false,Position=3)]
    [String]$psexecFilePath = "C:\Scripts\PSExec.exe"
)
process {
    foreach ($computer in $computers) {
        Write-Host "`nProcessing $computer..."
        if (!(test-connection -count 1 -Quiet -ComputerName $computer)) {
            Write-Host "`n$computer could not be contacted, please check make sure its online. Moving on...`n"
            Write-Host "+++++++++++++++++++"
        }
        else {            
            if (!(Test-Path -Path "\\$computer\c$\Temp")) {
                    new-item -Path "\\$computer\c$\Temp" -ItemType directory | Out-Null
            }
            Copy-Item $scriptPath\$scriptName "\\$computer\c$\Temp"
                    & $psexecFilepath -s \\$computer powershell.exe -noprofile -command "C:\Temp\$scriptName" 
            }
            # Delete local copy of update package
            Remove-Item "\\$computer\c$\Temp\$scriptName"
            Write-Host "`n$computer Complete, Check error code file for explanation of results...`n"
                Write-Host "+++++++++++++++++++"
        }
    }