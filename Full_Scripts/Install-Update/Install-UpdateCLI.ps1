<#
.SYNOPSIS
     Pushes provided update to listed machines useing PSExec remote administration tool.

.DESCRIPTION
     Using passed parameters for the Computer name and and Hotfix Name; this script
     will copy the hotfix to the designated Computer or Computers, run the installation of the update,
     and return the exit/error code of the windows updates service.

     !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  computers
     (Required, No Default)

     This parameter is the active directory/workstation name or IP Address of the Computer that will receive the update. This designation
     can be a single Computer name, IP Address, or a function to read in a listing of Computers names, such as

     EX: "NB16576"
          "192.151.111.111"
          "(Get-Content ".\Computers.txt")" //File is located in same directory as the script//
          "(Get-Content "C:\UpdatePush\Computers.txt")"

.PARAMETER  hotFixName
     (Required, No Default)

     This parameter is the Name of the update file you wish to push, including the files extension. To ensure
     filename is read correctly due to spaces in the names, wrap the name in quotation marks.

     EX: "ms15-067 (part2)-x64.msu"

.PARAMETER hotFixPath
     (Not Required, Defaulted)

     This parameter is the directory location of the Hotfix file, exlcuding the filename and extension. (Just the file path
     up to the folder the file is located in.) This parameter is defaulted to the standard DAHC sys admin folder for
     manual patch installs. No entry is required unless you are loading an update from a seperate folder location
     than the default.

     Default Path: "C:\UpdatePush\Patches"

.PARAMETER psexecFilePath
     (Not Required, Defaulted)

     This parameter is the directory location of the PSExec.exe file, exlcuding the filename and extension. (Just the file path
     up to the folder the file is located in.) This parameter is defaulted to the standard DAHC sys admin folder for
     the PSExec.exe file. No entry is required unless you are loading the PSExec.exe file from a seperate folder location
     than the default.

     Default Path: "C:\UpdatePush\PSExec\PSExec.exe"

.EXAMPLE
     .\Install-Update.ps1 -computers WK19278 -hotFixName "ms15-067-x86.msu"

     Description:

          Installs HotFix "ms15-067-x86.msu" to the computer named "WK19278"

.EXAMPLE
     .\Install-Update.ps1 -computers 192.151.112.125 -hotFixName "ms15-067-x86.cab"

     Description:

          Installs HotFix "ms15-067-x86.cab" to the computer with IP "192.151.112.125"

.EXAMPLE
     .\Install-Update.ps1 -computers (Get-Content ".\Computers.txt") -hotFixName "ms15-067-x64.msu"

     Description:

          Using the powershell cmdlet "Get-Content" this example is reading a list of computers
          from the "computers.txt" file located in the current folder and loading them into a
          Powershell Object Array for processing. The script will then cycle through each object
          and install the "ms15-067-x64.msu" updates to each one.

.EXAMPLE
     .\Install-Update.ps1 -computers (Get-Content "c:\temp\computers.txt") -hotFixName "ms15-067-x64.exe" -hotFixPath "c:\patches\ms\windows\version\7" -psexecFilePath "c:\psexecfolder"

     Description:
          Using the powershell cmdlet "Get-Content" this example is reading a list of computers
          from the "computers.txt" file located in the c:\temp folder and loading them into a
          Powershell Object Array for processing. This will use the hotfix located in the
          c:\patches\ms\windows\version\7 folder, and use the PSExec.exe
          file located in the c:\psexecfolder Folder.

.NOTES
     Name: Install-Update
     Author: Jonathan Durant
     Version: 2.0
     DateUpdated: 2017-06-06

.INPUTS
     Computer names, Hotfix Names, file paths

.OUTPUTS
     Error/Exit codes to show completion status.
#>

[cmdletbinding()]

Param (
     [Parameter(mandatory = $true, Position = 0)]
     [array]$computers = "",
     [Parameter(mandatory = $true, Position = 1)]
     [string]$hotFixName = "",
     [Parameter(mandatory = $false, Position = 2)]
     [string]$hotFixPath = "C:\UpdatePush\Patches",
     [Parameter(mandatory = $false, Position = 3)]
     [String]$psexecFilePath = "C:\UpdatePush\PSExec\PSExec.exe"
)
process {

     $hotFixFileExt = $hotFixName.substring($hotFixName.length - 3, 3)

     if ($hotFixFileExt -eq "exe" -or $hotFixFileExt -eq "msu" -or $hotFixFileExt -eq "cab") {
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
                    Copy-Item $hotFixPath\$hotFixName "\\$computer\c$\Temp"
                    switch ($hotFixFileExt) {
                         exe {
                              & $psexecFilePath -s \\$computer "C:\Temp\$hotFixName" /q /norestart
                         }
                         msu {
                              & $psexecFilepath -s \\$computer wusa "C:\Temp\$hotFixName" /quiet /norestart
                         }
                         cab {
                              & $psexecFilePath -s \\$computer Dism.exe /Online /Add-Package /Quiet /NoRestart /packagepath:"C:\Temp\$hotFixName"
                         }
                    }
                    # Delete local copy of update package
                    Remove-Item "\\$computer\c$\Temp\$hotFixName"
                    Write-Host "`n$computer Complete, Check error code file for explanation of results...`n"
                    Write-Host "+++++++++++++++++++"
               }
          }
     }
     else {
          Write-Host "`n$hotFixName is of a unknown update filetype, please run manually on each computer. Moving on...."
          break
     }
}
