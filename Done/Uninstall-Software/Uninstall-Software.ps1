<#
.SYNOPSIS
     Uninstalls the given software from the given Computer or computers


.DESCRIPTION
     Using passed parameters for the Computer name and (if needed) Software Name; this script
     will use the WMI Class Win32_Product method Uninstall() to uninstall the given software from
     the given computer or computers

     !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  computers
     (Required, No Default)

     This parameter is the active directory/workstation name or IP Address of the Computer that will run the uninstall. This designation
     can be a single Computer name, IP Address, or a function to read in a listing of Computers names, such as

     EX: "Workstation1"
          "192.168.111.111"
          "(Get-Content "Computers.txt")" //File is located in same directory as the script//
          "(Get-Content "C:\Temp\List\Computers.txt")"

.PARAMETER  softwareName
     (Required,No Default)

     This parameter is the Name of the software you wish to uninstall. The software name must be exactly what the software is installed
     to the system as. You can find this info in "Programs and Features" under the "NAME" column, or you can use the
     "Get-WMIObject -class Win32_Product | Select-Object Name" command on a machine with the software installed.

     EX: "Internet Explorer 11"

.EXAMPLE
     .\Uninstall-Software.ps1 -computers Workstation1 -softwareName "Internet Explorer 11"

     Description:

          Uninstalls Internet Explorer 11 from the Workstation1 Machine.

.EXAMPLE
     .\Uninstall-Software.ps1 -computers 192.168.112.125 -softwareName "Internet Explorer 11"

     Description:

          Uninstalls Internet Explorer 11 from the computer with IP "192.168.112.125".

.EXAMPLE
     .\Uninstall-Software.ps1 -computers (Get-Content "Computers.txt") -softwareName "Internet Exporer 11"

     Description:

          Uninstalls Internet Explorer 11 from all Computers listed in the Computers.txt file.

.NOTES
     Name: Uninstall-Software.ps1
     Author: Jonathan Durant
     Version: 1.0
     DateUpdated: 2016-04-21

.INPUTS
     Computer names, Software Names

.OUTPUTS
     None
#>

[cmdletbinding()]

Param (
     [Parameter(mandatory = $true, Position = 0)]
     [array]$computers = "",
     [Parameter(mandatory = $false, Position = 1)]
     [string]$softwareName = ""
)

foreach ($computer in $computers) {
     $ping = Test-Connection -ComputerName $computer -Quiet -Count 1

     if ($ping -eq $true) {
          try {
               Write-Host "Uninstalling $softwareName from $computer"
               $app = Get-WmiObject -Class Win32_Product -ComputerName $computer -filter "Name = '$softwareName'" -ErrorAction Stop
               $app.uninstall()
               Write-Host "=====================================`n"
          }
          Catch {
               Write-Host "Unable to Obtain WMI Object of $computer"
               Write-Host "`n=====================================`n"
          }
     }
}
If ($Ping -eq $False) {
     Write-Host "The $computer is not pingable"
     Write-Host "`n=====================================`n"
}

