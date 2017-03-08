<#
.SYNOPSIS
    Looks at the given machine and finds out if the given software is installed and outputs results to CSV, or
    if no software name is given all software on the machine is listed.

.DESCRIPTION
    Using passed parameters for the Computer name and (if needed) Software Name; this script
    will use WMI to gather the installed software on the designated Computer or Computers,
    and output the results to a CSV file. If Software name is given, ONLY that software is
    searched for and outputted to the CSV file.

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  computers
    (Required, No Default)

    This parameter is the active directory/workstation name or IP Address of the Computer that will run the search. This designation
    can be a single Computer name, IP Address, or a function to read in a listing of Computers names, such as

    EX: "Workstation1"
        "192.168.1.25"
        "(Get-Content "Computers.txt")" //File is located in same directory as the script//
        "(Get-Content "C:\Temp\List\Computers.txt")"

.PARAMETER  softwareName
    (Not Required, Defaulted)

    This parameter is the Name of the software you wish to push. The software name must be exactly what the software is installed
    to the system as. You can find this info in "Programs and Features" under the "NAME" column, or you can use the
    "Get-WMIObject -class Win32_Product" command on a machine with the software installed. This Parameter is defaulted to list
    all software on the machine and set to "*". Only use this parameter if you want to search for a specific software.

    EX: "Internet Explorer 11"

.PARAMETER $resultsFileOutPath
    (Not Required, Defaulted)

    This parameter is the directory location of where the results file should be created, exlcuding the filename and extension.
    (Just the file path up to the folder the file will be created in.) This parameter is defaulted to the "C:\Temp". No entry
    is required unless you are loading an update from a seperate folder location than the default.

    Default Path: "C:\Temp"

.EXAMPLE
    .\Find-SoftwareOnWorstation.ps1 -computers Workstation1

    Description:

        Searchs for and exports a list of all software on the computer named "Workstation1" to the default CSV file location at "C:\Temp"

.EXAMPLE
    .\Find-SoftwareOnWorstation.ps1 -computers 192.168.1.23

    Description:

        Searchs for and exports a list of all software on the computer IP "192.168.1.23" to the default CSV file location at "C:\Temp"

.EXAMPLE
    .\Find-SoftwareOnWorstation.ps1 -computers (Get-Content "Computers.txt") -softwareName "Internet Exporer 11"

    Description:

        Searchs for Internet Explorer 11 on all Computers listed in the Computers.txt file and exports a list to the default CSV file location at "C:\Temp"

.EXAMPLE
    .\Find-SoftwareOnWorstation.ps1 -computers (Get-Content "Computers.txt") -softwareName "Internet Exporer 11" -resultsFileOutPath "C:\MySearchs\Day1"

    Description:

        Searchs for Internet Explorer 11 on all Computers listed in the Computers.txt file and exports a list to the CSV file location at "C:\MySearchs\Day1"

.NOTES
    Name: Find-SoftwareOnWorstation.ps1
    Author: Jonathan Durant
    Version: 1.0
    DateUpdated: 2016-04-21

.INPUTS
    Computer names, Software Names, file paths

.OUTPUTS
    Results CSV File.
#>

[cmdletbinding()]

Param (
     [Parameter(mandatory=$true,Position=0)]
     [array]$computers = "",
     [Parameter(mandatory=$false,Position=1)]
     [string]$softwareName = "*",
     [Parameter(Mandatory=$false,Position=2)]
     [string]$resultsFileOutPath = "c:\temp\"
)

Begin {
     function Check-IfNotPathCreate([string]$FolderPath) {
          # Check the past folder path, if is does not exist create it.
          if (!(Test-Path -Path $FolderPath)) {
               New-Item -Path $FolderPath -ItemType directory | Out-Null
          }
     }

     Check-IfNotPathCreate($resultsFileOutPath)
}

Process {
     foreach ($computer in $computers) {

          $ping = Test-Connection -ComputerName $computer -Quiet -Count 1

          if ($ping -eq $true) {
               try {
                    if ($softwareName -eq "*") {

                         $resultsFileOutPath = "c:\temp\$computer-All-Software-Search-Results.csv"

                         Write-Host "Checking $computer"
                         Get-WmiObject -Class Win32_Product -ComputerName $computer -ErrorAction Stop | Select-Object @{N="ComputerName";E={$computer}},Vendor,Name | Export-Csv $resultsFileOutPath -Append -NoTypeInformation
                    }
                    else {

                         $outputFileName = $softwareName -replace '\s','_'
                         $resultsFileOutPath = "c:\temp\$computer-$outputFileName-Software-Search-Results.csv"

                         Write-Host "Checking $computer"
                         Get-WmiObject -Class Win32_Product -ComputerName $computer -filter "Name = '$softwareName'" -ErrorAction Stop | Select-Object @{N="ComputerName";E={$computer}},Vendor,Name | Export-Csv $resultsFileOutPath -Append -NoTypeInformation
                    }
               }
               Catch {
                    Write-Host "Unable to Obtain WMI Object of $computer"
               }
          }
     }
     if ($ping -eq $false) {
          Write-Host "The $computer is not pingable"
     }
}

