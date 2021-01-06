<#
.SYNOPSIS
    Get information about the installed version of PowerShell/Windows Mangement Framework

.DESCRIPTION
    Checks the given machine for the installed version of POSH/WMF

.PARAMETER  computerName
    (Default = (Get-Content 'List.txt'))

    The name of the computer you want the script to gather information about.
    This is defaulted to the a list of computers in the same directory as the
    script in a file called 'List.txt'. Can be a single computername as a
    paramter sent to the command.

.PARAMETER  csvReportPath
    (Default = 'C:Temp')

    The path to the folder you want to save the CSV report to. Defaults to the
    'C:\Temp' folder.

.PARAMETER  csvReportName
    (Default = 'PSVersion.csv')

    The name you want the csv report to be called. Defaults to 'PSVersion.csv'

.EXAMPLE
    .\Check-PSVersion.ps1

    Description:

        Gets the PS Version on the list of computers in 'List.txt and reports to the
        'C:\Temp\PSVersion.csv' file.

.EXAMPLE
    .\Check-PSVersion.ps1 computer1, server2, server3

    Description:

        Gets the PS Version on computer1, server2, and server3 and reports to the
        'C:\Temp\PSVersion.csv' file.

.EXAMPLE
    .\Check-PSVersion.ps1 -computerName Server1 -csvReportPath 'C:\Reports'

    Description:

        Gets the PS Version on Server1 and reports to the C:\Reports folder location.

.EXAMPLE
    .\Check-PSVersion.ps1 -computerName Server1 -csvReportPath 'C:\Reports' -csvReportName "Server1_Report.csv"

    Description:

        Gets the PS Version on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report.csv.

.NOTES
    Name: Check-PsVersion.ps1
    Author: Jonathan Durant
    Version: 1.0
    DateUpdated: 21 March 2017

.INPUTS
    Single object or Array of objects

.OUTPUTS
    CSV File
#>

[CmdletBinding()]
param
(
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$computerName = (Get-Content "List.txt"),
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$csvReportPath = 'C:\Temp',
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$csvReportName = 'PSVersion.csv'
)

BEGIN {
     function Test-IfNotPathCreate([string]$FolderPath) {
          # Check the passed folder path, if is does not exist create it.
          if (!(Test-Path -Path $FolderPath)) {
               New-Item -Path $FolderPath -ItemType directory | Out-Null
          }
     }

     # Set main object array for report to blank
     $Objs = @()
     # Check Report Path and create if necessary
     Test-IfNotPathCreate($csvReportPath)
}

PROCESS {
     ForEach ($computer in $computerName) {
          try {
               # Test if Computer is online and if so continue to check the PS Version Else Skip Check
               if (Test-Connection -ComputerName $computer -quiet -Count 1) {
                    # Run Invoke command on Remote machine to get PSVersion
                    $result = Invoke-Command -Computername $computer -Scriptblock {
                         $PSVersionTable.psversion
                    } -ErrorVariable $checkError 2>$null # Push off error if cannot connect
                    # Create PS object for report
                    $obj = New-Object -TypeName PSObject -Property @{
                         ComputerName = $computer
                         MajorVers = $result.major
                         MinorVers = $result.minor
                    }
                    # If PS Remoting error set object to report it.
                    If ($null -eq $result.major) {
                         $obj.MajorVers = "[!] Failed to connect"
                         $obj.MinorVers = "PowerShell Remoting Not configured"
                    }
               }
               Else {
                    # If cannot connect to computer set PS Object to report it.
                    $obj = New-Object -TypeName PSObject -Property @{
                         ComputerName = $computer
                         MajorVers = "[!] Failed to connect"
                         MinorVers = "Computer is not online"
                    }
               }
               # Update Main Object Array with current computer result
               $Objs += $obj
               # Reset $result variable
               $result = $null
          }
          Catch {
               # Throw error on host if there is a problem
               $Error
          }
     }
}

END {
     # Write Main objec tto a CSV report
     $Objs | Select-Object ComputerName, MajorVers, MinorVers | Export-csv $csvReportPath\$csvReportName -NoTypeInformation
}
