<#
.SYNOPSIS
     Resolve IPs given to Hostnames.

.DESCRIPTION
     This will take the IPs passed and resolve them to their Hostnames, as well
     as check if they are pingable, and write that information to a CSV file in the C:\Temp
     directory by default. The report can also be a created in TXT format by passing the
     the textReport switch (add -textReport to the end of the command)

.PARAMETER  ipList
     (Mandatory, No Default)

     The IP list you want to resolve to hostnames. This can be a single IP, comma seperated
     list of IPs, or a list of IPs in a text file read in using the Get-Content cmdlet.

.PARAMETER  reportPath
     (Default = 'C:Temp')

     The path to the folder you want to save the report to. Defaults to the
     'C:\Temp' folder.

.PARAMETER  reportName
     (Default = 'Resolve-IPToHostname_Report')

     The name you want the report to be called minus the extension.

.PARAMETER textReport
     (Default = False)

     This is a switch to create the output report as a TXT file instead of CSV, It will default
     to False, and will turn True when added to the end of the command (add -textReport)

.EXAMPLE
     .\Resolve-IPToHostname.ps1 -ipList 192.168.1.1

     Description:

          Resolves the IP 192.168.1.1 to its hostname and checks if it can be pinged. Outputs results to
          C:\Temp\Resolve-IPToHostname_Report.csv file

.EXAMPLE
     .\Resolve-IPToHostname.ps1 -ipList 192.168.1.1, 192.168.1.2, 192.168.1.3 -reportPath "C:\IPCheck"

     Description:

          Resolves the IP 192.168.1.1, 192.168.1.2, and 192.168.1.3 to their hostnames and checks if
          they can be pinged. Outputs results to C:\IPCheck\Resolve-IPToHostname_Report.csv file

.EXAMPLE
     .\Resolve-IPToHostname.ps1 -ipList 192.168.1.1, 192.168.1.2, 192.168.1.3 -reportPath "C:\IPCheck" -reportName "My-IP-Check-List"

     Description:

          Resolves the IP 192.168.1.1, 192.168.1.2, and 192.168.1.3 to their hostnames and checks if
          they can be pinged. Outputs results to C:\IPCheck\My-IP-Check-List.csv file

.EXAMPLE
     .\Resolve-IPToHostname.ps1 -ipList (Get-Content 'iplist.txt') -reportPath "C:\IPCheck" -reportName "My-IP-Check-List"

     Description:

          Resolves all IPs listed in the iplist.txt file to their hostnames and checks if
          they can be pinged. Outputs results to C:\IPCheck\My-IP-Check-List.csv file

.EXAMPLE
     .\Resolve-IPToHostname.ps1 -ipList (Get-Content 'iplist.txt') -reportPath "C:\IPCheck" -reportName "My-IP-Check-List" -textReport

     Description:

          Resolves all IPs listed in the iplist.txt file to their hostnames and checks if
          they can be pinged. Outputs results to C:\IPCheck\My-IP-Check-List.txt file

.NOTES
     Name: Resolve-IPToHostname.ps1
     Author: Jonathan Durant
     Version: 1.0
     DateUpdated: 17 March 2016

.INPUTS
     Single object or Array of objects

.OUTPUTS
     CSV File, TXT File
#>

[CmdletBinding()]
param
(
     [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
     [string[]]$ipList,
     [Parameter(Mandatory = $false)]
     [string[]]$reportPath = 'C:\Temp',
     [Parameter(Mandatory = $false)]
     [string[]]$reportName = 'Resolve-IPToHostname_Report',
     [Parameter(Mandatory = $false)]
     [switch]$textReport # Should the report be TXT or the standard CSV
)

Begin {
     # INTERNAL FUNCTIONS

     function Check-IfNotPathCreate([string]$FolderPath) {
          # Check the past folder path, if is does not exist create it.
          if (!(Test-Path -Path $FolderPath)) {
               New-Item -Path $FolderPath -ItemType directory | Out-Null
          }
     }

     # END INTERNAL FUNCTIONS

     # Check the report and log folder paths and create if necessary
     Check-IfNotPathCreate($reportPath)
     # Set Script Scopes Variables

}
Process {
     ForEach ($ip in $ipList) {

          # Set PingCheck to no as default
          $pingCheck = "No"

          # Create Answer Object to hold IP info
          $ipAnswerList = New-Object -TypeName PSObject

          # Check to see if the IP is pingable, if so set $pingCheck to Yes
          if (Test-Connection -ComputerName $ip -quiet -Count 1) {
               $pingCheck = "Yes"
          }
          Try {
               if (!($pingCheck -eq "Yes")) {
                    Throw "IP Not Pingable, Skipping Resolving of the Hostname"
               }

               $hostName = ([system.net.Dns]::GetHostByAddress($ip)).hostname

               #If $hostname command had no errors, fill out the info for this IP
               $ipAnswerList  | Add-Member -MemberType NoteProperty -Name 'IP' -Value $ip
               $ipAnswerList  | Add-Member -MemberType NoteProperty -Name 'Resolved Name' -Value $hostName
               $ipAnswerList  | Add-Member -MemberType NoteProperty -Name 'Answers Ping' -Value $pingCheck
          }
          Catch {
               # if $hostName command had errors, fill in the info for a non resolvable IP, include if it can be ping or not.
               $ipAnswerList  | Add-Member -MemberType NoteProperty -Name 'IP' -Value $ip
               $ipAnswerList  | Add-Member -MemberType NoteProperty -Name 'Resolved Name' -Value "Cannot Resolve Host"
               $ipAnswerList  | Add-Member -MemberType NoteProperty -Name 'Answers Ping' -Value $pingCheck
          }
          # Place IP info Answers into a single variable array.
          [array]$results += $ipAnswerList
     }
}

End {
     if ($null -ne $results) {
          # Check which report to create and create it.
          If ($textReport -eq $true) {
               $results | Out-File -FilePath "$reportPath\$reportName.txt"
          }
          Else {
               $results | Export-Csv -Path "$reportPath\$reportName.csv" -NoTypeInformation
          }
     }
}