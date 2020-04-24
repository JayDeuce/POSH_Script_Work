<#
.SYNOPSIS
     Get information about the stated computer(s)

.DESCRIPTION
     Using WMI this script will query the stated computer(s) and gather the following
     Information:

          - PC Name
          - PC Serial #
          - PC Manufacturer
          - PC Model     
          - PC Type (Virtual/Physical)          
          - Last Logged On User
          - Operating System Type
          - Operating System Name
          - Operating System Version
          - Operating System Build (Windows 10, 2016/Server, Etc.)
          - Operating System Architecture
          - Operating System Service Pack Version
          - CPU name
          - CPU # of Cores
          - CPU # of logical Processors
          - Total Memory in GB
          - Active Domain or Workgroup Name
          - Active IP Addresses
          - Active Network Card MAC Addresses

     This information will then be displayed in the console or exported to a CSV file
     with the name supplied or using the default name (See parameters)

.PARAMETER  computerName
     (Default = $env:Computername)

     The name of the computer you want the script to gather information about.
     This is defaulted to the local computer it is running on. This parameter can
     import more than one computer inline with the command seperated by commas or by
     using the "Get-Content" CMDLET.

.PARAMETER  csvReportPath
     (Default = "$ENV:USERPROFILE\Documents\Scripts\Logs")

     The path to the folder you want to save the CSV report to. Defaults to the
     "C:\Temp" folder.

.PARAMETER  csvReportName
     (Default = "Get-ComputerInfo")

     The name you want the csv report to be called, Script automtically adds Date and time to end of filename. 
     Defaults to "Get-ComputerInfo-2020-04-13_10:31:41" (Date/Time change based on Sysem Date Variable when ran)

.PARAMETER errorLogPath
     (Default = "$ENV:USERPROFILE\Documents\Scripts\Logs")

     The path to the folder you want to save the error log to. Defaults to the
     "C:\Temp" folder.

.PARAMETER  errorLogName
     (Default = "Get-ComputerInfo-ErrorLog")

     The name you want the error log to be called. Defaults to "Get-ComputerInfo-ErrorLog"

.PARAMETER consoleOutput
     (Default = $false)

     Switch to set whether the output should go to the Console screen instead of a file
     
.PARAMETER xlsxOutput
     (Default = $false)

     Switch to set whether the output should be exported to an XLSX (Excel) file in the Reports folder
     Instead of CSV

.EXAMPLE
     .\Get-ComputerInfo.ps1

     Description:

          Gathers info on the local machine and reports a CSV in the standard folder locations. The
          error log is also created at the default location.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -consoleOutput

     Description:
          Gathers info on the local machine and reports to the console window. The
          error log is also created at the default location.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -xlsxOutput

     Description:
          Gathers info on the local machine and reports toand XLXS (Excel) File in the standard folder locations.
          The error log is also created at the default location.


.EXAMPLE
     .\Get-ComputerInfo.ps1 computer1, server2, server3

     Description:

          Gathers info on computer1, server2, and server3 and reports to the standard folder locations. The
          error log is also created at the default location.

.EXAMPLE
     .\Get-ComputerInfo.ps1 (Get-Content "computers.txt")

     Description:

          Gathers info on all computers listed in the computers.txt file, reading one line at a time,
          and reports to the standard folder locations. The error log is also created at the default location.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath "C:\Reports"

     Description:

          Gathers info on Server1 and reports to the C:\Reports folder location. The
          error log is created at the default location.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath "C:\Reports" -csvReportName "Server1_Report"
     get-exec
     Description:

          Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report-2020-04-13_10:31:41. The
          error log is created at the default location.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath "C:\Reports" -csvReportName "Server1_Report" -errorLogPath "C:\Logs"

     Description:

          Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report-2020-04-13_10:31:41. The
          error log is created at the C:\Logs folder location with the standard name.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath "C:\Reports" -csvReportName "Server1_Report" -errorLogPath "C:\Logs" -errorLogName "GC_Error_Log"

     Description:

          Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report2020-04-13_10:31:41. The
          error log is created at the C:\Logs folder location with the GC_Error_Log name.

.NOTES
     Name: Get-ComputerInfo.ps1
     Author: Jonathan Durant
     Version: 1.8
     DateUpdated: 13 Apr 2020

.INPUTS
     Single object or Array of objects

.OUTPUTS
     CSV File, LOG File
#>

[CmdletBinding()]
param
(
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$computerName = $ENV:COMPUTERNAME,
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$csvReportPath = "$ENV:USERPROFILE\Documents\ScriptReports",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$csvReportName = "Get-ComputerInfo",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$errorLogPath = "$ENV:USERPROFILE\Documents\ScriptLogs",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$errorLogName = "Get-ComputerInfo-ErrorLog",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [switch]$consoleOutput = $false,
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [switch]$xlsxOutput = $false
)

BEGIN {
     # INTERNAL FUNCTIONS
     function Test-IfNotPathCreate([string]$FolderPath) {
          # Check the passed folder path, if is does not exist create it
          if (!(Test-Path -Path $FolderPath)) {
               New-Item -Path $FolderPath -ItemType directory | Out-Null
          }
     }
     function Test-IfNotLogCreate([string]$logPath) {
          # Check the passed logfile path, if is does not exist create it.
          if (!(Test-Path -Path $logPath)) {
               New-Item -Path $logPath -Force -ItemType File | Out-Null
          }
     }
     function Write-ToLogFile([string]$logMessage) {
          Test-IfNotLogCreate($errorLog)
          $logMessage | Out-File -FilePath $errorLog -Append
     }

     function Get-FormattedDate() {
          [string]$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
          return $date
     }

     function Convert-CsvToXls {

          Param (
               [string]$inputFilePath,
               [string]$outputFilePath
          )

          # Create a new Excel Workbook with one empty sheet
          $excel = New-Object -ComObject excel.application
          $excel.DisplayAlerts = $false
          $workbook = $excel.Workbooks.Add(1)
          $worksheet = $workbook.worksheets.Item(1)

          # Build the QueryTables.Add command
          # QueryTables does the same as when clicking "Data » From Text" in Excel
          $txtConnector = ("TEXT;" + $inputFilePath)
          $connector = $worksheet.QueryTables.add($txtConnector, $worksheet.Range("A1"))
          $query = $worksheet.QueryTables.item($connector.name)

          # Set the delimiter (, or ;) according to your regional settings
          $query.TextFileOtherDelimiter = $excel.Application.International(5)

          # Set the format to delimited and text for every column
          # A trick to create an array of 2s is used with the preceding comma
          $query.TextFileParseType = 1
          $query.TextFileColumnDataTypes = , 2 * $worksheet.Cells.Columns.Count
          $query.AdjustColumnWidth = 1


          # Execute & delete the import query
          $query.Refresh()
          $query.Delete()

          # Save & close the Workbook as XLSX. Change the output extension for Excel 2003
          $workbook.SaveAs($outputFilePath, 51)
          $excel.Quit()
     }

     # END INTERNAL FUNCTIONS

     # Set Script Scope Variables
     if ($xlsxOutput -eq $True) {
          [string]$xlsxReport = "$csvReportPath\$csvReportName" + "-" + (Get-FormattedDate) + ".xlsx"
     }
     else {
          [string]$csvReport = "$csvReportPath\$csvReportName" + "-" + (Get-FormattedDate) + ".csv"
     }
     [string]$errorLog = "$errorLogPath\$errorLogName.log"
     
     # Check the report and log folder paths and create if necessary
     Test-IfNotPathCreate($csvReportPath)
     Test-IfNotPathCreate($errorLogPath)

}

PROCESS {

     # Loop to test each PC in the $computerName array
     ForEach ($pc in $computerName) {
          [string]$formattedDate = Get-FormattedDate
          
          # Send a progress message to user in the console
          Write-Host -Object "Getting System Info for $pc"

          # Connection Tests
          # Test connections to IP address and to WMI to see if the computer is online and can communicate
          try {
               Test-Connection -ComputerName $pc -Count 1 -ErrorAction Stop | Out-Null
          }
          catch {
               $logMessage = "$formattedDate - $($pc.ToUpper()): Not online"
               Write-Host $logMessage
               Write-ToLogFile($logMessage)
               Add-Content "$errorLogPath\fail_list.txt" $pc
               continue
          }
          try {
               $cimSession = New-CimSession -ComputerName $pc -ErrorAction Stop
          }
          catch {
               $logMessage = "$formattedDate - $($pc.ToUpper()): CIM failed to connect, usaully due to DNS issues or you do not have Administrator rights to the remote machine"
               Write-Host $logMessage
               Write-ToLogFile($logMessage)               
               Add-Content "$errorLogPath\fail_list.txt" $pc
               continue
          }
          
          # Main Program
          # Gather all computer info
          # Run CIM Calls
          $osInfo = Get-CimInstance -ClassName Win32_operatingSystem -CimSession $cimSession
          $sys = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $cimSession -Property Name, Model, Manufacturer, Domain
          #$netAdapter = Get-CimInstance -ClassName Win32_NetworkAdapter -CimSession $cimSession | Where-Object { ($_.PhysicalAdapter -eq $true) }
          $netIpAddress = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -CimSession $cimSession | Where-Object { $_.IPEnabled -eq $True }
          $RAM = Get-CimInstance -ClassName Win32_PhysicalMemory -CimSession $cimSession
          $cpu = Get-CimInstance -ClassName Win32_Processor -CimSession $cimSession
          $bios = Get-CimInstance -ClassName Win32_Bios -CimSession $cimSession

          # Set data info variables
          $pcName = $sys.Name
          $pcModel = $sys.Model
          $pcMfr = $sys.Manufacturer
          $pcSerial = $bios.Serialnumber
          if ($sys.model.ToUpper().contains("VIRTUAL")) {
               $pcType = "Virtual Machine"
          }
          else {
               $pcType = "Physical"
          }
          $domain = $sys.Domain
          $lastLoggedUserSAM = (Get-ChildItem "\\$pc\c$\Users" | Sort-Object LastWriteTime -Descending | Select-Object Name -first 1).Name
          if ($lastLoggedUserSAM -notmatch '^(\d{10})\.(civ|mil|wa|oa|sa|ln)') {
               $lastLoggedUser = $lastLoggedUserSAM
          }
          else {
               $lastLoggedUserDisplay = ([adsisearcher]"samaccountname=$lastLoggedUserSAM").FindOne().Properties["displayname"] 
               $lastLoggedUser = $lastLoggedUserDisplay.Item(0).ToString()
          }
          $osCapt = $osInfo.Caption
          $osVer = $osInfo.Version
          if ($osInfo.ProductType -eq '1') {
               $osProdType = "Workstation"
          }
          Elseif ($osInfo.ProductType -eq '2') {
               $osProdType = "Domain Controller"
          }
          Else {
               $osProdType = "Member Server"
          }
          switch ($osVer) {
               "10.0.10240" {
                    $osBuild = "1507"
               }
               "10.0.10586" {
                    $osBuild = "1511"
               }
               "10.0.14393" {
                    $osBuild = "1607"
               }
               "10.0.15063" {
                    $osBuild = "1703"
               }
               "10.0.16299" {
                    $osBuild = "1709"
               }
               "10.0.17134" {
                    $osBuild = "1803"
               }
               "10.0.17763" {
                    $osBuild = "1809"
               }
               "10.0.18362" {
                    $osBuild = "1903"
               }
               "10.0.18363" {
                    $osBuild = "1909"
               }
               "10.0.19041" {
                    $osBuild = "2004"
               }
               Default {
                    $osBuild = "N/A v10.0 only"
               }
          }
          $osArct = $osInfo.OSArchitecture
          $spMajVer = $osInfo.ServicePackMajorVersion.ToString()
          $spMinVer = $osInfo.ServicePackMinorVersion.ToString()
          $serPackVer = "SP $spMajVer.$spMinVer"
          $cpuName = $cpu.Name | Get-Unique 
          $cpuCores = $cpu.NumberOfCores | Get-Unique
          $cpuLogProc = $cpu.NumberOfLogicalProcessors | Get-Unique
          $totalRam = [Math]::Round((($ram | Measure-Object -property "Capacity" -Sum).Sum) / 1GB)
          $actIPAddr = (@($netIpAddress | Select-Object -ExpandProperty IPAddress | Where-Object { $_ -notlike "*:*" }) -join ", ")
          $actMacAddr = (@($netIpAddress.MACAddress) -join ", ")
          #$allMacAddr = (@($netAdapter.MACAddress) -join ", ")

          # Create new custom object and add all porperties
          $sysInfo = New-Object -TypeName PSObject
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Name" -Value $pcName
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Serial #" -Value $pcSerial
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Manufacturer" -Value $pcMfr
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Model" -Value $pcModel
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Type" -Value $pcType
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Last Logged On User" -Value $lastLoggedUser
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Operating System Type" -Value $osProdType
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $osCapt
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Operating System Version" -Value $osVer
          $sysInfo | add-member -memberType NoteProperty -Name "Operating System Build" -Value $osBuild
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Operating System Architecture" -Value $osArct
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Service Pack Version" -Value $serPackVer
          $sysInfo | Add-Member -MemberType NoteProperty -Name "CPU Name" -Value $cpuName
          $sysInfo | Add-Member -MemberType NoteProperty -Name "CPU # of Cores" -Value $cpuCores
          $sysInfo | Add-Member -MemberType NoteProperty -Name "CPU # of Logical Processors" -Value $cpuLogProc
          $sysInfo | add-member -memberType NoteProperty -Name "Total Memory (GB)" -value $totalRam
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Active Domain Or WorkGroup" -Value $domain
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Active IP Address" -Value $actIPAddr          
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Active Network Adapter MAC Address" -Value $actMacAddr
          #$sysInfo | Add-Member -MemberType NoteProperty -Name "ALL Network Adapter MAC Addresses" -Value $allMacAddr

          # Add custom object to results array
          [array]$results += $sysInfo
          Get-CimSession | Remove-CimSession
          Clear-Variable formattedDate, sys, netIpAddress, pcName, ram, cpu, bios, pcType, pcModel, pcMfr, pcSerial, osProdType, osCapt, osVer, osBuild, osArct, spMajVer, spMinVer, serPackVer, cpuName, cpuCores, cpuLogProc, totalRam, domain, actIPAddr, actMacAddr, cimSession, lastLoggedUserSAM, lastLoggedUserDisplay, lastLoggedUser #allMacAddr, netAdapter,
     }
}

END {
     # Check if results array is null, if not export to CSV file
     if ($null -ne $results) {
          if ($consoleOutput -eq $True) {
               Out-Host -InputObject $results
          }
          Elseif ($xlsxOutput -eq $true) {
               $tempCSV = "$env:TEMP\Get-ComputerInfoTempReport.csv"
               $results | Export-Csv -Path $tempCSV -NoTypeInformation -Force
               Convert-CsvToXls -inputFilePath $tempCSV -outputFilePath $xlsxReport
               Remove-Item -Path $tempCSV
          }
          Else {
               
               $results | Export-Csv -Path $csvReport -NoTypeInformation -Force
          }
     }
}
