<#
.SYNOPSIS
     Get information about the stated computer(s)

.DESCRIPTION
     Using WMI this script will query the stated computer(s) and gather the following
     Information:

          - PC Name
          - PC Type (Virtual/Physical)
          - PC Model
          - PC Manufacturer
          - PC Type (Virtual/Physical)
          - PC Serial #
          - Last Logged On User
          - Operating System Type
          - Operating System Name
          - Operating System Version
          - Operating System Build (Windows 10, 2016/Server)
          - Operating System Architecture
          - Operating System Service Pack
          - CPU Type
          - Total Memory in GB
          - Active Domain or Workgroup Name
          - Active Ip Addresses
          - Active Network Card MAC Addresses
          - All Network Card MAC Addresses

     This information will then be displayed in the console or exported to a CSV file
     with the name supplied or using the default name (See parameters)

.PARAMETER  computerName
     (Default = $env:Computername)

     The name of the computer you want the script to gather information about.
     This is defaulted to the local computer it is running on. This parameter can
     import more than one computer inline with the command seperated by commas or by
     using the "Get-Content" CMDLET.

.PARAMETER consoleOutput
     (Default = $false)

     Switch to set whether the output should go to the Console screen or to a CSV file in the Reports folder

.PARAMETER  csvReportPath
     (Default = "$ENV:USERPROFILE\Documents\Scripts\Logs")

     The path to the folder you want to save the CSV report to. Defaults to the
     "C:\Temp" folder.

.PARAMETER  csvReportName
     (Default = "Get-ComputerInfo.csv")

     The name you want the csv report to be called. Defaults to "Get-ComputerInfo.csv"

.PARAMETER errorLogPath
     (Default = "$ENV:USERPROFILE\Documents\Scripts\Logs")

     The path to the folder you want to save the error log to. Defaults to the
     "C:\Temp" folder.

.PARAMETER  errorLogName
     (Default = "Get-ComputerInfo-ErrorLog.log")

     The name you want the error log to be called. Defaults to "Get-ComputerInfo-ErrorLog.log"

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
     .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath "C:\Reports" -csvReportName "Server1_Report.csv"
     get-exec
     Description:

          Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report.csv. The
          error log is created at the default location.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath "C:\Reports" -csvReportName "Server1_Report.csv" -errorLogPath "C:\Logs"

     Description:

          Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report.csv. The
          error log is created at the C:\Logs folder location with the standard name.

.EXAMPLE
     .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath "C:\Reports" -csvReportName "Server1_Report.csv" -errorLogPath "C:\Logs" -errorLogName "GC_Error_Log.log"

     Description:

          Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report.csv. The
          error log is created at the C:\Logs folder location with the GC_Error_Log.log name.

.NOTES
     Name: Get-ComputerInfo.ps1
     Author: Jonathan Durant
     Version: 1.5
     DateUpdated: 26 Sept 2018

.INPUTS
     Single object or Array of objects

.OUTPUTS
     CSV File, TXT File
#>

[CmdletBinding()]
param
(
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$computerName = $ENV:COMPUTERNAME,
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$csvReportPath = "$ENV:USERPROFILE\Documents\Scripts\Reports",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$csvReportName = "Get-ComputerInfo.csv",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$errorLogPath = "$ENV:USERPROFILE\Documents\Scripts\Logs",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [string[]]$errorLogName = "Get-ComputerInfo-ErrorLog.log",
     [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
     [switch]$consoleOutput = $false
)

BEGIN {
     # INTERNAL FUNCTIONS
     function Check-IfNotPathCreate([string]$FolderPath) {
          # Check the passed folder path, if is does not exist create it.
          if (!(Test-Path -Path $FolderPath)) {
               New-Item -Path $FolderPath -ItemType directory | Out-Null
          }
     }
     function Check-IfNotLogCreate([string]$logPath) {
          # Check the passed logfile path, if is does not exist create it.
          if (!(Test-Path -Path $logPath)) {
               New-Item -Path $logPath -Force -ItemType File | Out-Null
          }
     }
     function Write-ToLogFile([string]$logMessage) {
          Check-IfNotLogCreate($errorLog)
          $logMessage | Out-File -FilePath $errorLog -Append
     }
     # END INTERNAL FUNCTIONS

     # Set Script Scope Variables
     [string]$errorLog = "$errorLogPath\$errorLogName"
     [string]$csvReport = "$csvReportPath\$csvReportName"
     [string]$formattedDate = Get-Date -Format "yyyy-MM-dd_HH:mm:ss"

     # Check the report and log folder paths and create if necessary
     Check-IfNotPathCreate($csvReportPath)
     Check-IfNotPathCreate($errorLogPath)

}

PROCESS {
     # Loop to test each PC in the $computerName array
     ForEach ($pc in $computerName) {
          # Send a progress message to user in the console
          Write-Host -Object "Getting System Info for $pc"

          # Connection Tests
          # Test connections to IP address and to WMI to see if the computer is online and can communicate
          try {
               Test-Connection -ComputerName $pc -Count 1 -ErrorAction Stop | Out-Null
          }
          catch {
               $logMessage = "$formattedDate - $($pc.ToUpper()): Not online"
               Write-ToLogFile($logMessage)
               break
          }
          try {
               $osInfo = Get-WmiObject -Class Win32_operatingSystem -ComputerName $pc -ErrorAction Stop
          }
          catch {
               $logMessage = "$formattedDate - $($pc.ToUpper()): WMI failed to connect (Either is not Windows Or WMI is corrupted)"
               Write-ToLogFile($logMessage)
               break
          }

          # Main Program
          # Gather all computer info
          # Run WMI Calls
          $sys = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $pc
          $netAdapter = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $pc | Where-Object {($_.PhysicalAdapter -eq $true)}
          $netIpAddress = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $pc | Where-Object {$_.IPEnabled -eq $True}
          $RAM = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $pc
          $cpu = Get-WmiObject -Class Win32_Processor -ComputerName $pc
          $bios = Get-WmiObject -Class Win32_Bios -computer $pc

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
          if ($lastLoggedUserSAM.length -eq 0) {
               $lastLoggedUser = "***NO LAST LOGGED ON USER DATA***"
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
               Default {
                    $osBuild = "N/A v10.0 only"
               }
          }
          $osArct = $osInfo.OSArchitecture
          $spMajVer = $osInfo.ServicePackMajorVersion.ToString()
          $spMinVer = $osInfo.ServicePackMinorVersion.ToString()
          $serPackVer = "SP $spMajVer.$spMinVer"
          $cpuType = (@($cpu.Name) -join "`r`n")
          $totalRam = [Math]::Round((($ram | Measure-Object -property "Capacity" -Sum).Sum) / 1GB)
          $actIPAddr = (@($netIpAddress | Select-Object -ExpandProperty IPAddress | Where-Object {$_ -notlike "*:*"}) -join "`r`n")
          $actMacAddr = (@($netIpAddress.MACAddress) -join "`r`n")
          $allMacAddr = (@($netAdapter.MACAddress) -join "`r`n")

          # Create new custom object and add all porperties
          $sysInfo = New-Object -TypeName PSObject
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Name" -Value $pcName
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Type" -Value $pcType
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Model" -Value $pcModel
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Manufacturer" -Value $pcMfr
          $sysInfo | Add-Member -MemberType NoteProperty -Name "PC Serial #" -Value $pcSerial
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Last Logged On User" -Value $lastLoggedUser
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Operating System Type" -Value $osProdType
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $osCapt
          $sysInfo | Add-Member -MemberType NoteProperty -Name "OS Version" -Value $osVer
          $sysInfo | add-member -memberType NoteProperty -Name "OS Build" -Value $osBuild
          $sysInfo | Add-Member -MemberType NoteProperty -Name "OS Architecture" -Value $osArct
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Service Pack Version" -Value $serPackVer
          $sysInfo | Add-Member -MemberType NoteProperty -Name "CPU Type" -Value $cpuType
          $sysInfo | add-member -memberType NoteProperty -Name "Total Memory (GB)" -value $totalRam
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Active Domain Or WorkGroup" -Value $domain
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Active IP Address" -Value $actIPAddr
          $sysInfo | Add-Member -MemberType NoteProperty -Name "Active Network Adapter MAC Address" -Value $actMacAddr
          $sysInfo | Add-Member -MemberType NoteProperty -Name "ALL Network Adapter MAC Addresses" -Value $allMacAddr

          # Add custom object to results array
          [array]$results += $sysInfo
          Write-Host "1: $pcType $ram"
          Clear-Variable sys, netAdapter, netIpAddress, pcName, ram, cpu, bios, pcType, pcModel, pcMfr, pcSerial, lastLoggedUserSAM, lastLoggedUserDisplay, lastLoggedUser, osProdType, osCapt, osVer, osBuild, osArct, spMajVer, spMinVer, serPackVer, cpuType, totalRam, domain, actIPAddr, actMacAddr, allMacAddr
          write-host "2: $pcType $ram"
     }
}

END {
     # Check if results array is null, if not export to CSV file
     if ($null -ne $results) {
          if ($consoleOutput -eq $True) {
               Out-Host -InputObject $results
          }
          Else {
               $results | Export-Csv -Path $csvReport -NoTypeInformation -Force
          }
     }
}
