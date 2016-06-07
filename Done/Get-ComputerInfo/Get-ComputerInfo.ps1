<#
.SYNOPSIS
    Get information about the stated computer(s)

.DESCRIPTION
    Using WMI this script will query the stated computer(s) and gather the following
    Information:

        - PC Name
        - PC Model
        - PC Manufacturer
        - PC Type (Virtual/Physical)
        - PC Serial #
        - Active Domain or Workgroup Name
        - Operating System Name
        - Operating System Version
        - Operating System Architecture
        - Operating System Service Pack
        - CPU Type
        - Total Memory in GB
        - Active Ip Addresses
        - Active Network Card MAC Addresses
        - All Network Card MAC Addresses

    This information will then be exported to a CSV file with the name supplied or
    using the default name (See parameters)

.PARAMETER  computerName
    (Default = $env:Computername)

    The name of the computer you want the script to gather information about.
    This is defaulted to the local computer it is running on. This parameter can
    import more than one computer using the 'Get-Content' CMDLET.

.PARAMETER  csvReportPath
    (Default = 'C:Temp')

    The path to the folder you want to save the CSV report to. Defaults to the 
    'C:\Temp' folder.

.PARAMETER  csvReportName
    (Default = 'Get-ComputerInfo.csv')

    The name you want the csv report to be called. Defaults to 'Get-ComputerInfo.csv'

.PARAMETER errorLogPath
    (Default = 'C:Temp')

    The path to the folder you want to save the error log to. Defaults to the 
    'C:\Temp' folder.

.PARAMETER  errorLogName
    (Default = 'Get-ComputerInfo-ErrorLog.log')

    The name you want the error log to be called. Defaults to 'Get-ComputerInfo-ErrorLog.log'

.EXAMPLE
    .\Get-ComputerInfo.ps1 

    Description:

        Gathers info on the local machine and reports to the standard folder locations. The
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
    .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath 'C:\Reports'

    Description:

        Gathers info on Server1 and reports to the C:\Reports folder location. The
        error log is created at the default location.

.EXAMPLE
    .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath 'C:\Reports' -csvReportName "Server1_Report.csv"

    Description:

        Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report.csv. The
        error log is created at the default location.

.EXAMPLE
    .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath 'C:\Reports' -csvReportName "Server1_Report.csv" -errorLogPath 'C:\Logs'

    Description:

        Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report.csv. The
        error log is created at the C:\Logs folder location with the standard name.

.EXAMPLE
    .\Get-ComputerInfo.ps1 -computerName Server1 -csvReportPath 'C:\Reports' -csvReportName "Server1_Report.csv" -errorLogPath 'C:\Logs' -errorLogName 'GC_Error_Log.log'

    Description:

        Gathers info on Server1 and reports to the C:\Reports folder location with the report name of Server1_Report.csv. The
        error log is created at the C:\Logs folder location with the GC_Error_Log.log name.

.NOTES
    Name: Get-ComputerInfo.ps1
    Author: Jonathan Durant
    Version: 1.0
    DateUpdated: 14 March 2016      

.INPUTS
    Single object or Array of objects

.OUTPUTS
    CSV File, TXT File
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory=$false,ValueFromPipeline=$true)]
	[string[]]$computerName = $env:COMPUTERNAME,
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
	[string[]]$csvReportPath = 'C:\Temp',
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
	[string[]]$csvReportName = 'Get-ComputerInfo.csv',
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
	[string[]]$errorLogPath = 'C:\Temp',
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
	[string[]]$errorLogName = 'Get-ComputerInfo-ErrorLog.log'
)

BEGIN {
    # INTERNAL FUNCTIONS

    function Check-IfNotPathCreate([string]$FolderPath) {
        # Check the past folder path, if is does not exist create it.
        if (!(Test-Path -Path $FolderPath)) {
            New-Item -Path $FolderPath -ItemType directory | Out-Null
        }
    }

    # END INTERNAL FUNCTIONS

    # Set Script Scope Variables
    $errorLog = "$errorLogPath\$errorLogName"    
    $csvReport = "$csvReportPath\$csvReportName"
    $formattedDate = Get-Date -Format "yyyy-MM-dd_HH:mm:ss"

    # Check the report and log folder paths and create if necessary
    Check-IfNotPathCreate($csvReportPath)
    Check-IfNotPathCreate($errorLogPath)

    # Create Error Log File
    New-Item -Path $errorLogPath\$errorLogName -Force -ItemType File | Out-Null    
}

PROCESS {
    # Loop to test each PC in the computername array
	ForEach ($pc in $computerName) 
	{   
        # Set the Connection Test Variables to false, they will turn to True if the tests failed
        $pingCheck = $false
        $wmiCheck = ''

        # Send a message to user in the console
        Write-Host -Object "Getting computers info for $pc"
        
        # Test connections to IP address and to WMI to see if the computer is online and can communicate
        try {            
            if (Test-Connection -ComputerName $pc -quiet -Count 1){
                $pingCheck = $true # Change to True if computer responds
            }
            # Check WMI, Stop and go to Catch if there is an error
            $osInfo = Get-WmiObject -Class Win32_operatingSystem -ComputerName $pc -ErrorAction Stop 
            $wmiCheck = $true # Change to True if computer responds        
        }
            catch {
                $wmiCheck = $false # Change to False if WMI fails
        }

        # Main program procedure - Check both Ping and WMI Checks is passes moves to gather all data.
        if ($pingCheck -eq $false) { 
            # Check Ping test variable and write log entry if matches False
            "$formattedDate - $($pc.ToUpper()): Not online" | Out-File -FilePath $errorLog -Append
        }
        elseif ($wmiCheck -eq $false) {
            # Check WMI test variable and write log entry if matches False
            "$formattedDate - $($pc.ToUpper()): WMI failed to connect (Either is not Windows Or WMI is corrupted)" | Out-File -FilePath $errorLog -Append
        }
        else {
            # Gather all computer info

            # Run WMI Calls
			$sys = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $pc
            $netAdapter = Get-WmiObject -Class Win32_NetworkAdapter -ComputerName $pc | Where-Object {($_.PhysicalAdapter -eq $true)}
            $netIpAddress = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $PC | Where-Object {$_.IPEnabled -eq $True}
            $RAM = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $pc
            $cpu = Get-WmiObject -Class Win32_Processor -ComputerName $pc
            $bios = Get-WmiObject -Class Win32_Bios -computer $PC

            # Set data info variables
			$pcName = $sys.Name			
            $pcModel = $sys.Model
			$pcMfr = $sys.Manufacturer
            $pcSerial = $bios.Serialnumber
            if ($sys.model.ToUpper().contains('VIRTUAL')) {
                $pcType = 'Virtual Machine'
            }
            else {
                $pcType = 'Physical'
            }            
            $domain = $sys.Domain
            $osCapt = $osInfo.Caption
            $osVer = $osInfo.Version	
			$osArct = $osInfo.OSArchitecture
            $spMajVer = $osInfo.ServicePackMajorVersion.ToString()
            $spMinVer = $osInfo.ServicePackMinorVersion.ToString()
            $serPackVer = "SP $spMajVer.$spMinVer"
            $cpuType = (@($cpu.Name) -join "`r`n")
			$totalRam = [Math]::Round((($ram | Measure-Object -property 'Capacity' -Sum).Sum) / 1GB)
            $actIPAddr = (@($netIpAddress | Select-Object -ExpandProperty IPAddress | Where-Object {$_ -notlike "*:*"}) -join "`r`n")
            $actMacAddr = (@($netIpAddress.MACAddress) -join "`r`n")
            $allMacAddr = (@($netAdapter.MACAddress) -join "`r`n")

            # Create new custom object and add all porperties
			$sysInfo = New-Object -TypeName PSObject
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'PC Name' -Value $pcName
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'PC Model' -Value $pcModel
			$sysInfo | Add-Member -MemberType NoteProperty -Name 'PC Manufacturer' -Value $pcMfr
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'PC Type' -Value $pcType
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'PC Serial #' -Value $pcSerial
			$sysInfo | Add-Member -MemberType NoteProperty -Name 'Active Domain Or WorkGroup' -Value $domain 
			$sysInfo | Add-Member -MemberType NoteProperty -Name 'Operating System' -Value $osCapt
			$sysInfo | Add-Member -MemberType NoteProperty -Name 'OS Version' -Value $osVer
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'OS Architecture' -Value $osArct
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'Service Pack Version' -Value $serPackVer
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'CPU Type' -Value $cpuType
            $sysInfo | add-member -memberType NoteProperty -Name 'Total Memory (GB)' -value $totalRam         
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'Active IP Address' -Value $actIPAddr
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'Active Network Adapter MAC Address' -Value $actMacAddr
            $sysInfo | Add-Member -MemberType NoteProperty -Name 'ALL Network Adapter MAC Addresses' -Value $allMacAddr
           
            # Add custom object to results array
            [array]$results += $sysInfo         
		}		    
    }    
}

END {
    # Check if results array is null, if not exprt to CSV file
    if ($results -ne $null) {
        $results | Export-Csv -Path $csvReport -NoTypeInformation
    }
}