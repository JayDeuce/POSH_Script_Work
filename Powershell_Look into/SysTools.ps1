<# 
    .SYNOPSIS 
        This is a simple Powershell script to provide a set of 15 tools to make our life easy. 
 
    .DESCRIPTION 
        You can use the below key words to get the corresponding tools 
         R - To restart any remote server 
         I - To do IISRESET on a remote web server 
         L - To log off a specific Terminal Service session to a remote machine (You will be prompted for Server Name) 
         S - To start a specific service on a remote machine (Make sure to provide the 'Service Name'properly- as shown in the properties page of the services) 
         X - To stop a specific service on a remote machine (Make sure to provide the 'Service Name'properly- as shown in the properties page of the services) 
         T - To get the status of scheduled tasks running on a remote machine (You will be prompted for Server Name) 
         J - To get the status of SQL Job starting with some string we prefer (You will be prompted to give Server name and Job name string) 
         B - To get Basic computer Inventory 
         P - To list the process running on a remote server 
         H - To get the whois lookup for any public IP address 
         F - To find and replace a string in any file on a computer  
         M - To monitor the health of computers  
         E - To get event log details of a server 
         D - To download huge files using BITS Transfer Jobs  
         Q - To Exit 
 
    .NOTES 
        You need an administrative account with domain privileges and SQL access to use this tool. 
        Make sure to exit this application / Lock your computer, when you are about to leave for the day- to avoid any possible security problems. 
        Much coding is not done to handle all the errors and exceptions, so you might receive errors in non standard scenarios. 
        This program uses terminal service powershell modules which need to be installed separately on the machine on which this is planning to be used. 
 
    .LINK 
        http://technet.microsoft.com/hi-in/scriptcenter/powershell(en-us).aspx 
        http://www.powershellcommunity.org/ 
        http://en.wikipedia.org/wiki/Windows_PowerShell 
#> 
 
 
######################################## DISCLAIMER ################################### 
# The free software programs provided by Shaiju J.S may be freely distributed,     ####  
# provided that no charge above the cost of distribution is levied, and that the ###### 
# disclaimer below is always attached to it.                                     ###### 
# The programs are provided as is without any guarantees or warranty.            ######  
# Although the author has attempted to find and correct any bugs in the free software # 
# programs, the author is not responsible for any damage or losses of any kind caused # 
# by the use or misuse of the programs. The author is under no obligation to provide ## 
# support, service, corrections, or upgrades to the free software programs. ###########  
####################################################################################### 
 
 
write-host "###########################################################################" -ForegroundColor DarkCyan 
write-host "###################### TOOL KIT USING POWERSHELL ##########################" -ForegroundColor Yellow 
write-host "################### 15 TOOLS TO AVOID RDP SESSIONS ########################" -ForegroundColor Green 
write-host "####################### USE ADMIN ACCOUNT TO START ########################" -ForegroundColor Yellow 
write-host "## RM Tool Version 3.0.0 #### By Shaiju J.S ########## 10 Oct 2015 ########" -ForegroundColor Green 
write-host "###########################################################################" -ForegroundColor DarkCyan 
 
####################################################################################### 
$a = (Get-Host).PrivateData 
$a.WarningBackgroundColor = "red" 
$a.WarningForegroundColor = "white" 
#To change your screen or background color set the following: 
#$Host.Ui.RawUi.BackGroundColor = "Blue" 
# To change your test or foreground color set the following: 
#$Host.Ui.RawUi.ForeGroundColor = "Yellow" 
 
gc env:computername  
Get-Date  
 
function Read-Choice { 
    PARAM([string]$message, [string[]]$choices, [int]$defaultChoice=15, [string]$Title=$null ) 
      $Host.UI.PromptForChoice( $caption, $message, [Management.Automation.Host.ChoiceDescription[]]$choices, $defaultChoice ) 
} 
 
switch(Read-Choice "Use Shortcut Keys:[]" "&Restart","&IISRESET","&Log-off TS","&Start Service","&X-Stop Service","&Task- Scheduled","&Job- SQL","&Basic Computer Inventory","&Application List","&Process List","&Event Logs","&H-Whois Lookup","&FindnReplace","&Monitoring Health","&Download Huge Files","&Quit"){ 
    0 {  
        Write-Host "You have selected the option to restart a server" -ForegroundColor Yellow 
        $ServerName = Read-Host "Enter the name of the server to be restarted" 
        if (Test-connection $ServerName) { 
            Get-Date 
            write-host "$ServerName is reachable" 
            Write-Host "$ServerName is getting restarted" 
            Get-Date  
            restart-computer -computername $ServerName -Force  
            Write-Host "Starting continuous ping to test the status"   
            Test-Connection -ComputerName $ServerName -Count 100 | select StatusCode          
            Start-Sleep -s 300 
            Write-Host "Here is the last reboot time: "  
            $wmi=Get-WmiObject -class Win32_OperatingSystem -computer $ServerName  
            $LBTime=$wmi.ConvertToDateTime($wmi.Lastbootuptime) 
            $LBTime 
             
        } 
        else { 
                Get-Date 
                write-host "$ServerName is not reachable, please check this manually" 
                exit 
        } 
 
    }  
    1 {  
        Write-Host "You have selected the option to do IISRESET" -ForegroundColor Yellow 
        $Server1 = Read-Host "Enter the server name on which iis need to be reset" 
        Invoke-Command -ComputerName $Server1 -ScriptBlock {iisreset} 
    }  
    2 { 
        Write-Host "You have selected the option to list and log off terminal service sessions" -ForegroundColor Yellow 
        Import-Module PSTerminalServices 
        $server9 = Read-Host "Enter Remote Server Name" 
        $session = Get-TSSession -ComputerName $server9 | SELECT "SessionID","State","IPAddress","ClientName","WindowStationName","UserName" | Out-GridView     
        $session 
        $s = Read-Host "Enter Session ID, if you want to log off any session" 
        Get-TSSession -ComputerName $server9 -filter {$_.SessionID -eq $s} | Stop-TSSession -Force 
    } 
    3 { 
        Write-Host "You have selected the option to start a service" -ForegroundColor Yellow 
        $Server6 = Read-host "Enter the remote computer name" 
        Get-Service * -computername $Server6 | where {$_.Status -eq "Stopped"} | Out-GridView   
        $svc6 = Read-host "Enter the name of the service to be started" 
        $Status1 = (Get-WmiObject -computer $Server6 Win32_Service -Filter "Name='$svc6'").InvokeMethod("StartService",$null) 
        If ($Status1 -eq '0') { 
        Write-Host "The $svc6 has started successfully" 
        } 
        else { 
        Write-Host "Unable to start $svc6, please troubleshoot" 
        } 
    } 
    4 { 
        Write-Host "You have selected the option to stop a service" -ForegroundColor Yellow 
        $Server7 = Read-host "Enter the remote computer name" 
        Get-Service * -computername $Server7 | where {$_.Status -eq "Running"} | Out-GridView   
        $svc7 = Read-host "Enter the name of the service to be stopped" 
        $Status2 = (Get-WmiObject -computer $Server7 Win32_Service -Filter "Name='$svc7'").InvokeMethod("StopService",$null) 
        If ($Status2 -eq '0') { 
        Write-Host "The $svc7 has stopped successfully" 
        } 
        else { 
        Write-Host "Unable to stop $svc7, please troubleshoot" 
        } 
    } 
    5 { 
        Write-Host "You have selected the option to get the scheduled task status list" -ForegroundColor Yellow 
        $Server8 = Read-host "Enter the remote computer name" 
        $t = Read-Host "Task names starting with.....or press Enter Key to get the full list" 
        if ($t -ne $null) { 
        $Result1 = schtasks /query /S $Server8 | ?{$_ -like "$t*"} 
        $Result1 
        Write-Host "Please wait for the result in grid view." 
        $Result1 | Out-GridView 
        } 
        else { 
        Write-Host "Please wait for the result in grid view." 
        schtasks /query /S $Server8 /FO TABLE /V | Out-GridView 
        } 
    } 
    6 {  
        Write-Host "You have selected the option to get the status of SQL job" -ForegroundColor Yellow 
        write-host "Hope you are logged in with an account having SQL access privilege" 
        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null 
        $instance = Read-Host "Enter the server name" 
        $j = Read-Host "Job names starting with.....or press Enter Key to get the full list" 
        If ($j -ne $null) { 
        $s = New-Object ('Microsoft.SqlServer.Management.Smo.Server') $instance 
        $Result2 = $s.JobServer.Jobs |Where-Object {$_.Name -ilike "$j*"}| SELECT NAME, LASTRUNOUTCOME, LASTRUNDATE  
        $Result2 
        Write-Host "Please wait for the result in grid view." 
        $Result2 | Out-GridView 
        } 
        else { 
        Write-Host "Please wait for the result in grid view." 
        $s.JobServer.Jobs | SELECT NAME, LASTRUNOUTCOME, LASTRUNDATE | Out-GridView 
        } 
    }  
    7 { 
        Write-Host "You have selected the option to get basic computer inventory" -ForegroundColor Yellow 
        $server10 = Read-Host "Enter Remote Server Name" 
        $Inventory = Read-Host "Enter a drive name to store the log- Eg: D:" 
        Clear-Content -Path $Inventory'\Inventory.csv' 
        Start-Transcript -Path $Inventory'\Inventory.csv' 
        $OS = (Get-WmiObject -class Win32_OperatingSystem -Computer $Server10).caption 
        $CPU = Get-WmiObject -Class Win32_processor -Computer $server10 | foreach {$_.Name} 
        $MothBoard = Get-WmiObject -Class Win32_baseboard -Computer $server10 | foreach {$_.SerialNumber} 
        $Domain = Get-WmiObject -Class Win32_Computersystem -Computer $server10 | foreach {$_.Domain} 
        $Vendor = Get-WmiObject -Class Win32_Computersystem -Computer $server10 | foreach {$_.Manufacturer} 
        $Model = Get-WmiObject -Class Win32_Computersystem -Computer $server10 | foreach {$_.Model} 
        $Memory = Get-WmiObject -Class Win32_Computersystem -Computer $server10 | foreach {[math]::truncate($_.TotalPhysicalMemory / 1MB)} 
        $ServiceTag = gwmi win32_bios -ComputerName $server10 | foreach {$_.SerialNumber} 
        $ip = gwmi win32_networkadapterconfiguration -filter "IPEnabled=True" -Computer $server10 | foreach {$_.IPAddress} 
        $mac = gwmi win32_networkadapterconfiguration -filter "IPEnabled=True" -Computer $server10 | foreach {$_.description,$_.macaddress}  
        $gw = gwmi win32_networkadapterconfiguration -filter "IPEnabled=True" -Computer $server10 | foreach {$_.DefaultIPGateway}  
                                 
            function Get-DellWarranty{  
            [CmdletBinding()]  
            [OutputType([System.Object])]  
            Param(  
                # Name should be a valid computer name or IP address.  
                [Parameter(Mandatory=$False,   
                       ValueFromPipeline=$true,  
                       ValueFromPipelineByPropertyName=$true,   
                       ValueFromRemainingArguments=$false)]  
          
                [Alias('HostName', 'Identity', 'DNSHostName', 'ComputerName')]  
                [string[]]$Name,  
          
                 # ServiceTag should be a valid Dell Service tag. Enter one or more values.  
                 [Parameter(Mandatory=$false,   
                        ValueFromPipeline=$false)]  
                 [string[]]$ServiceTag  
                 )  
 
            Begin{  
                 }  
            Process{  
                if($ServiceTag -eq $Null){  
                    foreach($C in $Name){  
                        $test = Test-Connection -ComputerName $c -Count 1 -Quiet  
                            if($test -eq $true){  
                                $service = New-WebServiceProxy -Uri http://xserv.dell.com/services/assetservice.asmx?WSDL  
                                $system = Get-WmiObject -ComputerName $C win32_bios -ErrorAction SilentlyContinue  
                                $serial =  $system.serialnumber  
                                $guid = [guid]::NewGuid()  
                                $info = $service.GetAssetInformation($guid,'check_warranty.ps1',$serial)  
                          
                                $Result=@{  
                                'ComputerName'=$c  
                                'ServiceLevel'=$info[0].Entitlements[0].ServiceLevelDescription.ToString()  
                                'EndDate'=$info[0].Entitlements[0].EndDate.ToShortDateString()  
                                'StartDate'=$info[0].Entitlements[0].StartDate.ToShortDateString()  
                                'DaysLeft'=$info[0].Entitlements[0].DaysLeft  
                                'ServiceTag'=$info[0].AssetHeaderData.ServiceTag  
                                'Type'=$info[0].AssetHeaderData.SystemType  
                                'Model'=$info[0].AssetHeaderData.SystemModel  
                                'ShipDate'=$info[0].AssetHeaderData.SystemShipDate.ToShortDateString()  
                                }  
                      
                                $obj = New-Object -TypeName psobject -Property $result  
                                Write-Output $obj  
                     
                                $Result=$Null  
                                $system=$Null  
                                $serial = $null  
                                $guid=$Null  
                                $service=$Null  
                                $info=$Null  
                                $test=$Null   
                                $c=$Null  
                            }   
                            else{  
                                Write-Warning "$c is offline"  
                                $c=$Null  
                                }          
 
                        }  
                }  
                elseif($ServiceTag -ne $Null){  
                    foreach($s in $ServiceTag){  
                                $service = New-WebServiceProxy -Uri http://xserv.dell.com/services/assetservice.asmx?WSDL  
                                $guid = [guid]::NewGuid()  
                                $info = $service.GetAssetInformation($guid,'check_warranty.ps1',$S)  
                          
                                if($info -ne $Null){  
                          
                                    $Result=@{  
                                    'ServiceLevel'=$info[0].Entitlements[0].ServiceLevelDescription.ToString()  
                                    'EndDate'=$info[0].Entitlements[0].EndDate.ToShortDateString()  
                                    'StartDate'=$info[0].Entitlements[0].StartDate.ToShortDateString()  
                                    'DaysLeft'=$info[0].Entitlements[0].DaysLeft  
                                    'ServiceTag'=$info[0].AssetHeaderData.ServiceTag  
                                    'Type'=$info[0].AssetHeaderData.SystemType  
                                    'Model'=$info[0].AssetHeaderData.SystemModel  
                                    'ShipDate'=$info[0].AssetHeaderData.SystemShipDate.ToShortDateString()  
                                    }  
                                }  
                                else{  
                                Write-Warning "$S is not a valid Dell Service Tag."  
                                }  
                      
                                $obj = New-Object -TypeName psobject -Property $result  
                                Write-Output $obj  
                     
                                $Result=$Null  
                                $system=$Null  
                                $serial=$Null  
                                $guid=$Null  
                                $service=$Null  
                                $s=$Null  
                                $info=$Null  
                          
                               }  
                    }  
            }  
            End  
            {  
            }  
        } 
        Write-Host "`nBasic Information as below" -ForegroundColor Yellow 
        "PC Name               : $server10"  
        "Manufacturer          : $Vendor"     
        "Model                 : $Model"      
        "Service Tag           : $ServiceTag"  
        "Main Board            : $MothBoard"   
        "Processor             : $CPU"         
        "Physical Memory in MB : $Memory"      
        "Operating System      : $OS"          
        "Domain Name           : $Domain"      
        "IP address               : $ip"          
        "Mac address           : $mac"         
        "Gateway address       : $gw"  
        Write-Host "`nWarranty Details as below" -ForegroundColor Yellow 
        Get-DellWarranty -ServiceTag $ServiceTag | select ServiceLevel, EndDate, StartDate, DaysLeft, ServiceTag, Type, Model, ShipDate  
        Write-Host "`nDisk Space Information as below" -ForegroundColor Yellow 
        Get-WmiObject win32_logicaldisk -ComputerName $server10 | ForEach-Object {  
        Write-host "`n`nDrive ID         : " $_.DeviceID  
        Write-host "Drive Size       : " ([math]::truncate($_.size / 1GB)) "GB"  
        Write-host "Drive Free Space : " ([math]::truncate($_.FreeSpace / 1GB)) "GB"  
        } 
    Stop-Transcript 
    Get-Content -Path $Inventory'\Inventory.csv' | Out-GridView  
    } 
    8 { 
        Write-Host "The option to List the Applications installed on a remote machine" -ForegroundColor Yellow 
        # This script will Query the Uninstall Key on a computer specified in $computername and list the applications installed there  
        # $Branch contains the branch of the registry being accessed  
        $computername= Read-Host "Enter the computer name" 
        # Branch of the Registry  
        $Branch='LocalMachine'  
        # Main Sub Branch you need to open  
        $SubBranch="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"  
        $registry=[microsoft.win32.registrykey]::OpenRemoteBaseKey('Localmachine',$computername)  
        $registrykey=$registry.OpenSubKey($Subbranch)  
        $SubKeys=$registrykey.GetSubKeyNames()  
        $date = Get-date -uformat "%d%m%y" 
        $applog = Read-host "Specify the drive location -Eg: D:" 
        if (Test-Path -Path $applog'\applist.csv') { 
        Clear-Content -Path $applog'\applist.csv'     
        } 
        else { 
        write-host "Logging details to $applog" 
        } 
        # Drill through each key from the list and pull out the value of  
        # “DisplayName” – Write to the Host console the name of the computer  
        # with the application beside it 
        Foreach ($key in $subkeys) {  
            $exactkey=$key  
            $NewSubKey=$SubBranch+"\\"+$exactkey  
            $ReadUninstall=$registry.OpenSubKey($NewSubKey)  
            $Value=$ReadUninstall.GetValue("DisplayName")  
            $Value | Out-File -FilePath $applog'\applist.csv' -Append 
        }  
        Start-Sleep -Seconds 30 
        Write-Host "Please wait for the result in grid view." 
        Get-content -Path $applog'\applist.csv' | Out-GridView 
    } 
    9 { 
    Write-Host "You have selected the option to get process details of a remote server" -ForegroundColor Yellow 
    $server12 = Read-Host "Enter the remote machine name" 
    Write-Host "Please wait for the result in grid view." 
    Get-Process -ComputerName $server12 | Out-GridView   
    } 
    10 { 
    Write-Host "You have selected the option to get the event log details of a server" -ForegroundColor Yellow 
    $opt3 = Read-Host "Do you want to export details to excel (Y/N)?" 
    $server14 = Read-Host "Enter server name" 
    [int]$n = Read-Host "Last how many Hours?" 
    $event = Read-host "Application / Security / System ?" 
    $start1 = (Get-Date).addHours(-[int]$n)    
    $start2 = (Get-Date) 
    $strdat = (get-date).ToString() 
        if ($opt3 -eq 'Y') { 
            If ($event -eq 'Security') { 
            $entry2 = Read-Host "FailureAudit / SuccessAudit ?" 
            $location1 = Read-Host "Enter a drive location for the report" 
            get-eventlog -logname $event -EntryType $entry2 -after $start1 -before $start2 -ComputerName $server14 | Export-csv -Force -Path "$location1\$(Get-Date -Format 'dd_MM_yyyy')-$Event Log-$entry2-$server14.csv" 
            Invoke-Item "$location1\$(Get-Date -Format 'dd_MM_yyyy')-$Event Log-$entry2-$server14.csv" 
            } 
            else { 
            $entry0 = Read-Host "Information / Warning / Error ?" 
            $location2 = Read-Host "Enter a drive location for the report" 
            get-eventlog -logname $event -EntryType $entry0 -after $start1 -before $start2 -ComputerName $server14 | Export-csv -Force -Path "$location2\$(Get-Date -Format 'dd_MM_yyyy')-$Event Log-$entry0-$server14.csv" 
            Invoke-Item "$location2\$(Get-Date -Format 'dd_MM_yyyy')-$Event Log-$entry0-$server14.csv" 
            } 
        } 
        else { 
            If ($event -eq 'Security') { 
            $entry3 = Read-Host "FailureAudit / SuccessAudit ?" 
            Write-Host "Please wait for the result in grid view." 
            get-eventlog -logname $event -EntryType $entry3 -after $start1 -before $start2 -ComputerName $server14 | Out-GridView 
            } 
            else { 
            $entry1 = Read-Host "Information / Warning / Error ?" 
            Write-Host "Please wait for the result in grid view." 
            get-eventlog -logname $event -EntryType $entry1 -after $start1 -before $start2 -ComputerName $server14 | Out-GridView 
            } 
        } 
    } 
    11 { 
        Write-Host "You have selected the option to get the whois lookup details for a public IP" -ForegroundColor Yellow 
        Write-Host "You need internet access to proceed further, wait for the browser pop-up" -ForegroundColor Red 
        $fp = Read-Host "Enter the IP address for whois lookup" 
        Start-Process -FilePath "http://who.is/whois-ip/ip-address/$fp" 
    } 
    12 {     
        Write-Host "You have selected the option to find a specified string on files in a directory and replace it with another" -ForegroundColor Yellow 
        $Path = Read-host "Please specify the search location Eg: D:\Posh"  
        $Find = Read-Host "Find..." 
        Get-Childitem -Path "$Path" -Recurse -include "*.bat","*.config","*.cfg","*.txt","*.ps1","*.cmd","*.vbs" -ErrorAction SilentlyContinue | Select-String -Pattern "$Find" | Out-GridView  
        $opt1 = Read-Host "Do you want to replace the string (Y/N)?" 
        if ($opt1 -eq 'Y') { 
            $Replace = Read-host "Replace with..." 
            Get-Childitem -Path "$Path" -Recurse -include "*.bat","*.config","*.cfg","*.txt","*.ps1","*.cmd","*.vbs" | foreach-object { $a = $_.fullname; ( get-content $a ) | foreach-object { $_ -replace "$Find","$Replace" } | set-content $a } 
            Get-Childitem -Path "$Path" -Recurse -include "*.bat","*.config","*.cfg","*.txt","*.ps1","*.cmd","*.vbs" | Select-String -Pattern "$Replace" | Out-GridView 
        } 
        else { 
        $opt2 = Read-Host "Do you want to log the details for detailed verification (Y/N)?" 
        if ($opt2 -eq 'Y') { 
            $log = Read-host "Please specify the log location" 
            $date = Get-date -uformat "%d%m%y" 
            Get-Childitem -Path "$Path" -Recurse -include "*.bat","*.config","*.cfg","*.txt","*.ps1","*.cmd","*.vbs" | Select-String -Pattern "$Find" | Format-Table -Wrap -AutoSize | Out-String -Width 4096 | Out-file "$log\log$date.txt"  
        } 
        else { 
        Write-Host "Thank you for using this tool" 
        exit 
        } 
        } 
    } 
    13 { 
    Write-Host "You have selected the option to get the basic health status of a computer" -ForegroundColor Yellow 
    $mserver = Read-Host "Server name for monitoring?" 
    Get-WmiObject -Class Win32_ComputerSystem -ComputerName $mserver | Select NumberOfProcessors 
    Get-WmiObject win32_processor -Computer $mserver | select LoadPercentage 
    Get-WmiObject win32_OperatingSystem -Computer $mserver |%{"Total Physical Memory: {0}KB`nFree Physical Memory : {1}KB`nTotal Virtual Memory : {2}KB`nFree Virtual Memory  : {3}KB" -f $_.totalvisiblememorysize, $_.freephysicalmemory, $_.totalvirtualmemorysize, $_.freevirtualmemory} 
    Get-WmiObject -Class Win32_logicaldisk -Computer $mserver | Select DeviceID, FreeSpace, Size 
     
    } 
    14 { 
    Write-Host "You have selected the option to download huge files without file corruption using BITS Transfer" -ForegroundColor Yellow 
    Import-Module Bitstransfer 
    $src = Read-Host "Enter source path Eg: \\remotepc\e$\filename.iso" 
    $dst = Read-Host "Enter the destination path Eg: \\mypcname\d$"  
    Write-Host "Please check the progress bar and wait until the transfer completes, " -ForegroundColor Green 
    $start_time = Get-Date 
    Start-BitsTransfer -Source $src -Destination $dst -Priority Foreground    
    Write-Output "This download has taken $((Get-Date).Subtract($start_time).Seconds) second(s) to complete"  
    } 
    15 { 
        Write-Host "You have selected the option to exit the tool, Thank you for using this !!!" -ForegroundColor Yellow     
        exit 
    } 
     
} 