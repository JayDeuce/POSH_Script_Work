<#
.SYNOPSIS
    LogCentralizer.ps1 - Fetches IIS logfiles, Application, Security and System eventlogs (exported in text files) for the day before from remote computer(s)
.DESCRIPTION
    LogCentralizer.ps1 - Fetches IIS logfiles, Application, Security and System eventlogs (exported in text files) for the day before from remote computer(s)
    Has to be scheduled each day to be "like a inverted-syslog"
.PARAMETER Servers
    Defines the server list
    Default is "(Get-Content .\servers.txt)".
.NOTES
    File Name   : LogCentralizer.ps1
    Author      : Fabrice ZERROUKI - fabricezerrouki@hotmail.com
.EXAMPLE
    PS D:\> .\LogCentralizer.ps1 -Servers COMPUTER1, COMPUTER2
    From the remote computers COMPUTER1 and COMPUTER2; fetches yesterday's IIS logfiles, Application, Security and System eventlogs (each exported in a text file)
    and place them in the following directory structure (if today is 20/12/2012):
    COMPUTER1
            └───Events
                     └───20122012
                                └───ApplicationEvents.zip
                                └───SecurityEvents.zip
                                └───SystemEvents.zip
            └───IIS
                  └───20122012
                             └───ex121219.zip
    COMPUTER2
            └───Events
                     └───20122012
                                └───ApplicationEvents.zip
                                └───SecurityEvents.zip
                                └───SystemEvents.zip
            └───IIS
                  └───20122012
                             └───ex121219.zip
#>
Param(
    [Parameter(Mandatory=$true, HelpMessage="You must provide at least one server to get logs from. Could be a list of computers (comma separated or put the list in a .\servers.txt file)")]
    $Servers=(Get-Content .\servers.txt)
)
 
function New-Zip
{
    param([string]$zipfilename)
    set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipfilename).IsReadOnly = $false
}
  
function Add-Zip
{
    param([string]$zipfilename)
  
    if(-not (test-path($zipfilename)))
    {
        set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (dir $zipfilename).IsReadOnly = $false
    }
      
    $shellApplication = new-object -com shell.application
    $zipPackage = $shellApplication.NameSpace($zipfilename)
      
    foreach($file in $input)
    {
            $zipPackage.CopyHere($file.FullName)
            Start-sleep -milliseconds 500
    }
}
 
$i=0
# We assume all the IIS websites logs are located under the same folder for all the servers...
$IISLogsRootPath="D:\LOGS\HTTP"
$Today=Get-Date -Day $((Get-Date -Format "dd") - 1) -Hour 0 -Minute 0 -Second 0 -Format "ddMMyyyy"
$YesterdayIIS=Get-Date -Day $((Get-Date -Format "dd") - 1) -Hour 0 -Minute 0 -Second 0 -Format "yyMMdd"
$Start=([datetime]::Today).AddDays(-1)
$End=([datetime]::Today).AddDays(-2)
 
ForEach ($Server in $Servers) 
{
$i++
Write-Progress -Id 1 -Activity "Collecting yesterday's logs from $Server. Please wait..." -Status "Progress:" -PercentComplete ($i/($Servers.Count)*100)
if(!(Test-Path .\$Server\IIS\$Today)) {New-Item -ItemType Directory -Path .\$Server\IIS\$Today | Out-Null}

$LogFolders=Get-ChildItem -Path "D:\Logs\HTTP" | Where {($_.PSIsContainer)} | Select -ExpandProperty Name

ForEach ($LogFolder in $LogFolders) {
    $LogLocation=$IISLogsRootPath + "\" + $LogFolder
    $LogLocation=$LogLocation -replace ":","$"
    $LogFile=$LogLocation + "\ex" + $YesterdayIIS + ".log"
    $YesterdayLog="\\$Server\$LogFile"
    Write-Progress -Id 2 -Activity "Copying yesterday's IIS logfile ($LogFile) from $Server. Please wait..." -Status "Progress:" -PercentComplete (1/4*100)
    Copy-Item $YesterdayLog (".\$Server\IIS\$Today\" + "ex" + $YesterdayIIS + ".log")
    $ZipName="\ex" + $YesterdayIIS + ".zip"
    New-Zip $ZipName
    $FileName=".\$Server\IIS\$Today\" + "ex" + $YesterdayIIS + ".log"
    $FileName | Add-Zip $ZipName
    Remove-Item $FileName
}
 
if(!(Test-Path .\$Server\Events\$Today)) {New-Item -ItemType Directory -Path .\$Server\Events\$Today | Out-Null}
Write-Progress -Id 2 -Activity "Exporting yesterday's Application Eventlog from $Server. Please wait..." -Status "Progress:" -PercentComplete (1/4*100)
$ApplicationEvents=Get-EventLog -logName Application -ComputerName $Server -Before $Start -After $End | Format-Table -Wrap -Property TimeWritten, EntryType, Source, EventID, Message -Autosize
$ApplicationEvents | Out-File .\$Server\Events\$Today\ApplicationEvents.txt
$ZipName=".\$Server\Events\$Today\ApplicationEvents.zip"
New-Zip $ZipName
$FileName=".\$Server\Events\$Today\ApplicationEvents.txt"
$FileName | Add-Zip $ZipName
Remove-Item $FileName
 
if(!(Test-Path .\$Server\Events\$Today)) {New-Item -ItemType Directory -Path .\$Server\Events\$Today | Out-Null}
Write-Progress -Id 2 -Activity "Exporting yesterday's Security Eventlog from $Server. Please wait..." -Status "Progress:" -PercentComplete (1/4*100)
$SecurityEvents=Get-EventLog -logName Security -ComputerName $Server -Before $Start -After $End | Format-Table -Wrap -Property TimeWritten, EntryType, Source, EventID, Message -Autosize
$SecurityEvents | Out-File .\$Server\Events\$Today\SecurityEvents.txt
$ZipName=".\$Server\Events\$Today\SecurityEvents.zip"
New-Zip $ZipName
$FileName=".\$Server\Events\$Today\SecurityEvents.txt"
$FileName | Add-Zip $ZipName
Remove-Item $FileName
 
if(!(Test-Path .\$Server\Events\$Today)) {New-Item -ItemType Directory -Path .\$Server\Events\$Today | Out-Null}
Write-Progress -Id 2 -Activity "Exporting yesterday's System Eventlog from $Server. Please wait..." -Status "Progress:" -PercentComplete (1/4*100)
$SystemEvents=Get-EventLog -logName System -ComputerName $Server -Before $Start -After $End | Format-Table -Wrap -Property TimeWritten, EntryType, Source, EventID, Message -Autosize
$SystemEvents | Out-File .\$Server\Events\$Today\SystemEvents.txt
$ZipName=".\$Server\Events\$Today\SystemEvents.zip"
New-Zip $ZipName
$FileName=".\$Server\Events\$Today\SystemEvents.txt"
$FileName | Add-Zip $ZipName
Remove-Item $FileName
}