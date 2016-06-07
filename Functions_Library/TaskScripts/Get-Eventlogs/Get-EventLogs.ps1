$computers = (Get-Content "C:\Admin_Scripts\servers.txt")
$dateTime = Get-Date -format MM-dd-yyyy_HH-mm

$seperationText = " ********************************** "

function Check-IfNotPathCreate([string]$FolderPath) {
    if (!(Test-Path -Path $FolderPath)) {
        new-item -Path $FolderPath -ItemType directory | Out-Null
    }
} 

foreach ($computer in $computers) {
    try {
	
	$progressLog = "C:\Admin_Scripts\TaskProgressLog.txt"               
	$startText = "Start $computer $dateTime..."
	$endText = "...Finished $computer."

	$startText | Out-File $progressLog -Append -Force

        $logs = Get-WmiObject -EnableAllPrivileges -Class Win32_NTEventLogFile -ComputerName $computer -ErrorAction stop
        $logTemp = "\\$computer\c$\Logtemp"
        $folder = "$dateTime-$computer"
        $dataStore = "\\dahcbueventlog\e$\$computer"   

        foreach ($log in $logs) {

            $logTempFolder = "$logTemp\$folder"
            $dataStoreFolder = "$dataStore\$folder"
            Check-IfNotPathCreate($logTempFolder)
            Check-IfNotPathCreate($dataStoreFolder)

            $path = "{0}\{1}.evtx" -f $logTempFolder,$log.LogFileName

            $log.BackupEventLog($path) | Out-Null

            Copy-Item $path $dataStoreFolder -force         
        }

        $endText | Out-File $progressLog -Append -Force

        Remove-Item $logTemp -Recurse -Force
    }
    catch {
	$error | Out-File $progressLog -Append -Force
    }
}
$seperationText | Out-File $progressLog -Append -Force
