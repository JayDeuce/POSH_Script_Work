<#
.SYNOPSIS
     Starts the Transcript cmdlet and sets logging to a chosen Logfile

.DESCRIPTION
     The Script will check if a LogFile Path was passed into it, if not
     it will ask the location and run start-transcript with a logfile
     name using a Date-Time Stamp for the filename.

.PARAMETER  -LogPath
     The location to save the logfile.

.EXAMPLE
     Start-ConsoleTranscript

     Description:
          Script will run and ask for a path to save the logfile

.EXAMPLE
     Start-ConsoleTranscript -LogPath "C:\Log"

     Description:
          Script will run and start the Transcript saving the File
          to the C:\Log Directory.

.INPUTS
     System.String

.OUTPUTS
     Logfile Named As: %(Current)Date-Time%.txt

.NOTES
     Based on the Start-Trans Function Created by Ed Wilson, The Scripting Guy over at the
     Hey, Scripting Guy! Blog http://blogs.technet.com/b/heyscriptingguy/

     Link to Specific Post:
     http://blogs.technet.com/b/heyscriptingguy/archive/2010/09/25/create-a-transcript-of-commands-from-the-windows-powershell-ise.aspx
#>
[cmdletbinding()]
Param
(
     [String]$LogFilePath = ""
)

# Set the ComputerName+Date-Time $LogFileName variable for the Log File Name
[string]$ComputerName = get-content env:computername
[string]$Date = [dateTime]::now.toshortdatestring().Replace("/", "-")
[string]$Time = [dateTime]::now.toshorttimestring().Replace(":", "-").Replace(" ", "-")
[string]$DateTime = $Date + "_" + $Time
[string]$LogFileName = $ComputerName + "_" + $DateTime + ".txt"

function checkfolderPath([string]$FolderPath) {
     if (!(Test-Path -Path $FolderPath)) {
          new-item -Path $FolderPath -ItemType directory | Out-Null
     }
}

If ($LogFilePath -eq "") {
     # Ask Where to Save
     [string]$LogFilePath = [Microsoft.VisualBasic.Interaction]::InputBox("Where would you like to save the transcript file?", "Trascript Path")
     If ($LogFilePath -eq "") {
          # Error handling if answer is blank or Cancel is clicked
          Write-Host "`n"
          Write-Host "HEY!!! No transcript save file location set. Exiting!!" -BackgroundColor "darkgreen" -ForegroundColor "yellow"
          Write-Host "`n"
          [Microsoft.VisualBasic.Interaction]::MsgBox("No transcript save file location set. Exiting!", "OKOnly, Exclamation", "Transcript Status") | Out-Null
     }
     Else {
          $LogFilePath = $LogFilePath.TrimEnd("\")
          checkfolderPath($LogFilePath)
          Write-Host "`n"
          start-transcript -path "$LogFilePath\$LogFileName" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
          Write-Host "`n"
     }
}
Else {
     checkfolderPath($LogFilePath)
     Write-Host "`n"
     start-transcript -path "$LogFilePath\$LogFileName" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
     Write-Host "`n"
}