# ====================================================================================
# Pulled from autospsourcebuilder, Created by Brian Lalancette (@brianlala)
# http://autospsourcebuilder.codeplex.com
# Remove Read-Only Attribute
# ====================================================================================

Function Remove-ReadOnlyAttribute ($Path) {
     ForEach ($item in (Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue)) {
          $attributes = @((Get-ItemProperty -Path $item.FullName).Attributes)
          If ($attributes -match "ReadOnly") {
               # Set the file to just have the 'Archive' attribute
               Write-Host "  - Removing Read-Only attribute from file: $item"
               Set-ItemProperty -Path $item.FullName -Name Attributes -Value "Archive"
          }
     }
}

#------------------

# ===================================================================================
# Pulled from autospsourcebuilder, Created by Brian Lalancette (@brianlala)
# http://autospsourcebuilder.codeplex.com
# Func: Pause
# Desc: Wait for user to press a key - normally used after an error has occured or input is required
# ===================================================================================
Function Pause($action, $key) {
     # From http://www.microsoft.com/technet/scriptcenter/resources/pstips/jan08/pstip0118.mspx
     if ($key -eq "any" -or ([string]::IsNullOrEmpty($key))) {
          $actionString = "Press any key to $action..."
          Write-Host $actionString
          $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
     }
     else {
          $actionString = "Enter `"$key`" to $action"
          $continue = Read-Host -Prompt $actionString
          if ($continue -ne $key) {
               pause $action $key
          }

     }
}

#-----------------

# ====================================================================================
# First check if we are running this under an elevated session.
# Pulled from the script at http://gallery.technet.microsoft.com/scriptcenter/1b5df952-9e10-470f-ad7c-dc2bdc2ac946
# ====================================================================================

If (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
     Write-Warning " - You must run this script under an elevated PowerShell prompt. Launch an elevated PowerShell prompt by right-clicking the PowerShell shortcut and selecting `"Run as Administrator`"."
     break
}

#------------------

# ====================================================================================
# Pulled from autospsourcebuilder, Created by Brian Lalancette (@brianlala)
# http://autospsourcebuilder.codeplex.com
# Func: EnsureFolder
# Desc: Checks for the existence and validity of a given path, and attempts to create if it doesn't exist.
# From: Modified from patch 9833 at http://autospinstaller.codeplex.com/SourceControl/list/patches by user timiun
# ====================================================================================

Function EnsureFolder ($Path) {
     If (!(Test-Path -Path $Path -PathType Container)) {
          Write-Host -ForegroundColor White " - $Path doesn't exist; creating..."
          Try {
               New-Item -Path $Path -ItemType Directory | Out-Null
          }
          Catch {
               Write-Warning " - $($_.Exception.Message)"
               Throw " - Could not create folder $Path!"
               $errorWarning = $true
          }
     }
}

#-----------------

# ====================================================================================
# Check is a folder path is there, if not create it.
# ====================================================================================
function checkfolderPath([string]$FolderPath) {
     if (!(Test-Path -Path $FolderPath)) {
          new-item -Path $FolderPath -ItemType directory | Out-Null
     }
}

#-----------------

# ====================================================================================
# Build a FileName out of Now Date and Time. Vatiable Build
# ====================================================================================

[string]$Date = [dateTime]::now.toshortdatestring().Replace("/", "-")
[string]$Time = [dateTime]::now.toshorttimestring().Replace(":", "-").Replace(" ", "-")
[string]$DateTime = $Date + "_" + $Time

# OR

Get-Date -Format "dd-mm-2014_HH-mm"

#-----------------

# ====================================================================================
# Start ConsoleTrancript Function - Used in Profile Scripts
# ====================================================================================

function Start-ConsoleTranscript() {
     # Check if I want to start a Transcript, Used if I am trying out commands
     [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
     [string]$TranYN = [Microsoft.VisualBasic.Interaction]::MsgBox("Would you Like to Start a Transcript?", "YesNo, DefaultButton2", "Start Transcript?")
     # Check the answer for "Yes" or "No" and act upon the choice
     switch ($TranYN) {
          Yes {
               [string]$Date = [dateTime]::now.toshortdatestring().Replace("/", "-")
               [string]$Time = [dateTime]::now.toshorttimestring().Replace(":", "-").Replace(" ", "-")
               [string]$DateTime = $Date + "_" + $Time

               # Check for my Standard Scripts Folder, and act accordingly to its presence or not
               If (Test-Path -Path "scripts:\") {
                    start-transcript -path "Scripts:\Transcripts\$DateTime.txt" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
               }
               Else {
                    # If Standard Scripts Path is not available, ask where to save
                    [string]$LogPath = [Microsoft.VisualBasic.Interaction]::InputBox("Where would you like to save the transcript file?", "Trascript Path")
                    If ($LogPath -eq "") {
                         # Error handling is answer is blank or I hit cancel
                         Write-Warning "HEY!!!, No transcript save file location set, defaulting to Transcript Off!!"
                         Write-host "`n"
                         [Microsoft.VisualBasic.Interaction]::MsgBox("Transcript is off!", "OKOnly, Exclamation", "Transcript Status") | Out-Null
                    }
                    else {
                         $LogPath = $LogPath.TrimEnd("\")
                         checkfolderPath($LogPath)
                         start-transcript -path "$LogPath\$DateTime.txt" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
                         Write-host "`n"
                    }
               }
          }
          No {
               Write-host "Transcripting is turned off for this console session." -BackgroundColor "darkgreen" -ForegroundColor "yellow"
          }
     }
}

#-----------------

# ====================================================================================
#  Email Variable Creation and Code to Send
# ====================================================================================

$Recipients = '"User1 <user1@web.com>"', '"User1 <user1@web.com>"', '"User1 <user1@web.com>"'
$Sender = '"Sender1 <Sender1@web.com>"'
$SMTPServer = "000.000.000.000 <ServerAddress>"
$Subject = "Subject"
$Body = @"
Good Morning,

BODY OF THE EMAIL

"@

#Build the Notification Email and send it to the Recipients
Send-MailMessage -To $Recipients -From $Sender -Subject $Subject -Body $Body -SmtpServer $SMTPServer

#-----------------

# ====================================================================================
#  Find-Files
#  Func: One Liner
#  Desc: Find File in File Path
#  From: https://gallery.technet.microsoft.com/scriptcenter/Search-for-Files-Using-340397aa
# ====================================================================================

Get-ChildItem -Recurse -Force $filePath -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -like "*$fileName*") } | Select-Object Name, Directory| Format-Table -AutoSize *

#-----------------

# ====================================================================================
#  Get-Parameters
#  Func: One Liner
#  Desc: List all parameters of a Script or cmdlet in a neat little format
# ====================================================================================

# List Just Parameter Names
(Get-Help %SCRIPT/CMDLETNAME%).Syntax | SELECT-OBJECT –ExpandProperty SyntaxItem | SELECT-OBJECT –ExpandProperty parameter | SELECT-OBJECT name

# List Parameters and their proerties (Description, Mandatory, etc)
(Get-Help %SCRIPT/CMDLETNAME%).Syntax | SELECT-OBJECT –ExpandProperty SyntaxItem | SELECT-OBJECT –ExpandProperty parameter

#-----------------

# ====================================================================================
#  Replace-Character
#  Func: oneliner with scriptblock
#  Desc: Get all objects (Folders/Files) in a folder and replace any Character
#        found with another character
# ====================================================================================

Get-ChildItem -recurse | Where-Object {$_.name -match " "} | ForEach-Object {
     $New = $_.name.Replace(" ", "_")
     Rename-Item -path $_.Fullname -newname $New -passthru
}

#-----------------

# ====================================================================================
#  Copy-HistoryToISE
#  Func: ISE Function to add command to copy history to a new ISE tab
#  Desc: Get all commands in the history of the powershell ISE and anything that
#        completed succesfully copy to a new ISE tab
# ====================================================================================

function Copy-HistoryToISE {
     $file = $psise.CurrentPowerShellTab.Files.Add()
     $file.Editor.Text = (Get-History | Where-Object ExecutionStatus -eq Completed).CommandLine
     $file.Editor.SetCaretPosition(1, 1)
}

$psise.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('Copy History', { Copy-HistoryToISE }, 'CTRL+ALT+H')

#-----------------

# ====================================================================================
#  Relaunch-AsAdmin
#  Func: Relaunch current script as admin
#  Desc: Checks to see if the current script is running as admin or not and relaunchs
#        If it is not
# ====================================================================================

Function Relaunch-AsAdmin {
     # Get the ID and security principal of the current user account
     $myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
     $myWindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
     # Get the security principal for the Administrator role
     $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
     # Check to see if we are currently running "as Administrator"
     if (!($myWindowsPrincipal.IsInRole($adminRole))) {
          Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList "-NonInteractive", "-WindowStyle Hidden", "-File $PSCommandPath"\
          exit
     }
}

# ====================================================================================
#  Check-RunningProcessesAndPercentages
#  Func: Check the listed machines running processes and get a report
#  Desc: Check the listed machines running processes and get a report showing them an
#  	    the percentages of CPU and Memory Usage
# ====================================================================================

Get-WmiObject Win32_PerfFormattedData_PerfProc_Process -ComputerName localhost |
     Where-Object {
     $_.name -inotmatch '_total|idle'
} |
     ForEach-Object {
     "Process={0,-25} CPU_Usage={1,-12} Memory_Usage_(MB)={2,-16}" -f $_.Name, $_.PercentProcessorTime, ([math]::Round($_.WorkingSetPrivate / 1Mb, 2))
}

# ====================================================================================
#  Check-ForStaticIP
#  Func: Check the listed machines and see if they are Set to Staic IP
#  Desc: Check the listed machines and see if they are Set to Staic IP and the Output
#        to CSV Report
# ====================================================================================

foreach ($ip in $iplist) {
     Write-Host "Processing: $ip`n"
     Get-WMIObject Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue -computername $ip |
          Where-Object {
          $_.IPEnabled -eq $true -and $_.DHCPEnabled -eq $false
     } |
          Select-Object @{
          Name = 'IpAddress'; Expression = {$_.IpAddress -join '; '}
     }, DNSHostname |
          Export-Csv "list.csv" -Append -NoTypeInformation
}