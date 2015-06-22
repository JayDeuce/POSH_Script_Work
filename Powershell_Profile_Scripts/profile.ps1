#region Set Windows Size and Shape in Console

if ($host.Name -eq "ConsoleHost") {
    # Set Variable
    $pshost = get-host
    $pswindow = $pshost.ui.rawui
    # Set Screen Buffer Size
    $newsize = $pswindow.buffersize
    $newsize.height = 3000
    $newsize.width = 150
    $pswindow.buffersize = $newsize
    # Set Windows Size
    $newsize = $pswindow.windowsize
    $newsize.height = 50
    $newsize.width = 150
    $pswindow.windowsize = $newsize
}

#endregion

#=======================================================================

#region Set Colors of "Write-" Messages

$colors = $host.PrivateData
$colors.VerboseForegroundColor = "white"
$colors.VerboseBackgroundColor = "blue"
$colors.WarningForegroundColor = "yellow"
$colors.WarningBackgroundColor = "darkgreen"
# Comment out the next two lines to set the default Fore=Red, Back=black
$colors.ErrorForegroundColor = "white"
$colors.ErrorBackgroundColor = "red"

#endregion

#=======================================================================

#region Set PSModulePath and PSDrives

# Checks paths to my Coding folders, and sets "Coding" PSDrive and $env:PSModulesPath Variable to them if found
# Allows loading of modules, not in the default directory. Defaults to "UserProfile" and no $env:PSModulesPath change.
If (Test-Path -Path "Z:\Wravien\Dropbox\Jon\Coding") {
    $env:PSModulePath = $env:PSModulePath + ";Z:\Wravien\Dropbox\Jon\Coding\SW_Config_Files\PowerShell\PSModules"
    new-psdrive -name Coding -psprovider FileSystem -root "Z:\Wravien\Dropbox\Jon\Coding" | Out-Null
    Set-Location Coding:
}
ElseIf (Test-Path -Path "F:\dropbox\jon\coding") {
    $env:PSModulePath = $env:PSModulePath + ";F:\Dropbox\Jon\Coding\SW_Config_Files\PowerShell\PSModules"
    new-psdrive -name Coding -psprovider FileSystem -root "F:\dropbox\jon\coding" | Out-Null
    Set-Location Coding:
}
Else {
    set-location $env:USERPROFILE
}

# In Dev: Command to find the Coding Folder Path - ISSUES: Very slow to search, fails using a VM on a MAC while sharing Mac Folders.
# $CodingPath = Get-ChildItem -path (Get-PSDrive -PSProvider FileSystem | Where-Object { $_.name -ne "Coding"}).root -Directory "*psmodules*" -Recurse -ErrorAction SilentlyContinue

#endregion

#=======================================================================

#region Import Modules

#region User Module Directory

# Load PSReadLine module from user modules directory if available
Import-Module PSReadLine

#endregion
#-----------------------------------------------------------------------
#region System Module Directory

# Load ActiveDirectory module from system modules directory if available
if(get-module -ListAvailable -Name "ActiveDirectory") {
    Import-Module ActiveDirectory
}

#endregion

#endregion

#=======================================================================

#region Helper Functions

#-----------------------------------------------------------------------
# Check-IfNotPathCreate Function
function Check-IfNotPathCreate([string]$FolderPath) {
    if (!(Test-Path -Path $FolderPath)) {
        new-item -Path $FolderPath -ItemType directory | Out-Null
    }
}
# End Check-IfNotPathCreate Function
#-----------------------------------------------------------------------
# Start-ConsoleTranscript Function
function Start-ConsoleTranscript() {

    # Check if I want to start a Transcript, Used if I am trying out commands
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    [string]$tranYN = [Microsoft.VisualBasic.Interaction]::MsgBox("Would you Like to Start a Transcript?", "YesNo, DefaultButton2",  "Start Transcript?")
    # Check the answer for "Yes" or "No" and act upon the choice
    switch ($tranYN) {
        Yes {
            [string]$computerName = $env:COMPUTERNAME
            [string]$dateTime = get-date -format MM-dd-yyyy_HH-mm
            [string]$logFileName = $computerName + "_" + $dateTime + ".txt"
            # Check for my Standard Coding Folder, and act accordingly to its presence or not
            If (Test-Path -Path "Coding:\") {
                start-transcript -path "Coding:\Source_Code\PowerShell\Transcripts\$logFileName" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
            }
            Else {
                # If Standard Coding Path is not available, ask where to save
                [string]$logPath = [Microsoft.VisualBasic.Interaction]::InputBox("Where would you like to save the transcript file?", "Trascript Path")
                If ($logPath -eq "") { # Error handling if answer is blank or Cancel is hit
                    Write-Warning "HEY!!!, No transcript save file location set, defaulting to Transcript Off!!"
                    [Microsoft.VisualBasic.Interaction]::MsgBox("Transcript is off!", "OKOnly, Exclamation",  "Transcript Status") | Out-Null
                }
                else {
                    $logPath = $logPath.TrimEnd("\")
                    Check-IfNotPathCreate($logPath)
                    start-transcript -path "$logPath\$logFileName" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
                }
            }
        }
        No {
            Write-Host "Transcripting is turned off for this console session." -BackgroundColor "darkgreen" -ForegroundColor "yellow"
        }
    }
}
# End Start-ConsoleTranscript Function
#-----------------------------------------------------------------------
# Edit-InSublime Function
function Edit-InSublime {
    param (
            [string]$fileName
        )
    . 'C:\Program Files\Sublime Text 3\sublime_text.exe' $fileName
}
# End Edit-InSublime Function
#-----------------------------------------------------------------------
# Start-RDP Function
function Start-RDP {
    param (
            [string]$ip
        )
    Start-Process -FilePath mstsc -ArgumentList "/admin /w:1024 /h:768 /v:$ip"
}
# End Start-RDP Function
#-----------------------------------------------------------------------
# Start-Chrome Function
function Start-Chrome {
    param (
            [string]$webAddy
        )
    if ($webAddy -eq "") {
        Start-Process -FilePath 'chrome.exe'
    }
    else {
        Start-Process -FilePath 'chrome.exe' $webAddy
    }
}
# End Start-Chrome Function
#-----------------------------------------------------------------------
# PSReadLine Keybind Functions

# Matched history completion using same up/down arrows
Set-PSReadlineKeyHandler -Key UpArrow -Function HistorySearchBackward
Set-PSReadlineKeyHandler -Key DownArrow -Function HistorySearchForward
Set-PSReadlineOption -HistorySearchCursorMovesToEnd

# Automatic Double "" and '', works with ctrl-z (Undo) as well
Set-PSReadlineKeyHandler -Chord 'Oem7','Shift+Oem7' `
                         -BriefDescription SmartInsertQuote `
                         -LongDescription "Insert paired quotes if not already on a quote" `
                         -ScriptBlock {
    param (
          $key, $arg
        )

    $line = $null
    $cursor = $null
    [PSConsoleUtilities.PSConsoleReadline]::GetBufferState([ref]$line, [ref]$cursor)

    if ($line[$cursor] -eq $key.KeyChar) {
        # Just move the cursor
        [PSConsoleUtilities.PSConsoleReadline]::SetCursorPosition($cursor + 1)
    }
    else {
        # Insert matching quotes, move cursor to be in between the quotes
        [PSConsoleUtilities.PSConsoleReadline]::Insert("$($key.KeyChar)" * 2)
        [PSConsoleUtilities.PSConsoleReadline]::GetBufferState([ref]$line, [ref]$cursor)
        [PSConsoleUtilities.PSConsoleReadline]::SetCursorPosition($cursor - 1)
    }
}
# End PSReadLine Keybind Functions
#-----------------------------------------------------------------------
# Script Browser Setup
if ($host.Name -eq "Windows PowerShell ISE Host") {
    #Version: 1.2.1
    Add-Type -Path 'C:\Program Files (x86)\Microsoft Corporation\Microsoft Script Browser\System.Windows.Interactivity.dll'
    Add-Type -Path 'C:\Program Files (x86)\Microsoft Corporation\Microsoft Script Browser\ScriptBrowser.dll'
    Add-Type -Path 'C:\Program Files (x86)\Microsoft Corporation\Microsoft Script Browser\BestPractices.dll'
    $scriptBrowser = $psISE.CurrentPowerShellTab.VerticalAddOnTools.Add('Script Browser', [ScriptExplorer.Views.MainView], $true)
    $scriptAnalyzer = $psISE.CurrentPowerShellTab.VerticalAddOnTools.Add('Script Analyzer', [BestPractices.Views.BestPracticesView], $true)
    $psISE.CurrentPowerShellTab.VisibleVerticalAddOnTools.SelectedAddOnTool = $scriptBrowser
}
# End Script Browser
#-----------------------------------------------------------------------

#endregion

#=======================================================================

#region Prompt Function

function prompt {
    # Sets the indicator to show role as User($) or Admin(#)
    $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = new-object System.Security.principal.windowsprincipal($CurrentUser)
    [string]$DTFormat = Get-Date -format "ddd MMM d | hh:mm:ss"

    Write-Host "`n=======================================================================" -ForegroundColor Blue
    Write-Host " ($DTFormat)" -ForegroundColor Yellow
    if ($Principal.IsInRole("Administrators")) {
        Write-Host " {#}-{$($(Get-Location).Path)}" -ForegroundColor Red
    }
    else {
        Write-Host " {$}-{$($(Get-Location).Path.replace($HOME,'~'))}" -ForegroundColor Green
    }
    Write-Host $(if ($nestedpromptlevel -ge 1) { '>>' }) -NoNewline
    Write-Host "=======================================================================" -ForegroundColor Blue
    return "--> "
}

#endregion

#=======================================================================

#region Set Aliases

Set-Alias gh Get-Help
Set-Alias st Stop-Transcript
Set-Alias subl Edit-InSublime
Set-Alias rdp Start-RDP
Set-Alias chrome Start-Chrome

#endregion

#=======================================================================

#region Call Commands/Functions

# Clear Screen
#cls

# Start the console Transcript command
if ($host.Name -eq "ConsoleHost") {
    Start-ConsoleTranscript
}
#endregion

