#region Set Windows Size and Shape in Console

if ($host.Name -eq 'ConsoleHost') {
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

#region Set Colors of 'Write-' Messages

$colors = $host.PrivateData
$colors.VerboseForegroundColor = 'white'
$colors.VerboseBackgroundColor = 'blue'
$colors.WarningForegroundColor = 'yellow'
$colors.WarningBackgroundColor = 'darkgreen'
# Comment out the next two lines to set the default Fore=Red, Back=black
$colors.ErrorForegroundColor = 'white'
$colors.ErrorBackgroundColor = 'red'

#endregion

#=======================================================================

#region Set PSDrives

# Checks paths to my Coding folders, and sets 'Coding' PSDrive to them if found
If (Test-Path -Path 'Z:\Deuce\Dropbox\Jon\Coding') {
    new-psdrive -name Coding -psprovider FileSystem -root 'Z:\Deuce\Dropbox\Jon\Coding' | Out-Null
    Set-Location Coding:
}
ElseIf (Test-Path -Path 'F:\dropbox\jon\coding') {
    new-psdrive -name Coding -psprovider FileSystem -root 'F:\dropbox\jon\coding' | Out-Null
    Set-Location Coding:
}
Else {
    set-location $env:USERPROFILE\desktop
}

#endregion

#=======================================================================

#region Import Modules

#region User Module Directory

#endregion
#-----------------------------------------------------------------------
#region System Module Directory

# Load ActiveDirectory module from system modules directory if available
if(get-module -ListAvailable -Name 'ActiveDirectory') {
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
function Start-ConsoleTranscript {

    # Check if I want to start a Transcript, Used if I am trying out commands
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    [string]$tranYN = [Microsoft.VisualBasic.Interaction]::MsgBox('Would you Like to Start a Transcript?', 'YesNo, DefaultButton2',  'Start Transcript?')
    # Check the answer for 'Yes' or 'No' and act upon the choice
    switch ($tranYN) {
        Yes {
            [string]$computerName = $env:COMPUTERNAME
            [string]$dateTime = get-date -format MM-dd-yyyy_HH-mm
            [string]$logFileName = $computerName + '_' + $dateTime + '.txt'
            # Check for my Standard Coding Folder, and act accordingly to its presence or not
            If (Test-Path -Path 'Coding:\') {
                start-transcript -path 'Coding:\Source_Code\PowerShell\Transcripts\$logFileName' | Write-Host -BackgroundColor 'darkgreen' -ForegroundColor 'yellow'
            }
            Else {
                # If Standard Coding Path is not available, ask where to save
                [string]$logPath = [Microsoft.VisualBasic.Interaction]::InputBox('Where would you like to save the transcript file?', 'Trascript Path')
                If ($logPath -eq '') { # Error handling if answer is blank or Cancel is hit
                    Write-Warning 'HEY!!!, No transcript save file location set, defaulting to Transcript Off!!'
                    [Microsoft.VisualBasic.Interaction]::MsgBox('Transcript is off!', 'OKOnly, Exclamation',  'Transcript Status') | Out-Null
                }
                else {
                    $logPath = $logPath.TrimEnd('\')
                    Check-IfNotPathCreate($logPath)
                    start-transcript -path '$logPath\$logFileName' | Write-Host -BackgroundColor 'darkgreen' -ForegroundColor 'yellow'
                }
            }
        }
        No {
            Write-Host 'Transcripting is turned off for this console session.' -BackgroundColor 'darkgreen' -ForegroundColor 'yellow'
        }
    }
}
# End Start-ConsoleTranscript Function
#-----------------------------------------------------------------------
# Start-RDP Function
function Start-RDP {
    param (
            [string]$ip
        )
    Start-Process -FilePath mstsc -ArgumentList '/admin /w:1024 /h:768 /v:$ip'
}
# End Start-RDP Function
#-----------------------------------------------------------------------
# Start-Chrome Function
function Start-Chrome {
    param (
            [string]$webAddy
        )
    if ($webAddy -eq '') {
        Start-Process -FilePath 'chrome.exe'
    }
    else {
        Start-Process -FilePath 'chrome.exe' $webAddy
    }
}
# End Start-Chrome Function
#-----------------------------------------------------------------------
# Goto-Desktop Function
function Goto-Desktop {
    Set-Location $env:USERPROFILE\Desktop
}
# End Goto-Dekstop Function
#-----------------------------------------------------------------------
# Goto-Repos Function
function Goto-Repos {
    Set-Location Coding:\Repos
}
# End Goto-Repos Function
#-----------------------------------------------------------------------
# Out-Notepad Function
function Out-Notepad {
      param (
            [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
            [Object]
            [AllowEmptyString()] 
            $Object,

            [Int]
            $Width = 150
      )

      begin {
                $al = New-Object System.Collections.ArrayList
      }

      process {
        $null = $al.Add($Object)
      }
      end {
            $text = $al | 
            Format-Table -AutoSize -Wrap | 
            Out-String -Width $Width

            $process = Start-Process notepad -PassThru
            $null = $process.WaitForInputIdle()


            $sig = '
              [DllImport("user32.dll", EntryPoint = "FindWindowEx")]public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
              [DllImport("User32.dll")]public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);
            '

            $type = Add-Type -MemberDefinition $sig -Name APISendMessage2 -PassThru
            $hwnd = $process.MainWindowHandle
            [IntPtr]$child = $type::FindWindowEx($hwnd, [IntPtr]::Zero, "Edit", $null)
            $null = $type::SendMessage($child, 0x000C, 0, $text)
      }
}
# End Out-Notepad Function
#-----------------------------------------------------------------------
# Start Get-GitInfoForDirectory Function
function Get-GitInfoForDirectory {

    param (
    )

    begin {
        $gitBranch = (git branch)
        $gitStatus = (git status)
        $gitTextLine = ""
    }

    process {
        try {
            foreach ($branch in $gitBranch) {
                if ($branch -match '^\* (.*)') {
                    $gitBranchName = 'Git Repo - Branch: ' + $matches[1].ToUpper()
    	        }
            }
    
            if (!($gitStatus -like '*working directory clean*')) {
                $gitStatusMark = ' ' + '/' + ' Status: ' + 'NEEDS UPDATING'
            }
            elseif ($gitStatus -like '*Your branch is ahead*') {
                $gitStatusMark = ' ' + '/' + ' Status: ' + 'PUBLISH COMMITS'
            }
            else {
                $gitStatusMark = ' ' + '/' + ' Status: ' + 'UP TO DATE'
            }
        }
        catch {
        }
    }

    end {
        if ($gitBranch) { 
            $gitTextLine = "`n" + ' {' + $gitBranchName + $gitStatusMark + '}'            
        }
        return $gitTextLine       
    }    
}
# End Get-GitInfoForDirectory Function
#-----------------------------------------------------------------------
# Show-MyAliasList Function
function Show-MyAliasList {
    Write-Output {
            MY CUSTOM ALIASES/FUNCTIONS

            chrome  Start-Chrome
            gh      Get-Help
            gtd     Goto-Desktop
            notepad Out-Notepad
            rdp     Start-RDP
            repos   Goto-Repo
            rt      Start-ConsoleTranscript
            st      Stop-Transcript
         }
}
# End Show-MyAliasList Function

#endregion

#=======================================================================

#region Prompt Function

function prompt {
    # Sets the indicator to show role as User($) or Admin(#)
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = new-object System.Security.principal.windowsprincipal($currentUser)
    [string]$dateFormat = Get-Date -format "ddd MMM d | hh:mm:ss"

    Write-Host "`n=======================================================================" -ForegroundColor Blue
    Write-Host " ($dateFormat)" -ForegroundColor Yellow
    if ($principal.IsInRole('Administrators')) {
        Write-Host " {#}-{$($(Get-Location).Path)}" -NoNewline -ForegroundColor Red        
    }
    else {
        Write-Host " {$}-{$($(Get-Location).Path.replace($HOME,'~'))}" -NoNewline -ForegroundColor Green        
    }
    write-Host $(Get-GitInfoForDirectory) -ForegroundColor Magenta
    Write-Host $(if ($nestedpromptlevel -ge 1) { '>>' }) -NoNewline
    Write-Host "=======================================================================" -ForegroundColor Blue
    return "--> "
}

#endregion

#=======================================================================

#region ISE Addon Configurations

# Script Browser Setup
if ($host.Name -eq 'Windows PowerShell ISE Host') {
    #Version: 1.3.2
    Add-Type -Path 'C:\Program Files (x86)\Microsoft Corporation\Microsoft Script Browser\System.Windows.Interactivity.dll'
    Add-Type -Path 'C:\Program Files (x86)\Microsoft Corporation\Microsoft Script Browser\ScriptBrowser.dll'
    Add-Type -Path 'C:\Program Files (x86)\Microsoft Corporation\Microsoft Script Browser\BestPractices.dll'
    $scriptBrowser = $psISE.CurrentPowerShellTab.VerticalAddOnTools.Add('Script Browser', [ScriptExplorer.Views.MainView], $true)
    $scriptAnalyzer = $psISE.CurrentPowerShellTab.VerticalAddOnTools.Add('Script Analyzer', [BestPractices.Views.BestPracticesView], $true)
    $psISE.CurrentPowerShellTab.VisibleVerticalAddOnTools.SelectedAddOnTool = $scriptBrowser
}
# End Script Browser Setup

#endregion

#=======================================================================

#region Set Aliases

Set-Alias chrome Start-Chrome
Set-Alias gh Get-Help
Set-Alias gtd Goto-Desktop
Set-Alias ma Show-MyAliasList
Set-Alias notepad Out-Notepad
Set-Alias repos Goto-Repos
Set-Alias rdp Start-RDP
Set-Alias rt Start-ConsoleTranscript
Set-Alias st Stop-Transcript

#endregion

#=======================================================================

#region Command/Function Call

if ($host.Name -eq 'ConsoleHost') {
    Start-ConsoleTranscript
    Show-MyAliasList
}

#endregion

