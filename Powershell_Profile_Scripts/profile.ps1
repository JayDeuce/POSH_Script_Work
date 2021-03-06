#region Set Windows Size and Shape in Console

if ($host.Name -eq "ConsoleHost") {
    # Set Variable
    $pshost = Get-Host
    $pswindow = $pshost.ui.rawui
    if (!($ENV:WT_SESSION)) {
        # Set Screen Buffer Size
        $newsize = $pswindow.buffersize
        $newsize.height = 300
        $newsize.width = 150
        $pswindow.buffersize = $newsize

        # Set Windows Size
        $newsize = $pswindow.windowsize
        $newsize.height = 38
        $newsize.width = 150
        $pswindow.windowsize = $newsize
    }
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

#region Script Variables

$timeset = $null # For use in Get-GitInfoForDirectory

#endregion

#=======================================================================

#region Import Modules

if (Get-Module -ListAvailable -Name "ActiveDirectory") {
    Import-Module ActiveDirectory
}

if (Get-Module -ListAvailable -Name "posh-sshell") {
    Import-Module Posh-SSHell
    $sshell = $true
}
else {
    $sshell = $false
}

#endregion

#=======================================================================

#region Helper Functions

#-----------------------------------------------------------------------
# Test-IfNotPathCreate Function
function Test-IfNotPathCreate([string]$FolderPath) {
    if (!(Test-Path -Path $FolderPath)) {
        New-Item -Path $FolderPath -ItemType directory | Out-Null
    }
}
# End Test-IfNotPathCreate Function
#-----------------------------------------------------------------------
# Start-ConsoleTranscript Function
function Start-ConsoleTranscript {
    # Check if I want to start a Transcript, Used if trying out commands
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
    [string]$tranYN = [Microsoft.VisualBasic.Interaction]::MsgBox("Would you Like to Start a Transcript?", "YesNo, DefaultButton2", "Start Transcript?")
    # Check the answer for "Yes" or "No" and act upon the choice
    switch ($tranYN) {
        Yes {
            [string]$computerName = $env:COMPUTERNAME
            [string]$dateTime = Get-Date -Format MM-dd-yyyy_HH-mm
            [string]$logFileName = $computerName + "_" + $dateTime + ".txt"

            # Ask where to save
            [string]$logPath = [Microsoft.VisualBasic.Interaction]::InputBox("Where would you like to save the transcript file?", "Trascript Path")
            If ($logPath -eq "") {
                # Error handling if answer is blank or Cancel is hit
                Write-Warning "HEY!!!, No transcript save file location set, defaulting to Transcript Off!!"
                [Microsoft.VisualBasic.Interaction]::MsgBox("Transcript is off!", "OKOnly, Exclamation", "Transcript Status") | Out-Null
            }
            else {
                $logPath = $logPath.TrimEnd("\")
                Test-IfNotPathCreate($logPath)
                Start-Transcript -Path "$logPath\$logFileName" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
            }
        }
        No {
            Write-Host "Transcripting is turned off for this console session." -BackgroundColor "darkgreen" -ForegroundColor "yellow"
        }
    }
}
# End Start-ConsoleTranscript Function
#-----------------------------------------------------------------------
# Start-RDP Function
function Start-RDP {
    <#
    .SYNOPSIS
        Start Remote Desktop Protocol
    .PARAMETER ip
        The IP or network name of the remote machine
    .EXAMPLE
        "192.168.1.2" | rdp
    .EXAMPLE
        rdp "192.168.1.2"
    #>
    param (
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true)]
        [string]$IP
    )
    process {
        Start-Process -FilePath mstsc -ArgumentList "/admin /v:$ip"
    }
}
# End Start-RDP Function
#-----------------------------------------------------------------------
# Start-Edge Function
function Start-Edge {
    param (
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true)]
        [string]$webAddress
    )
    if ($webAddress -eq "") {
        Start-Process -FilePath "msedge.exe"
    }

    else {
        Start-Process -FilePath "msedge.exe" $webAddress
    }
}
# End Start-Edge Function
#-----------------------------------------------------------------------
# Set-DesktopPath Function
function Set-DesktopPath {
    Set-Location "$env:USERPROFILE\Desktop"
}
# End Set-DesktopPath Function
#-----------------------------------------------------------------------
# Set-RepoPath Function
function Set-RepoPath {
    if ((Test-Path -Path F:\Coding)) {
        Set-Location "F:\Coding\"
    }
    else {
        Set-Location "$env:USERPROFILE\Documents\Coding"
    }

}
# End Set-RepoPath Function
#-----------------------------------------------------------------------
# Out-Notepad Function
# Send pipeline data to Notepad, does not open files
function Out-Notepad {
    <#
    .SYNOPSIS
        Send object as string to Notepad App

    .DESCRIPTION
        This function takes the parameter object, converts
        it to a formatted string, and sends it to a notepad
        file for easier viewing and manipulation of output.
        (Can also be sent as pipline output by sending through
        Out-String first.)

    .PARAMETER object
        (Required, No default)

        THis parameter is the object data you wish to convert and send to notepad

        Examples: $object
                  (get-process)
                  "text in a sentence"
                  "Textword"

    .EXAMPLE
        Out-Notepad $object

        Takes the $object (Can be anything) and sends it to notepad

    .EXAMPLE
        Out-Notepad (Get-Process)

        Runs the command Get-process and sends it to notepad.

    .EXAMPLE
        $object | Out-Notepad

        Takes the $object Variable, converts it to string and sends it to notepad.

    .EXAMPLE
        (Get-Process) | Out-Notepad

        Runs the command Get-process, converts it to string and sends it to notepad.
    #>
    param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true)]
        [Object][AllowEmptyString()]$Object
    )

    begin {
        [Int]$Width = 150
        $al = New-Object System.Collections.ArrayList
    }

    Process {
        $null = $al.Add($Object)
        $text = $al | Format-Table -AutoSize -Wrap | Out-String -Width $Width
    }

    end {
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
    # Initialization of Variables
    if (($null -eq $timeset) -or (($timeset).minutes -le (Get-date).TimeOfDay.Minutes - 5)) {
        git remote update | Out-Null
        $script:timeset = (Get-Date).TimeOfDay
    }
    $gitRepo = ((Split-Path -Leaf (git remote get-url origin)).Split(".")[0]).ToUpper()
    $gitBranch = (git branch)
    $gitStatus = (git status)
    $gitTextLine = ""
    # Gather Data on Repo and build Variables
    try {
        foreach ($branch in $gitBranch) {
            if ($branch -match "^\* (.*)") {
                $gitRepoBranchName = "Repo: " + $gitRepo + " - Branch: " + $matches[1].ToUpper()
            }
        }
        if (!($gitStatus -like "*working tree clean*")) {
            $gitStatusMark = " " + "/" + " Status: " + "NEEDS UPDATING"
        }
        elseif ($gitStatus -like "*Your branch is ahead*") {
            $gitStatusMark = " " + "/" + " Status: " + "PUBLISH COMMITS"
        }
        elseif ($gitstatus -like "*Your branch is behind*") {
            $gitstatusMark = " " + "/" + " Status: " + "NEED TO PULL"
        }
        else {
            $gitStatusMark = " " + "/" + " Status: " + "UP TO DATE"
        }
    }
    catch {

    }
    # Return proper git text line for prompt
    if ($gitBranch) {
        $gitTextLine = "{" + $gitRepoBranchName + $gitStatusMark + "}"
    }
    return $gitTextLine
}
# End Get-GitInfoForDirectory Function
#-----------------------------------------------------------------------
# Start Get-PSVersion Function
function Get-PSVersion {
    $version = $PSVersionTable
    Write-Host ($version | Format-Table | Out-String)
}
# End Get-PSVersion Function
#-----------------------------------------------------------------------
# Start Show-MyAliasList Function
function Show-MyAliasList {
    Write-Output {
        MY CUSTOM ALIASES/FUNCTIONS

        edge    Start-Edge          rdp     Start-RDP
        gh      Get-Help            repo    Set-RepoPath
        gtd     Set-DesktopPath     staTr   Start-ConsoleTranscript
        notepad Out-Notepad         stoTr   Stop-Transcript
        psv     Get-PSVersion       sma     Show-MyAliasList
                                    smga    Show-MyGitAlias *When Git Installed*
    }
}
# End Show-MyAliasList Function
#-----------------------------------------------------------------------
# Start Git Functions
function Get-GitStatus {
    git status
}
function set-GitAdd {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$filename
    )
    git add $filename
}
function set-GitAddAll {
    git add -A
}
function Set-GitCommitMessage {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$message
    )
    git commit -m $message
}
function Set-GitCommitAmend {
    git commit --amend
}
function Set-GitPush {
    git push
}
function Get-GitPull {
    git pull
}
function Get-GitLog {
    git log
}
function Get-GitFetch {
    git fetch
}
function Get-GitDiff {
    git diff
}
function Show-MyGitAlias {
    Write-Output {
        MY GIT ALIASES

        gits   Get-GitStatus          gitf   Get-GitFetch
        gita   Set-GitAdd             gitpl  Get-GitPull
        gitaa  set-GitAddAll          gitpu  Set-GitPush
        gitcm  Set-GitCommitMessage   gitl   Get-GitLog
        gitcam Set-GitCommitAmend
    }

}

#endregion

#=======================================================================

#region Prompt Function

function prompt {
    # Sets the indicator to show role as User($) or Admin(#)
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.principal.windowsprincipal($currentUser)
    [string]$dateFormat = Get-Date -Format "ddd MMM d | hh:mm:ss"

    Write-Host "`n=======================================================================" -ForegroundColor Blue
    Write-Host "($dateFormat)" -ForegroundColor Yellow
    if ($principal.IsInRole("Administrators")) {
        Write-Host "{#}-{$($(Get-Location).Path)}" -ForegroundColor Red
    }
    else {
        Write-Host "{$}-{$($(Get-Location).Path.replace($HOME,"~"))}" -ForegroundColor Green
    }
    if (Test-Path -Path ".\.git") {
        Write-Host $(Get-GitInfoForDirectory) -ForegroundColor Magenta
    }
    Write-Host "=======================================================================" -ForegroundColor Blue

    if ($nestedpromptlevel -ge 1) {
        Write-Host ">> (Nested)" $nestedpromptlevel "<<" -ForegroundColor Cyan
    }
    return "--> "
}

#endregion

#=======================================================================

#region Set Aliases

Set-Alias edge Start-Edge
Set-Alias gh Get-Help
Set-Alias gtd Set-DesktopPath
Set-Alias ma Show-MyAliasList
Set-Alias notepad Out-Notepad
Set-Alias psv Get-PSVersion
Set-Alias repo Set-RepoPath
Set-Alias rdp Start-RDP
Set-Alias staTr Start-ConsoleTranscript
Set-Alias stoTr Stop-Transcript
Set-Alias sma Show-MyAliasList

# Git Aliases
if ("C:\Program Files\Git\cmd\git.exe") {
    Set-Alias smga Show-MyGitAlias
    Set-Alias gits Get-GitStatus
    Set-Alias gitpl Get-GitPull
    Set-Alias gita  Set-GitAdd
    Set-Alias gitaa Set-GitAddAll
    Set-Alias gitcm Set-GitCommitMessage
    Set-Alias gitcam Set-GitCommitAmend
    Set-Alias gitpu Set-GitPush
    Set-Alias gitl  Get-GitLog
    Set-Alias gitf Get-GitFetch
}
#endregion

#=======================================================================

#region Command/Function Call

if ($host.Name -eq "ConsoleHost") {

    Show-MyAliasList

    if (Get-Service -Name ssh-agent) {
        Start-SshAgent -Quiet
    }
    else {
        Write-Warning "Development Environment missing SSH-Agent. Install Git for Windows."
    }

    if ($sshell -eq $false) {
        Write-Warning "Development Environment missing Posh-SSHell. Install Module."
    }

}

#endregion
