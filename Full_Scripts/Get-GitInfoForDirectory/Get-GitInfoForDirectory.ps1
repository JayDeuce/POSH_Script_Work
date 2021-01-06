<#
.SYNOPSIS
    Uses Git for windows to gather info about the current directory and updates the prompt function with visual 'Status' and 'Branch Name' indicators.

.DESCRIPTION
    The function will use the installed Git For Windows application to
    gather repository data about the current directory and update the Prompt
    function to include visual "Status" and "Branch" indicators. If the current
    is not a Git Repository, nothing will change on the prompt, as it will run
    checks to see if it is a repository or not before changing the prompt.

    Installation and Requirements:
        The function requires Git For Windows to be installed and its install
        directory added to the PATH Variable. To use the function, it can be added
        to your PowerShell profile script directly or dot sourced.

        Here is an Example Prompt Function, where this function is called after
        the current location info is set:

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

    The function will run automatically each time the prompt function is called. Because of this,
    there are no Parameters associated with this function.

    To customize the Status Messages, change the strings of the $gitStatusMark variable.

.EXAMPLE
    Get-GitInfoForDirectory

    Description:

       Run the Script and update prompt with info.

.NOTES
    Name: Get-GitInfoForDirectory
    Author: Jonathan Durant
    Version: 1.1
    DateUpdated: 6 Jan 2021

.INPUTS
    NONE
.OUTPUTS
    Message String
#>

<#PSScriptInfo

.VERSION 1.1

.GUID fe3b786e-6bcb-409b-8f37-d1297bf2dd35

.AUTHOR JayDeuce

.COMPANYNAME

.COPYRIGHT

.TAGS git

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
Requires "Git For Windows" -> https://git-scm.com/download/win

#>


<#

.DESCRIPTION
 Uses Git for windows to gather info about the current directory and updates the prompt function with visual 'Status' and 'Branch Name' indicators.

#>

function Get-GitInfoForDirectory {

    param (
    )

    begin {
        git remote update | Out-Null
        $gitBranch = (git branch)
        $gitStatus = (git status)
        $gitTextLine = ""
    }

    process {
        try {
            foreach ($branch in $gitBranch) {
                if ($branch -match "^\* (.*)") {
                    $gitBranchName = "Git Repo - Branch: " + $matches[1].ToUpper()
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
    }

    end {
        if ($gitBranch) {
            $gitTextLine = "{" + $gitBranchName + $gitStatusMark + "}"
        }
        return $gitTextLine
    }
}