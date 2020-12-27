function Invoke-cmdWithProgressBar {
<#
    .SYNOPSIS
        Run a long running command with a progress bar, if there is no other way to use write-progress bar.

    .DESCRIPTION
        Run a long running command with a progress bar, if there is no other way to use write-progress bar.

        For Example  "Get-ChildItem $sbParam -Recurse -Directory"

        This command will take a long time to Enummerate the folders if there are many (100 or 1000s)
        folders in the searched location, and there is no way to give the user a way of knowing its running.

        This function will allow user feedback during the enumeration.

    .PARAMETER  Message
        Message to show on the progress bar while it runs. Defaults to "Running Command, Please wait..."

    .PARAMETER  sbParam
        A parameter passed to use in the scriptblock command

        This function is designed for 0 or 1 parameter to be passed, usually a path like in Get-ChildItem
        To increase parameters, add a param in the param block and then add a call to .addparameter in
        the ".addscript" line (See commenting examples addition in the function). Ensure the Scriptblock
        sent to "Task" has the parameter listed. Finally pass the parameter when calling the function.
#>
        param
        (
            [Parameter(Mandatory = $true)]
            [scriptblock]$task,
            [Parameter(Mandatory = $false)]
            [string]$message = "Running Command.",
            [Parameter(Mandatory = $false)]
            [string]$idNum = "1",
            [Parameter(Mandatory = $false)]
            [string]$sbParam
            <#
                Example:
                [Parameter(Mandatory = $false)]
                [string]$sbParam2
        #>
        )

        Begin {
            # Create PowerShell Runspace
            [powershell]$ps = [PowerShell]::Create()
            # Add script and parameters to runspace
            $null = $ps.AddScript($task).AddParameter("sbParam", $sbParam)#.AddParameter("sbParam2", $sbparam2)
        }
        Process {
            # Start the runspace and grab is report status
            $status = $ps.BeginInvoke()

            # Set loop count
            $count = 0

            # Loop to check if status is completed
            while (!$status.IsCompleted) {
                Write-Progress -Id $idNum -Activity $message -Status "Please Wait..." -PercentComplete ($count % 100)
                $count++
                Start-Sleep -Milliseconds 300
            }

            $ps.EndInvoke($status)
        }
        End {
            $ps.Stop()
            $ps.Runspace.Close()
            $ps.Dispose()
        }
    }


# Just for testing and example do not use except for example, Uncomment and change $paths value to test
<# $paths = Get-Content 'C:\pathList.txt'

foreach ($path in $paths) {

    $scriptBlock = {
        param([string]$sbParam)
        Get-ChildItem $sbParam -Recurse -Directory -ErrorAction SilentlyContinue
    }

    $folders = Invoke-cmdWithProgressBar -task $scriptBlock -message "Gathering Data on Folders, one Moment..." -sbParam $path

    Out-Host $folders
    Clear-Variable folders
} #>