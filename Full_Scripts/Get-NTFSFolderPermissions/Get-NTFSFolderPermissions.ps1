#Requires -Version 3

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string[]]$rootPath,
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [string[]]$csvReportPath = "$ENV:USERPROFILE\Documents",
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [string[]]$csvReportName = "Get-NTFSPermissions",
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [switch]$includeFlags = $false,
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [switch]$includeOwner = $false,
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [switch]$consoleOutput = $false,
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [switch]$xlsxOutput = $false
)

Begin {
    # INTERNAL FUNCTIONS
    function Test-IfNotPathCreate([string]$FolderPath) {
        # Check the passed folder path, if is does not exist create it
        if (!(Test-Path -Path $FolderPath)) {
            New-Item -Path $FolderPath -ItemType directory | Out-Null
        }
    }
    function Get-FormattedDate() {
        # Build a formatted date for use with variables
        [string]$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        return $date
    }
    function Convert-CsvToXls {
        # Function to convert CSV Files to XSLX. NOTE: Must have MS Office installed
        Param (
            [string]$inputFilePath,
            [string]$outputFilePath
        )

        # Create a new Excel Workbook with one empty sheet
        $excel = New-Object -ComObject excel.application
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add(1)
        $worksheet = $workbook.worksheets.Item(1)

        # Build the QueryTables.Add command
        # QueryTables does the same as when clicking "Data » From Text" in Excel
        $txtConnector = ("TEXT;" + $inputFilePath)
        $connector = $worksheet.QueryTables.add($txtConnector, $worksheet.Range("A1"))
        $query = $worksheet.QueryTables.item($connector.name)

        # Set the delimiter (, or ;) according to your regional settings
        $query.TextFileOtherDelimiter = $excel.Application.International(5)

        # Set the format to delimited and text for every column
        # A trick to create an array of 2s is used with the preceding comma
        $query.TextFileParseType = 1
        $query.TextFileColumnDataTypes = , 2 * $worksheet.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        # Execute & delete the import query
        $query.Refresh() | Out-Null
        $query.Delete()

        # Save & close the Workbook as XLSX. Change the output extension for Excel 2003
        $workbook.SaveAs($outputFilePath, 51)
        $excel.Quit()
    }
    function Invoke-cmdletWithProgressBar {
        <#    .SYNOPSIS
        Run a long running command with a progress bar, if there is no other way to use write-progress bar.

    .DESCRIPTION
        Run a long running command with a progress bar, if there is no other way to use write-progress bar.

        For Example  "Get-ChildItem $sbParam -Recurse -Directory"

        This command will take a long time to Enummerate the folders if there are many (100 or 1000s I mean)
        folders in the searched location, and there is no way to give the user a way of knowing its running.

        This function will allow user feedback during the enumeration.

    .PARAMETER  Message
        Message to show on the progress bar while it runs. Defaults to "Running Command, Please wait..."

    .PARAMETER  sbParam
        A parameter passed to use in the scriptblock command

        This function is designed for 0 or 1 parameter to be passed, usually a path like in Get-ChildItem
        To increase parameters, add a param in the param block and then add a call to .addparameter in
        the ".addscript" line (See commentign examples addition in the function). Ensure the Scriptblock
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
            #[Parameter(Mandatory = $false)]
            #[string]$sbParam2
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
    # END INTERNAL FUNCTIONS
}

Process {

    # Set Variables for $paths loop
    $allPaths = ($rootPath | Measure-Object).Count
    $currentPath = 0

    # Check the report folder paths and create if necessary
    Test-IfNotPathCreate($csvReportPath)

    # Paths loop
    foreach ($path in $rootPath) {

        # Invoke-cmdletWithProgressBar Parameters
        $scriptBlockPathLoop = {
            param([string]$sbParam)
            Get-ChildItem $sbParam -Recurse -Directory -ErrorAction SilentlyContinue
        }
        $messagePathLoop = "Gathering Data on Folders."
        $idNumPathLoop = "2"

        # Add to path loop Count
        $currentPath++

        # Set Filename for Output file, Uses the folder name as part of the file name.
        if ($xlsxOutput -eq $True) {
            [string]$xlsxReport = "$csvReportPath\$csvReportName" + "-" + (Split-Path $path -Leaf) + "-" + (Get-FormattedDate) + ".xlsx"
        }
        else {
            [string]$csvReport = "$csvReportPath\$csvReportName" + "-" + (Split-Path $path -Leaf) + "-" + (Get-FormattedDate) + ".csv"
        }

        # Progress Bar for Path Listing
        Write-Progress -Id 1 -Activity "Processing Path $currentpath / $allPaths" -Status "Working on on the following path: $path. One moment please...."  -PercentComplete (($currentPath / $allPaths) * 100)

        # Run Get-ChildItem cmdlet with a progress bar (Can be a long run time while gathering a large amount of folders), asigning output to $folders variable
        $folders = Invoke-cmdletWithProgressBar -task $scriptBlockPathLoop -message $messagePathloop -idNum $idNumPathLoop -sbParam $path

        # Set variables for $folders loop
        $totalFolders = ($folders | Measure-Object).Count
        $currentFolder = 0

        # Folders loop
        foreach ($folder in $folders) {
            # Add to folders loop count
            $currentFolder++

            # Progress bar for getting ACL Lists
            Write-Progress -Id 2 -Activity "Getting Access Control Lists for each Folder." -Status "Working on $currentFolder / $totalFolders" -PercentComplete (($currentFolder / $totalFolders) * 100)

            # Pull all ACL data
            $aclList = Get-Acl $folder.fullname -ErrorAction SilentlyContinue -ErrorVariable AccessDenied | ForEach-Object { $_.Access  }

            # ACLs Loop
            foreach ($acl in $aclList) {

                # Build ACL variables
                $fullname = ($folder.Fullname).replace(',', '<comma-removed>')
                $identityReference = $acl.IdentityReference
                $accessControlType = $acl.AccessControlType
                $fileSystemRights = (((($acl.FileSystemRights) | Out-String).Replace(",", "/")).Replace("`r`n", ""))
                $isInherited = $acl.IsInherited
                $inheritenceFlags = $acl.InheritanceFlags
                $propagationFlags = $acl.PropagationFlags
                $folderOwner = Get-Acl $folder.FullName -ErrorAction SilentlyContinue -ErrorVariable AccessDenied | ForEach-Object { $_.Owner }

                # Set acl Information to the $aclInfo variable
                $aclInfo = New-Object -TypeName PSObject
                $aclInfo | Add-Member -MemberType NoteProperty -Name "Full Folder Name" -Value $fullname
                $aclInfo | Add-Member -MemberType NoteProperty -Name "User-Group Name" -Value $identityReference
                $aclInfo | Add-Member -MemberType NoteProperty -Name "Access Control Type" -Value $accessControlType
                $aclInfo | Add-Member -MemberType NoteProperty -Name "File System Rights" -Value $fileSystemRights
                $aclInfo | Add-Member -MemberType NoteProperty -Name "Is Inherited from Parent" -Value $isInherited
                If ($includeFlags -eq $true) {
                    $aclInfo | Add-Member -MemberType NoteProperty -Name "Inherited Flags" -Value $inheritenceFlags
                    $aclInfo | Add-Member -MemberType NoteProperty -Name "Propagation Flags" -Value $propagationFlags
                }
                If ($includeOwner -eq $true) {
                    $aclInfo | Add-Member -MemberType NoteProperty -Name "Owner of the Folder" -Value $folderOwner
                }

                # Write results array
                [array]$results += $aclInfo

                # Clear used loop variables
                Clear-Variable fullname, identityReference, accessControlType, fileSystemRights, inheritenceFlags, propagationFlags, folderOwner
            }
        }

        # Per parameters push $results to output
        if ($null -ne $results) {
            if ($consoleOutput -eq $True) {
                Out-Host -InputObject $results
            }
            Elseif ($xlsxOutput -eq $true) {
                # Convert CSV to XSLX. NOTE: Must have MS Office installed
                $tempCSV = "$env:TEMP\Get-ComputerInfoTempReport.csv"
                $results | Export-Csv -Path $tempCSV -NoTypeInformation -Force
                Write-Progress -Id 2 -Activity "Exporting Results to Microsoft Excel." -Status "One Moment Please..." -PercentComplete 50
                Convert-CsvToXls -inputFilePath $tempCSV -outputFilePath $xlsxReport
                Write-Progress -Id 2 -Activity "Exporting Results to Microsoft Excel." -Status "One Moment Please..." -PercentComplete 100
                Remove-Item -Path $tempCSV
            }
            Else {
                $results | Export-Csv -Path $csvReport -NoTypeInformation -Force
            }
        }

        # Clear used loop variables
        Clear-Variable results
    }
}

End {

}