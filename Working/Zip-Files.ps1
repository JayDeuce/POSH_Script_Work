Param(
    [string]$zipfilename, 
    [string[]]$fileToAdd
)

$fAdd = Get-ChildItem -Path $fileToAdd 

if(-not (test-path($zipfilename))) {
    set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipfilename).IsReadOnly = $false
}

$zipFile = Resolve-Path $zipfilename

$shellApplication = new-object -com shell.application
$zipPackage = $shellApplication.NameSpace($zipFile.Path)

foreach($file in $fAdd) {       

        $zipPackage.CopyHere($file.FullName, 16)
        Start-sleep -milliseconds 500

}


# TODO:
# 1. Files ask for permission to copy via popup window
#    if file is already in archive. need to find a force switch or something
# 2. DONE More than one file at a time.
# 3. Keep Folder Structures