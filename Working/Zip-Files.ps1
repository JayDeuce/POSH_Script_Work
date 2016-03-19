Param(
    [string]$zipfilename, 
    [string]$fileToAdd
)

$fileAdd = Get-ChildItem -Path $fileToAdd 

if(-not (test-path($zipfilename))) {
    set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (dir $zipfilename).IsReadOnly = $false
}

$zipFile = Resolve-Path $zipfilename

$shellApplication = new-object -com shell.application
$zipPackage = $shellApplication.NameSpace($zipFile.Path)

foreach($file in $fileAdd) {
        

        $zipPackage.CopyHere($file.FullName)
        Start-sleep -milliseconds 500
}

