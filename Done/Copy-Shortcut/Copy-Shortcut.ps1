[cmdletbinding()]

Param (
    [Parameter(mandatory=$true,Position=0)]
    [array]$computers = "",
    [Parameter(mandatory=$true,Position=1)]
    [string]$ShortcutName = "",
    [Parameter(mandatory=$false,Position=2)]
    [string]$ShortcutPath = "C:\Scripts"
)
process {
    foreach ($computer in $computers) {
        Write-Host "`nProcessing $computer..."
        if (!(test-connection -count 1 -Quiet -ComputerName $computer)) {
            Write-Host "`n$computer could not be contacted, please check make sure its online. Moving on...`n"
            Write-Host "+++++++++++++++++++"
        }
        else {                        
            Copy-Item $ShortcutPath\$ShortcutName "\\$computer\c$\Users\Public\Desktop"                              
        }
    }
}
