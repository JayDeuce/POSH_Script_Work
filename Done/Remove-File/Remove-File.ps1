[cmdletbinding()]

Param (
     [Parameter(mandatory = $true, Position = 0)]
     [array]$computers = "",
     [Parameter(mandatory = $true, Position = 1)]
     [string]$ShortcutName = ""
)
process {
     foreach ($computer in $computers) {
          Write-Host "`nProcessing $computer..."
          if (!(test-connection -count 1 -Quiet -ComputerName $computer)) {
               Write-Host "`n$computer could not be contacted, please check make sure its online. Moving on...`n"
               Write-Host "+++++++++++++++++++"
          }
          else {
               Remove-Item  "\\$computer\c$\Users\*\Desktop\$ShortcutName" -Force
          }
     }
}
