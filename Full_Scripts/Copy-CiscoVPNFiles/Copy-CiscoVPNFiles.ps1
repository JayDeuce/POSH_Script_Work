[cmdletbinding()]

Param (
     [Parameter(mandatory = $true, Position = 0)]
     [array]$computers = "",
     [Parameter(mandatory = $false, Position = 1)]
     [string]$CiscoShortcutName = "Cisco AnyConnect Secure Mobility Client.lnk",
     [Parameter(mandatory = $false, Position = 2)]
     [string]$DMShortcutName = "Drivemapper.lnk",
     [Parameter(mandatory = $false, Position = 3)]
     [string]$DMFileName = "drivemapper.exe",
     [Parameter(mandatory = $false, Position = 4)]
     [string]$VPNFilesPath = "\\server\Copy-CiscoVPNFiles\VPN_Files"
)
process {
     foreach ($computer in $computers) {
          Write-Host "`nProcessing $computer..."
          if (!(test-connection -count 1 -Quiet -ComputerName $computer)) {
               Write-Host "`n$computer could not be contacted, please check make sure its online. Moving on...`n"
               Write-Host "+++++++++++++++++++"
          }
          else {
               Copy-Item $VPNFilesPath\$CiscoShortcutName "\\$computer\c$\Users\Public\Desktop"
               Copy-Item $VPNFilesPath\$DMShortcutName "\\$computer\c$\Users\Public\Desktop"
               Copy-Item $VPNFilesPath\$DMFileName "\\$computer\c$\Program Files (x86)"
          }
     }
}
