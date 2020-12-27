# Set Window Size
$psHost = Get-Host
$psWindow = $psHost.UI.RawUI
# Set Windows Size
$newSize = $psWindow.WindowSize
$newSize.height = 5
$newSize.width = 50
$psWindow.WindowSize = $newSize
# Set Window Title
$title = "Unlock Bitlocker Drives"
$psHost.UI.RawUI.WindowTitle = $title

#Set Read-Host a few lines off the top of window
Write-Host "`n"

# Request Key
$key = Read-Host "Enter Bitlocker Key" -AsSecureString

# Get All BL Encrypted Drives
$drives = Get-BitLockerVolume | Where-Object -Property "CapacityGB" -EQ "0.00"

# Unlock Each Drive
foreach ($drive in $drives) {
    Unlock-BitLocker -MountPoint $drive -Password $key | Out-Null
}