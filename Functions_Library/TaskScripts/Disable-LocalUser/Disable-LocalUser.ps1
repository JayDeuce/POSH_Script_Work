﻿# specify local user name.
$userName = "user"

# Read server NetBIOS names from csv file.
$computers =  Get-Content .\list.txt

ForEach ($computer In $computers) {
     try {
          Write-Host "Processing $computer..."
          # Connect to local user on remote server.
          $User = [ADSI]"WinNT://$computer/$UserName,user"
          # Disable the user.
          $User.AccountDisabled = $True
          $User.SetInfo()
          Write-Host "...$computer Complete."
          Write-Host "****************************"
     }
     catch {
          "Unable to connect to $UserName on $computer"
          Write-Host "****************************"
     }
}