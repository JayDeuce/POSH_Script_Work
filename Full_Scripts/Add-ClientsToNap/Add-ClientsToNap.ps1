<#
.SYNOPSIS
    Creates new local user accounts on the local NAP server and adds the new user to a specific vlan group; reads info from a CSV file.

.DESCRIPTION
   Creates new local user accounts on the local NAP server and adds the new user to a specific vlan group; reads info from a CSV file.
   The script assumes a 14 character Username (MAC Address), in Format XXXX.XXXX.XXXX. It will check for a 14 character Username, if found
   it will continue; if not it will check to see if it is 12 (MAC Address without any puncuation are 12 chars long) and insert the periods
   in the proper location and then continue with creation. If the Username is blank, or any other length than 12 or 14 it will error for that
   system.

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  systemList
     (Required, No Default)

     This parameter is the name of the csv where the data needed to create accounts is located. It should contain 3 columns of
     data, one for Username, one for Description, and one for Vlan Number. One set of data per line.

.EXAMPLE
    .\Add-ClientsToNap.ps1 -systemList 'List.csv"

    Description:

        Works line by line through the LIST.CSV file and creates the local user adding them to the proper vlan group as listed.

.NOTES
    Name: Add-ClientsToNap.ps1
    Author: Jonathan Durant
    Version: 1.1
    DateUpdated: 2017-10-18

.INPUTS
    CSV file located in same location as the script containing Username, Description, Vlan Columns and data; one set per line.

.OUTPUTS
    Status messages to host
#>

[cmdletbinding()]

Param (
     [Parameter(mandatory = $true, Position = 0)]
     [array]$systemList = ""
)

function format-name($nItem) {
     if ($nItem.Length -ne '12' -and $nItem.Length -ne '14') {
          Write-Host "'$nItem' is $($nitem.Length) Characters Long, must be only 12 chars 'xxxxxxxxxxxx' or 14 chars (including periods) 'xxxx.xxxx.xxxx'. Please review list.`n"
          $Script:format = $false
     }
     else {
          if ($nItem.Length -eq '12') {
               $Script:name = $nItem.Insert(4, '.').Insert(9, '.')
          }
     }
}

try {
     $list = import-csv $systemList -ErrorAction Stop | select-object Username, Description, Vlan
     $pass = ConvertTo-SecureString "<ADDLOCALPASSWORD>" -AsPlainText -Force -ErrorAction stop

     foreach ($item in $list) {
          $format = $true
          $name = $item.username
          $descrip = $item.description
          $vlan = $item.vlan
          if ($name -ne $null -and $name -ne '') {
               format-name($name)
               if ($format -eq $true) {
                    if (Get-WmiObject Win32_UserAccount -Filter "LocalAccount='True' and Name='$name'") {
                         Write-Host "'$name' is already a client on the NAP Server, Please check to ensure you have the right MAC or Delete the account.`n"
                    }
                    Else {
                         write-host "Creating $name"
                         New-LocalUser -Name $name -FullName $name -Description $descrip -Password $pass -PasswordNeverExpires -UserMayNotChangePassword | Out-Null
                         Add-LocalGroupMember -Group "Vlan $vlan" -Member $name
                         Add-LocalGroupMember -Group "Users" -Member $name
                    }
               }
          }
          else {
               Write-Host "There is an Empty UserName Field! Please reveiw CSV, all 'UserName' fields must have an entry!`n" -ForegroundColor Red
          }
     }
}
Catch {
     if ($_.exception.message -like "*Could not find file*") {
          Write-Host "`nCSV File not found! `n`nException Message:" $_.exception.message -ForegroundColor  Red
     }
}
