<#
.SYNOPSIS
    Gets the Network Card information for all computers listed in the computers.txt file and outputs them to a CSV file.

.DESCRIPTION
    Using the list of computernames/IPAddresses in the computers.txt file, the script will queries each computer and output the Computername, IP Address,
    MAC Address, and DefaultGateway to a CSV file for ease of veiwing.

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.NOTES
    Name: Get-IPMAC.ps1
    Author: Jonathan Durant
    Version: 1.0
    DateUpdated: 2016-01-20

.EXAMPLE
    .\Get-IPMAC.ps1

    Description:

        Loads the computers list from the computers.txt file and gets required info and creates the List.CSV file
        with all info for each machine

.INPUTS
    Computers.txt file located in same location as the script containing computer names or IP Address; one per line.

.OUTPUTS
    List.CSV file in the C:\TEMP Folder
#>


[cmdletbinding()]

Param (
)

$Computers = (Get-Content "Computers.txt") # Get list of Computers
$AllList = @() #Set empty PSObject to output to CSV

# Cycle through each line of the computers.txt list and run WMI queries to get NIC Info and input to PS Object
foreach ($Computer in $Computers) {

     # Create Object to hold all NIC Info
     $List = New-Object PSObject
     # Add Computer name to object for the current object
     $List | Add-Member -Name Computername -Value $Computer -MemberType NoteProperty

     if (test-connection -count 1 -Quiet -ComputerName $Computer) {

          # Get namespace and filter only by which adapters have an IP address enabled
          $colItems = Get-WmiObject -Class "Win32_NetworkAdapterConfiguration" -ComputerName $Computer -Filter "IpEnabled = TRUE" -ErrorAction SilentlyContinue


          # For the current object get the NIC Info and pass to Created $List Object
          ForEach ($objItem in $colItems) {
               $List | Add-Member -Name 'NIC Description' -Value $objItem.Description -MemberType NoteProperty
               $List | Add-Member -Name 'MAC Address' -Value $objItem.MacAddress -MemberType NoteProperty
               $List | Add-Member -Name 'IP Address' -Value $objItem.IPAddress[0] -MemberType NoteProperty
               $List | Add-Member -Name 'Default Gateway' -Value $objItem.DefaultIPGateway[0] -MemberType NoteProperty
          }
     }
     else {
          $List | Add-Member -Name 'NIC Description' -Value ':----- COMPUTER OFFLINE ----:' -MemberType NoteProperty
     }

     # Add current object retrieved info to the $AllList Object
     $AllList += $List
}
$AllList | Export-Csv -Path 'C:\Temp\List.csv' -NoTypeInformation
#$AllList | Out-GridView