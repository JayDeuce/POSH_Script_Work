<#
.SYNOPSIS
     Check Bitlocker Status

.DESCRIPTION
     Checks Bitlockers status for hard drives on given computername

.PARAMETER  ComputerName
     (Required)

     Description:
          Name of the computer you want to check

.EXAMPLE
     .\Get-BitlockerStatus.ps1 -ComputerName "TommysPC"

     Description:
          Gets the hard drive bitlocker status on Tommy's PC

.NOTES
     Name: Get-BitlockerStatus
     Author: Jonathan Durant
     Version: 1.0
     DateUpdated: 14 June 2017

.INPUTS
     NONE

.OUTPUTS
     NONE
#>

[cmdletbinding()]

Param (
     [parameter(Mandatory = $true)]
     [string]$ComputerName = ""
)

Begin {

}
Process {
     Get-WmiObject -namespace root\CIMv2\Security\MicrosoftVolumeEncryption -class Win32_EncryptableVolume -computername $ComputerName
}
End {

}