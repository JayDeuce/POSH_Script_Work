Param (
     [parameter(Mandatory=$true)][string]$ComputerName = ""
)

get-wmiobject -namespace root\CIMv2\Security\MicrosoftVolumeEncryption -class Win32_EncryptableVolume -computername $ComputerName