<#
.SYNOPSIS
    Sends Messages via pop-up dialog boxes to the listed locations

.DESCRIPTION
    Using passed parameters for the sendTo location and message; this script
    will send a message to the user via pop-up message box.

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  sendTo
    (Required, No Default)

    This parameter is the location of computers you want to send the message to. The 
    following options that can be used to send message to multiple computers in Specific locations:

    "Lan1"      -> Send to all Computers located on LAN1
    "Lan2"     -> Send to all Computers located on LAN2
    "All"      -> Send to all Computers located on all LANS
                                
                                OR
                                 
    Individual computers can also be passed using Active Directory Computername or IP Address.

        EX: "Workstation1"  
            "192.168.111.111"  

.PARAMETER message
    (Required, No Default)

    This parameter is the message you want to send to the user that will appear in the popup window.

    EX: "This is a message to the user!"

.EXAMPLE
    .\Send-Message -sendTo Workstation1 -message "Hello World!"

    Description:

        Sends the message "Hello World!" to the computer Workstation1 in a popup window.

.EXAMPLE
    .\Send-Message -sendTo LAN1 -message "Hello World!"

    Description:

        Sends the message "Hello World!" to the all computers in LAN1 in a popup window.

.EXAMPLE
    .\Send-Message -sendTo All -message "Hello World!"

    Description:

        Sends the message "Hello World!" to the all computers in the LAN Computer network boundaries in a popup window.

.NOTES
    Name: Send-Message.ps1
    Author: Jonathan Durant
    Version: 1.0
    DateUpdated: 2016-02-11      

.INPUTS
    Computer names, message string
.OUTPUTS
    NONE
#>

[cmdletbinding()]

Param (
    [Parameter(mandatory=$true,Position=0)]
        [string]$sendTo = "",
    [Parameter(mandatory=$true,Position=1)]
        [string]$message = ""
)

begin {
    # Checks the enter parameter of $sendTo and sets the $computers variable to the computer names or ip lists it pulls based on the parameter entered
    [array]$computers = ""
    switch($sendTo) {
        Fig {
                $computers = (Get-Content -Path ".\ip_ranges\Lan1.txt")
            }
        Lead {
                $computers = (Get-Content -Path ".\ip_ranges\lan2.txt")
            }
        All {
                $computers = (Get-Content -Path ".\ip_ranges\all.txt")
            }
        default {
                $computers = $sendTo
            }
    }
}
process {
    foreach ($computer in $computers) {
        # Test the connectivity of each connection and it live sends the message to the computer, if not live just quits
        if (test-connection -count 1 -Quiet -ComputerName $computer) {    
            Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * /time:86400 $message" -ComputerName $computer | Out-Null
        }
    }
}
