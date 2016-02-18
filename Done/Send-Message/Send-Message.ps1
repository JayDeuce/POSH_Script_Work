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

    "Fig"      -> Send to all Computers located on Ft. Indiantown Gap clinic LAN
    "Lead"     -> Send to all Computers located on Letterkeny LAN
    "Fillmore" -> Send to all COmputers located on the Fillmore LAN
    "Dunham"   -> Send to all Computers located on the Dunham Main Clinic LAN
    "All"      -> Send to all Computers located on all Dunham Controlled LANS
                                
                                OR
                                 
    Individual computers can also be passed using Active Directory Computername or IP Address.

        EX: "AMEDDAHCNB16576"  
            "160.151.111.111"  

.PARAMETER message
    (Required, No Default)

    This parameter is the message you want to send to the user that will appear in the popup window.

    EX: "This is a message to the user!"

.EXAMPLE
    .\Send-Message -sendTo AMEDDAHCWK002322 -message "Hello World!"

    Description:

        Sends the message "Hello World!" to the computer AMEDDAHCWK002322 in a popup window.

.EXAMPLE
    .\Send-Message -sendTo FIG -message "Hello World!"

    Description:

        Sends the message "Hello World!" to the all computers in Ft. Indiantown Gap in a popup window.

.EXAMPLE
    .\Send-Message -sendTo All -message "Hello World!"

    Description:

        Sends the message "Hello World!" to the all computers in the DAHC Computer network boundaries in a popup window.

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
                $computers = (Get-Content -Path "\\ameddahcfs02\imd$\~SysAdminFiles\POSH_Scripts\Send-Message\ip_ranges\fig.txt")
            }
        Lead {
                $computers = (Get-Content -Path "\\ameddahcfs02\imd$\~SysAdminFiles\POSH_Scripts\Send-Message\ip_ranges\lead.txt")
            }
        Fillmore {
                $computers = (Get-Content -Path "\\ameddahcfs02\imd$\~SysAdminFiles\POSH_Scripts\Send-Message\ip_ranges\fillmore.txt")
            }
        Dunham {
                $computers = (Get-Content -Path "\\ameddahcfs02\imd$\~SysAdminFiles\POSH_Scripts\Send-Message\ip_ranges\dunham.txt")
            }
        All {
                $computers = (Get-Content -Path "\\ameddahcfs02\imd$\~SysAdminFiles\POSH_Scripts\Send-Message\ip_ranges\all.txt")
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
            Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $message" -ComputerName $computer | Out-Null
        }
    }
}
