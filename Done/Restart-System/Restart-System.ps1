<#
.SYNOPSIS
    Restart the target computer using the 10 minute delay and user warning message.

.DESCRIPTION
    Using passed parameters for the computer name and user warning message; this script
    will send a message to the user and restart after 10 minutes have passed.

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  computers
    (Required, No Default)

    This parameter is the active directory/workstation name or IP Address of the Computer that will receive the update. 
    This designation can be a single Computer name, IP Address, or a function to read in a listing of Computers names, such as

    EX: "AMEDDAHCNB16576"  
        "160.151.111.111"  
        "(Get-Content "Computers.txt")" //File is located in same directory as the script//
        "(Get-Content "C:\Temp\List\Computers.txt")"

.PARAMETER  userMsg
    (Required, Defaulted)

    This parameter is the message you want to send to the user in the restart warning window.

    EX: "Computer needs to restart in 10 minutes. Please save your work." //This is the default message//

.PARAMETER  delayTime
    (Required, Defaulted)

    This parameter is the time (IN SECONDS) you want the system to delay until it restarts. NOTE: A pop-up message 
    will inform the logged on user at 10 minutes and 2 minutes to shutdown.

    EX: 600 //This is the default delay (IN SECONDS) set to 10 minutes//

.EXAMPLE
    .\Restart-System -computers AMEDDAHCWK00337

    Description:

        Restarts the computer named "AMEDDAHCWK19278" in the default 10 minutes with the default
        user message

.EXAMPLE
    .\Restart-System -computers AMEDDAHCWK00337 -userMsg "Restarting the computer.." -delayTime 21600

    Description:

        Restarts the computer named "AMEDDAHCWK19278" using a custom message in 6 hours (21600 seconds).

.EXAMPLE
    .\Restart-System -computers (Get-Content "computers.txt") -userMsg "Restarting the computer.."

    Description:

        Restarts the computers listed in the computers.txt file (Which is in the same folder location 
        as the Restart-System.ps1 file) in 10 minutes using a custom message.

.EXAMPLE
    .\Restart-System -computers (Get-Content "c:\temp\computers.txt) -userMsg "Restarting the computer.." -delayTime 900

    Description:

        Restarts the computers listed in the computers.txt file located in the temp folder in 15 minutes using a 
        custom message.

.NOTES
    Name: Restart-System
    Author: Jonathan Durant
    Version: 1.1
    DateUpdated: 2016-01-28      

.INPUTS
    Computer names, message string, number
.OUTPUTS
    NONE
#>

[cmdletbinding()]

Param (
[Parameter(mandatory=$true,Position=0)]
    [array]$computers = "",
    [Parameter(mandatory=$false,Position=1)]
    [string]$userMsg = "Computer needs to restart in 10 minutes. Please save your work.",
    [Parameter(mandatory=$false,Position=2)]
    [int]$delayTime = 600
)
process {    
    foreach ($computer in $computers) {
        & 'C:\Windows\System32\SHUTDOWN.exe' -m \\$computer -r -f -c $userMsg -t $delayTime
        write-host "`n$computer has started shutdown`n"
    }
}
