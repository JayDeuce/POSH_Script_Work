<#
.SYNOPSIS
    Creates a Message Pop Up Box with passed Parameters

.DESCRIPTION
    Using passed parameters for the message and title; this script
    will create a pop-up Box on the local computer.

.PARAMETER  message
    (Required, No Default)

    This Parametewr is the message the popup will show

.PARAMETER  title
    (Required, No Default)

    This parameter is title of the popup box

.EXAMPLE
    .\Create-MessageBox -message "Hello World -title "ATTENTION"

    Description:

        Pop-ups a Message box that says "Hello World" with a Window title of "ATTENTION"

.NOTES
    Name: Create-MessageBox
    Author: Jonathan Durant
    Version: 1.0
    DateUpdated: 2016-06/06        

.INPUTS
    Message text, Title Text

.OUTPUTS
    Popup Box
#>
[cmdletbinding()]

Param (
    [Parameter(Mandatory=$true)][string]$message,
    [Parameter(Mandatory=$true)][string]$title    
)      

Process {
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

    [System.Windows.Forms.MessageBoxButtons]$buttons="OK"
    [System.Windows.Forms.MessageBoxIcon]$icon="Information"

    [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
}