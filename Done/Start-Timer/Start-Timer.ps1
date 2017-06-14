<#
.SYNOPSIS
     Counts down the given minutes and beeps on 00:00 as an alarm.

.DESCRIPTION
     Using the passed parameter for a by the Minute Countdown, the script will countdown by the second
     and then beep the computers speakers to alarm you when the time is up.

.PARAMETER  countdownTime
     (Required, No Default)

     This parameter is the amount of minutes you want the timer to countdown, in whole numbers no decimals.

     ex: 10 or 1

.EXAMPLE
     .\Start-Timer.ps1 -countdown 15

     Description:

          Starts a time to countdown 15 minutes and sets off the alarm sounds when finished.

.NOTES
     Name: Start-Timer.ps1
     Author: Jonathan Durant
     Version: 1.0 (Based on PowerShellNap.ps1 by Jeff Hicks at
                    http://jdhitsolutions.com/blog/powershell/4830/friday-fun-a-powershell-nap/)
     DateUpdated: 2016-01-26

.INPUTS
     Number Time in Minutes

.OUTPUTS
     Host message and beep tones.
#>

[cmdletbinding()]

# Countdown time in minutes
Param (
     [Parameter(mandatory = $true, Position = 0)]
     [int]$countdownTime = 1
)

Process {
     try {
          # Set Endtime
          $endTime = (Get-Date).AddMinutes($countdownTime)

          # Loop until the time is >= the end time
          do {
               Clear-Host
               Write-host "Timer-Started....`n" -ForegroundColor Green

               # Trim off the milliseconds
               write-host ($endTime - (Get-Date)).ToString().Substring(0, 8) -NoNewline -ForegroundColor Cyan

               Start-Sleep -Seconds 1

          } Until ((Get-Date) -ge $endTime)

          # Play wake up sounds
          [console]::Beep(392, 1000)
          [console]::Beep((329.6 * 2), 1000)
          [console]::Beep(523.2, 1000)

          Write-Host "`n`nCountdown Over!" -ForegroundColor Yellow
     }
     catch {

     }
}