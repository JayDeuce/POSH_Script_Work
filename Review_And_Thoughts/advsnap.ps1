#requires -Version 4

Function Start-PSNap {

<#
.SYNOPSIS
Start a PowerShell napping session.
.DESCRIPTION
Use this command to start a short napping session. The command will alert you when you nap is up with a chime and a message. You have an option of displaying the message on the screen or having it spoken.
.PARAMETER Minutes
The number of minutes for your nap. This command has aliases of: nap,time
.PARAMETER ProgressBar
Indicate if you want to show a progress bar which includes a number of messages.
.PARAMETER Message
The text of the message to be displayed or spoken at the end of your nap.
.PARAMETER Voice
Specify the name of the installed voice to use. More most US desktops this will be David, Zira and perhaps Hazel. If you use this parameter the message will not be written to host.
.PARAMETER Rate
The voice speaking rate 
.EXAMPLE
PS C:\> Start-PSNap 10 -ProgressBar -message "Get back to work you lazy bum!"
Start a 10 minute nap with the progress bar and display the given message in the console host.
.EXAMPLE
PS C:\> Start-PSNap 15 -message "Wake up you fool and get back to work." -voice Zira
Start a 15 minute nap and use the computer voice Zira to speak the wake up message.
.NOTES
NAME        :  Start-PSNap
VERSION     :  1.0   
LAST UPDATED:  29 January 2016
AUTHOR      :  Jeff Hicks (@jeffhicks)

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************
.LINK
Start-Sleep
.INPUTS
none
.OUTPUTS
none
#>

[cmdletbinding(DefaultParameterSetName = "host")]
Param(
[Parameter(Position = 0,ParameterSetName = "voice")]
[Parameter(Position = 0,ParameterSetName = "host")]
[Alias("nap","time")]
[ValidateRange(1,30)]
[int]$Minutes = 1,

[Parameter(ParameterSetName = "voice")]
[Parameter(ParameterSetName = "host")]
[switch]$ProgressBar,

[Parameter(ParameterSetName = "voice")]
[Parameter(ParameterSetName = "host")]
[ValidateNotNullorEmpty()]
[string]$Message = "Get back to work sleepy head!",

[ValidateScript({
 #get voices
 Add-Type -AssemblyName System.speech
 $installed = [System.Speech.Synthesis.SpeechSynthesizer]::new().GetInstalledVoices().voiceinfo.Name
 if ($installed -match $_ ) {
    $True
 }
 else {
    [regex]$rx= "Microsoft\s+(?<name>\w+)\s+"
    #build a list of voices assuming the voice name is something like Microsoft David Desktop
    $choices = (($rx.Matches($installed)).foreach({$_.groups["name"].value})) -join ","
    Throw "Can't find an installed voice for $_. Possible values are: $choices"
 }

})]
[Parameter(ParameterSetName = "voice")]
[string]$Voice,

[Parameter(ParameterSetName = "voice")]
[ValidateRange(-5,5)]
[int]$Rate = 1
)

$wake = (Get-Date).AddMinutes($Minutes)
$remainingSeconds = $minutes * 60

#an array of status messages if using a progress bar
$ops = "I'm solving a PowerShell problem","I'm chasing cmdlets",
"Brilliance at work","Re-initializing my pipeline",
"Go away","I'm checking eyelid integrity","It can wait...",
"Don't you dare!","Spawning a new runspace","I'm multitasking",
"Nothing is that important","Unless you have beer for me, go away",
"I'm testing the new PSNap provider","I need this",
"I'm downloading my new matrix","Resource recyling in progress",
"Synaptic synch in progress","Neural network rebooting",
"If you can read this you shouldn't be here",
"$($env:username) has left the building"

#hashtable of parameter values to splat to Write-Progress
$progHash = @{
 Activity = "Ssssshhh..."
 Status = $ops[0]
 SecondsRemaining = $remainingSeconds
}

cls
#loop until the time is >= the wake up time
do {

if ($ProgressBar ) {
    Write-Progress @proghash
    #tick down $remainingseconds
    $proghash.SecondsRemaining = $remainingSeconds--
    #pick a new random status if remaining seconds is divisible by 10
    if ($remainingSeconds/10 -is [int]) {
        $proghash.status = $ops | Get-Random
    }
} #if
else {
    cls
    Write-Host "Ssshhhh...."

    #trim off the milliseconds
    Write-Host ($wake - (Get-Date)).ToString().Substring(0,8) -NoNewline
} #else

Start-Sleep -Seconds 1

} Until ( (Get-Date) -ge $wake)

#Play wake up music
[console]::Beep(392,1000)
[console]::Beep((329.6*2),1000)
[console]::Beep(523.2,1000)

If ($Voice) {
    Add-Type -AssemblyName System.speech
    $speech = New-Object System.Speech.Synthesis.SpeechSynthesizer
    #find the matching voice object
    $selected = [System.Speech.Synthesis.SpeechSynthesizer]::new().GetInstalledVoices().voiceinfo.Name | where {$_ -match $voice}
    $speech.SelectVoice($selected)
    $speech.Rate = $Rate
    $speech.SpeakAsync($message) | Out-Null
    #write a blank line to get a new prompt
    Write-Host "`n"
}
else {
    Write-Host "`n$Message" -ForegroundColor Yellow
}

} #end function