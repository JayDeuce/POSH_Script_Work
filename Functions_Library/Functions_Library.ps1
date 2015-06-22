# ====================================================================================
# Pulled from autospsourcebuilder, Created by Brian Lalancette (@brianlala)
# http://autospsourcebuilder.codeplex.com
# Remove Read-Only Attribute
# ====================================================================================

Function Remove-ReadOnlyAttribute ($Path)
{
    ForEach ($item in (Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue))
    {
        $attributes = @((Get-ItemProperty -Path $item.FullName).Attributes)
        If ($attributes -match "ReadOnly")
        {
            # Set the file to just have the 'Archive' attribute
            Write-Host "  - Removing Read-Only attribute from file: $item"
            Set-ItemProperty -Path $item.FullName -Name Attributes -Value "Archive"
        }
    }
}

#------------------

# ===================================================================================
# Pulled from autospsourcebuilder, Created by Brian Lalancette (@brianlala)
# http://autospsourcebuilder.codeplex.com
# Func: Pause
# Desc: Wait for user to press a key - normally used after an error has occured or input is required
# ===================================================================================
Function Pause($action, $key)
{
    # From http://www.microsoft.com/technet/scriptcenter/resources/pstips/jan08/pstip0118.mspx
    if ($key -eq "any" -or ([string]::IsNullOrEmpty($key)))
    {
        $actionString = "Press any key to $action..."
        Write-Host $actionString
        $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    else
    {
        $actionString = "Enter `"$key`" to $action"
        $continue = Read-Host -Prompt $actionString
        if ($continue -ne $key) {pause $action $key}

    }
}

#-----------------

# ====================================================================================
# First check if we are running this under an elevated session. 
# Pulled from the script at http://gallery.technet.microsoft.com/scriptcenter/1b5df952-9e10-470f-ad7c-dc2bdc2ac946
# ====================================================================================

If (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning " - You must run this script under an elevated PowerShell prompt. Launch an elevated PowerShell prompt by right-clicking the PowerShell shortcut and selecting `"Run as Administrator`"."
    break
}

#------------------

# ====================================================================================
# Pulled from autospsourcebuilder, Created by Brian Lalancette (@brianlala)
# http://autospsourcebuilder.codeplex.com
# Func: EnsureFolder
# Desc: Checks for the existence and validity of a given path, and attempts to create if it doesn't exist.
# From: Modified from patch 9833 at http://autospinstaller.codeplex.com/SourceControl/list/patches by user timiun
# ====================================================================================

Function EnsureFolder ($Path)
{
	If (!(Test-Path -Path $Path -PathType Container))
	{
		Write-Host -ForegroundColor White " - $Path doesn't exist; creating..."
		Try
		{
			New-Item -Path $Path -ItemType Directory | Out-Null
		}
		Catch
		{
			Write-Warning " - $($_.Exception.Message)"
			Throw " - Could not create folder $Path!"
            $errorWarning = $true
		}
	}
}

#-----------------

# ====================================================================================
# Check is a folder path is therr, if not create it.
# ====================================================================================
function checkfolderPath([string]$FolderPath)
{
	if (!(Test-Path -Path $FolderPath)) 
	{
		new-item -Path $FolderPath -ItemType directory | Out-Null
	}
}

#-----------------

# ====================================================================================
# Build a FileName out of Now Date and Time. Vatiable Build
# ====================================================================================

[string]$Date = [dateTime]::now.toshortdatestring().Replace("/","-")
[string]$Time = [dateTime]::now.toshorttimestring().Replace(":","-").Replace(" ","-")
[string]$DateTime = $Date + "_" + $Time

# OR

Get-Date -Format "dd-mm-2014_HH-mm"

#-----------------

# ====================================================================================
# Start ConsoleTrancript Function - Used in Profile Scripts
# ====================================================================================

function Start-ConsoleTranscript()
{
	# Check if I want to start a Transcript, Used if I am trying out commands
	[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
	[string]$TranYN = [Microsoft.VisualBasic.Interaction]::MsgBox("Would you Like to Start a Transcript?", "YesNo, DefaultButton2",  "Start Transcript?")
	# Check the answer for "Yes" or "No" and act upon the choice
	switch ($TranYN) 
	{
		Yes 
		{
			[string]$Date = [dateTime]::now.toshortdatestring().Replace("/","-")
			[string]$Time = [dateTime]::now.toshorttimestring().Replace(":","-").Replace(" ","-")
			[string]$DateTime = $Date + "_" + $Time
			
			# Check for my Standard Scripts Folder, and act accordingly to its presence or not
			If (Test-Path -Path "scripts:\")
			{	
				start-transcript -path "Scripts:\Transcripts\$DateTime.txt" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
			}
			Else
			{
				# If Standard Scripts Path is not available, ask where to save
				[string]$LogPath = [Microsoft.VisualBasic.Interaction]::InputBox("Where would you like to save the transcript file?", "Trascript Path")							
				If ($LogPath -eq "") # Error handling is answer is blank or I hit cancel
				{
					Write-Warning "HEY!!!, No transcript save file location set, defaulting to Transcript Off!!"
					Write-host "`n"
					[Microsoft.VisualBasic.Interaction]::MsgBox("Transcript is off!", "OKOnly, Exclamation",  "Transcript Status") | Out-Null
				}
				else
				{
					$LogPath = $LogPath.TrimEnd("\")
					checkfolderPath($LogPath)
					start-transcript -path "$LogPath\$DateTime.txt" | Write-Host -BackgroundColor "darkgreen" -ForegroundColor "yellow"
					Write-host "`n"
				}
			}
		}
		No 
		{
			Write-host "Transcripting is turned off for this console session." -BackgroundColor "darkgreen" -ForegroundColor "yellow"
		}
	}	
}

#-----------------

# ====================================================================================
#  Email Variable Creation and Code to Send
# ====================================================================================

$Recipients = '"User1 <user1@web.com>"','"User1 <user1@web.com>"','"User1 <user1@web.com>"'
$Sender = '"Sender1 <Sender1@web.com>"'
$SMTPServer = "000.000.000.000 <ServerAddress>"
$Subject = "Subject"
$Body = @"
Good Morning,

BODY OF THE EMAIL

"@

#Build the Notification Email and send it to the Recipients
Send-MailMessage -To $Recipients -From $Sender -Subject $Subject -Body $Body -SmtpServer $SMTPServer