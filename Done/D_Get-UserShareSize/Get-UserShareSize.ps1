<#
	!! NOTE: Currently Configured for use as a Scheduled task on USANATO User shares !!
	!! Use Standard Version For manual use.											 !!

.SYNOPSIS
    Get User Share folder sizes in specified tree.  
.DESCRIPTION
    Script creates an HTML report with Folder Name, Folder size, 
	and Count of ACL Groups.  By default script will only do 1
    level of folders.  Add Recurse to Main Function Calls in Main Script 
	to do all sub-folders.
	
	Parameter info is used to setup Variables and Function Calls in Main Script
.PARAMETER Path
    Specify the path you wish to report on
.PARAMETER ReportPath
    Specify where you want the HTML report to be saved
.PARAMETER Recurse
    Report on all sub-folders
.EXAMPLE
    .\Get-UserShareSize
    Run script and use defaults
.OUTPUTS
    1. "ReportPathVariable".HTML in specified Report Path
	2. Notification email sent to the Recipients List
#>

# Setup Script Scope Variables
$Report = @()
$TotSize = 0
$Date = Get-Date -Format "yyyyMMdd-HHmm"
# Paths Hard Coded for use in Task Scheduler
$PathAFN = "\\ktes2292apcs280\usanato_afnorth_users"
$PathAFS = "\\ktes2292apcs280\usanato_afsouth_users"
$PathBDE = "\\ktes2292apcs280\usanato_bdehq_users"
$AFNReport = "\\ktes2292apcs280\USANATO_TECHNICAL\srvr_reports\UserShareMonitoring\AFN_UFS-$Date.html"
$AFSReport	= "\\ktes2292apcs280\USANATO_TECHNICAL\srvr_reports\UserShareMonitoring\AFS_UFS-$Date.html"
$BDEReport	= "\\ktes2292apcs280\USANATO_TECHNICAL\srvr_reports\UserShareMonitoring\BDEHQ_UFS-$Date.html"
# Build Notification Email for Send-MailMessage Command
$Recipients = '"USARMY Sembach HQ USANATO Bde List G6 Admin Group <usarmy.sembach.hq-usanato-bde.list.g6-admin-group@mail.mil>"','"USARMY SHAPE HQ USANATO Bde List AFNORTH Bn S6 <usarmy.shape.hq-usanato-bde.list.afnorth-bn-s6@mail.mil>"','"USARMY Vicenza HQ USANATO Bde List HQUSANATO AFSOUTH S6 <usarmy.vicenza.hq-usanato-bde.list.hqusanato-afsouth-s6@mail.mil>"'
$Sender = '"BDE User Share Monitoring Script <noreply@bdenoreply.mil>"'
$SMTPServer = "136.215.66.134"
$Subject = "New User Share Size Monitoring Report Created - $Date"
$Body = @"
Good Morning,

A New Report on the size of the BDE User Shares has been created and saved to:

	\\ktes2292apcs280\USANATO_TECHNICAL\srvr_reports\UserShareMonitoring

Please look over the report for your AOR and ensure all users are compliant.

Any question or issues please contact the USANATO BDE KMO Officer at (314) 549-5425

Thank you and have a great day!

	Your Friendly Neighborhood User Share Montoring Script

-----------------------------------------------------------------------------------------------------
NOTE: PLEASE DO NOT REPLY TO THIS MESSAGE, THERE IS NO MAILBOX AS IT IS FOR INFORMATION PURPOSES ONLY
-----------------------------------------------------------------------------------------------------
"@

# Function Gather info on the Current Path, and set Report Object Variable Properties
Function AddObject {
    Param ( 
        $FileObject
    )
	# Gathers Sizes info on Current File Path Obejct, send it to CalculateSize function to convert it to 2 decimal string
    $Size = [double]($FSO.GetFolder($FileObject.FullName).Size)
    $Script:TotSize += $Size
    $Size = CalculateSize $Size
	
    # Gathers info on Access Control List Groups of the Current file path object and Set proper variables for use in the Report Object
    #$Owner = [string]((get-acl $FileObject.FullName).Access | ForEach-Object { $_.identityReference.value })
	$OwnerCount = ((get-acl $FileObject.FullName).Access).Count
	
	#Builds the Report Object in the Script Scope Report Variable
    $Script:Report += New-Object PSObject -Property @{
        'Folder Name' = $FileObject.Name
        'File Size' = $Size
        #'Assigned Groups' = $Owner
		'Groups Count' = $OwnerCount
    }
}

#Function Recieves a Raw Folder Size Number, and Converts it to a 2 decimal String with the Proper Byte identifier based on size
Function CalculateSize {
    Param (
        [double]$Size   
    )
    If ($Size -gt 1000000000)
    {   
        $ReturnSize = "{0:N2} GB" -f ($Size / 1GB)
    }
    ElseIf ($Size -ge 1000000 -And $Size -le 1000000000)
    {   
        $ReturnSize = "{0:N2} MB" -f ($Size / 1MB)
    }
    ElseIf ($Size -ge 1000 -And $Size -le 1000000)
    {
        $ReturnSize = "{0:N2} KB" -f ($size / 1KB)
    }
    Else
    {
        $ReturnSize = "{0:N0} Bytes" -f ($size)
    }
    Return $ReturnSize
}

# Function to Set Alternating row colors in the table listing when using HTML
Function Set-AlternatingRows {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [object[]]$Lines,
       
        [Parameter(Mandatory=$True)]
        [string]$CSSEvenClass,
       
        [Parameter(Mandatory=$True)]
        [string]$CSSOddClass
    )
    Begin {
        $ClassName = $CSSEvenClass
    }
    Process {
        ForEach ($Line in $Lines)
        {   $Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
            If ($ClassName -eq $CSSEvenClass)
            {   $ClassName = $CSSOddClass
            }
            Else
            {   $ClassName = $CSSEvenClass
            }
            Return $Line
        }
    }
}

# Main function that get the Folder sizes that is sent from the Root Path, calls other functions to complete job
function getFolderSize () {
	Param (
		[string]$Path = "",
		[string]$ReportPath = "",
   		[switch]$Recurse
	)
	
	# Build a FileSystemObject
	$FSO = New-Object -ComObject Scripting.FileSystemObject 

	# Get the properties of the starting path
	$Root = Get-Item -Path $Path 
	AddObject $Root

	# Loop through all the subfolders useing hte Build Parameter list of the Path and to Recurse or Not
	$ParamList = @{
    	Path = $Path
    	Recurse = $Recurse
	}
	ForEach ($Folder in (Get-ChildItem @ParamList | Where-Object { $_.PSisContainer }))
	{   
    	AddObject $Folder
	}
	
	$TotSize = CalculateSize $TotSize
	
	# Build HTML Header, Pre, and Post Code for report including embedded CSS
	$Header = @"
	<style>
	TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 6px;border-style: solid;border-color: black;background-color: #6495ED;}
	TD {border-width: 1px;padding: 6px;border-style: solid;border-color: black;}
	TABLE TR TH:nth-child(1) {width: 200px; word-break:break-all;}
	TABLE TR TD:nth-child(1) {width: 200px; word-break:break-all;}
	TABLE TR TH:nth-child(2) {width: 200px;text-align: center;}
	TABLE TR TD:nth-child(2) {width: 200px;text-align: center;}
	TABLE TR TH:nth-child(3) {width: 200px; text-align: center;}
	TABLE TR TD:nth-child(3) {width: 200px; text-align: center;}
	.odd  { background-color:#ffffff; }
	.even { background-color:#dddddd; }
	</style>
	<Title>
	Folder Sizes for "$Path"
	</Title>
"@
	$Pre = "<h1>Folder Sizes for ""$Path""</h1><h2>Run on $(Get-Date -f 'MM/dd/yyyy hh:mm:ss tt')</h2>"
	$Post = "<h2></h2>"
	$Post = "<h2>Total Space Used In ""$($Path)"":  $TotSize</h2></body></html>"

	# Build the report and save it to a file
	$HTML = $Report | Select-Object 'Folder Name','File Size', 'Groups Count' | Sort-Object 'Folder Name' | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $ReportPath
	
	
	# Cleanup/Reset the $Report Variable to Null, for Next Function Call
	$Script:Report = @()
}

# Call to getFolderSize to pull all three reports
getFolderSize $PathAFN $AFNReport
getFolderSize $PathAFS $AFSReport
getFolderSize $PathBDE $BDEReport

#Build the Notification Email and send it to the Recipients
Send-MailMessage -To $Recipients -From $Sender -Subject $Subject -Body $Body -SmtpServer $SMTPServer