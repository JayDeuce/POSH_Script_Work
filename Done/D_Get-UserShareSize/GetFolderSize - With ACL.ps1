<#
.SYNOPSIS
    Get folder sizes in specified tree.  
.DESCRIPTION
    Script creates an HTML report with owner information, when created, 
    when last updated and folder size.  By default script will only do 1
    level of folders.  Use Recurse to do all sub-folders.
    
    Update the PARAM section to match your environment.
.PARAMETER Path
    Specify the path you wish to report on
.PARAMETER ReportPath
    Specify where you want the HTML report to be saved
.PARAMETER Recurse
    Report on all sub-folders
.EXAMPLE
    .\Get-FolderSizes
    Run script and use defaults
.EXAMPLE
    .\Get-FolderSizes -Path "c:\Windows" -ReportPath "c:\Scripts"
    Run the script and report on all folers in C:\Windows.  Save the
    HTML report in C:\Scripts
.EXAMPLE
    .\Get-FolderSizes -Path "c:\Windows" -ReportPath "c:\Scripts" -Recurse
    Run the script and report on all folers in C:\Windows.  Save the
    HTML report in C:\Scripts.  Report on all sub-folders.
.OUTPUTS
    FolderSizes.HTML in specified Report Path
#>

# Setup Script Scope Variables
$Report = @()
$TotSize = 0
# Paths Hard Coded for use in Task Scheduler
$pathAFN = "\\ktes2292apcs280\usanato_afnorth_users"
$pathAFS = "\\ktes2292apcs280\usanato_afsouth_users"
$pathBDE = "\\ktes2292apcs280\usanato_bdehq_users"
$AFNReport = "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\AFN_Users_Folder_Size.html"
$AFSReport	= "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\AFS_Users_Folder_Size.html"
$BDEReport	= "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\BDE_Users_Folder_Size.html"

# Function Gather info on the Current Path,and set Report Object Variable Properties
Function AddObject {
    Param ( 
        $FileObject
    )
	# Gathers Sizes info on Current File Path Obejct, send it to CalculateSize function to convert it to 2 decimal string
    $Size = [double]($FSO.GetFolder($FileObject.FullName).Size)
    $Script:TotSize += $Size
    $Size = CalculateSize $Size
	
    # Gathers info on Access Control List Groups of the Current file path object and Set proper variables for use in the Report Object
    $Owner = [string]((get-acl $FileObject.FullName).Access | ForEach-Object { $_.identityReference.value })
	$OwnerCount = ((get-acl $FileObject.FullName).Access).Count
	
	#Builds the Report Object in the Script Scope Report Variable
    $script:Report += New-Object PSObject -Property @{
        'Folder Name' = $FileObject.Name
        'File Size' = $Size
        'Assigned Groups' = $Owner
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

function getFolderSize () {
	Param (
		[string]$Path = "",
		[string]$ReportPath = "",
   		[switch]$Recurse
	)
	
	
	$FSO = New-Object -ComObject Scripting.FileSystemObject 

	#First get the properties of the starting path
	$Root = Get-Item -Path $Path 
	AddObject $Root

	#Now loop through all the subfolders
	$ParamSplat = @{
    	Path = $Path
    	Recurse = $Recurse
	}
	ForEach ($Folder in (Get-ChildItem @ParamSplat | Where-Object { $_.PSisContainer }))
	{   
    	AddObject $Folder
	}
	
	#Create the HTML for our report
	$Header = @"
	<style>
	TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 6px;border-style: solid;border-color: black;background-color: #6495ED;}
	TD {border-width: 1px;padding: 6px;border-style: solid;border-color: black;}
	TABLE TR TH:nth-child(1) {width: 600px; word-break:break-all;}
	TABLE TR TD:nth-child(1) {width: 600px; word-break:break-all;}
	TABLE TR TH:nth-child(2) {width: 600px;}
	TABLE TR TD:nth-child(2) {width: 600px;}
	TABLE TR TH:nth-child(3) {width: 100px; text-align: center;}
	TABLE TR TD:nth-child(3) {width: 100px; text-align: center;}
	.odd  { background-color:#ffffff; }
	.even { background-color:#dddddd; }
	</style>
	<Title>
	Folder Sizes for "$Path"
	</Title>
"@

	# $TotSize = CalculateSize $TotSize

	$Pre = "<h1>Folder Sizes for ""$Path""</h1><h2>Run on $(Get-Date -f 'MM/dd/yyyy hh:mm:ss tt')</h2>"
	$Post = "<h2></h2>"
	$Post = "<h2>Total Space Used In ""$($Path)"":  $TotSize</h2></body></html>"

	#Create the report and save it to a file
	$HTML = $Report | Select-Object 'Folder Name', 'Assigned Groups', 'File Size', 'Groups Count' | Sort-Object 'Folder Name' | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $ReportPath
	
	# Cleanup/Reset the $Report Variable to Null, for Next Function Call
	$Script:Report = @()
}

getFolderSize $pathAFN $AFNReport
#getFolderSize $pathAFS $AFSReport
#getFolderSize $pathBDE $BDEReport