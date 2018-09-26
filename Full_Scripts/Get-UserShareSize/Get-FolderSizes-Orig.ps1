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
.PARAMETER Sort
    Specify which column you want the script to sort by.

    Valid colums are:
        Folder                  Sort by folder name
        Size                    Sort by folder size, largest to smallest
        Created                 Sort by the Created On column
        Changed                 Sort by the Last Updated column
        Owner                   Sort by Owner
.PARAMETER Descending
    Switch to control how you want things sorted.  By default the script
    will sort in an ascending order.  Use this switch to reverse that.
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
.NOTES
	Author:         Martin Pugh
	Twitter:        @thesurlyadm1n
	Spiceworks:     Martin9700
	Blog:           www.thesurlyadmin.com

	Changelog:
        1.41        @SPadminWV found a bug in the Total Size reporting.  Corrected.
        1.4         Add Sort and descending parameter
		1.3         Added Recurse parameter, default behavior is to now do 1 level of folders,
                    recurse will do all sub-folders.
        1.2         Added function to make the rows in the table alternating colors
		1.1         Updated to use ComObject Scripting.FileSystemObject which
		            should increase performance.  Inspired by MS Scripting Guy
					Ed Wilson.
		1.0         Initial Release
.LINK
	http://community.spiceworks.com/scripts/show/1738-get-foldersizes
.LINK
	http://community.spiceworks.com/topic/286820-how-to-export-list-all-folders-from-drive-the-list-should-include
.LINK
	http://blogs.technet.com/b/heyscriptingguy/archive/2013/01/05/weekend-scripter-sorting-folders-by-size.aspx
#>
Param (
     [string]$Path = "c:\dropbox\test",
     [string]$ReportPath = "c:\dropbox\test",
     [ValidateSet("Folder","Folders","Size","Created","Changed","Owner")]
     [string]$Sort = "Folder",
     [switch]$Descending,
     [switch]$Recurse
)

Function AddObject {
     Param (
          $FileObject
     )
     $Size = [double]($FSO.GetFolder($FileObject.FullName).Size)
     $Script:TotSize += $Size
     If ($Size) {
          $Size = CalculateSize $Size
     }
     Else {
          $Size = "0.00 MB"
     }
     $Script:Report += New-Object PSObject -Property @{
          'Folder Name' = $FileObject.FullName
          'Created on' = $FileObject.CreationTime
          'Last Updated' = $FileObject.LastWriteTime
          Size = $Size
          Owner = (Get-Acl $FileObject.FullName).Owner
     }
}

Function CalculateSize {
     Param (
          [double]$Size
     )
     If ($Size -gt 1000000000) {
          $ReturnSize = "{0:N2} GB" -f ($Size / 1GB)
     }
     Else {
          $ReturnSize = "{0:N2} MB" -f ($Size / 1MB)
     }
     Return $ReturnSize
}

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
          ForEach ($Line in $Lines) {
               $Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
               If ($ClassName -eq $CSSEvenClass) {
                    $ClassName = $CSSOddClass
               }
               Else {
                    $ClassName = $CSSEvenClass
               }
               Return $Line
          }
     }
}

Clear-Host

#Validate sort parameter
Switch -regex ($Sort) {
     "^folder.?$" {
          $SortBy = "Folder Name";Break
     }
     "created" {
          $SortBy = "Created On";Break
     }
     "changed" {
          $SortBy = "Last Updated";Break
     }
     default {
          $SortBy = $Sort
     }
}

$Report = @()
$TotSize = 0
$FSO = New-Object -ComObject Scripting.FileSystemObject

#First get the properties of the starting path
$Root = Get-Item -Path $Path
AddObject $Root
$TotalSize = CalculateSize $TotSize

#Now loop through all the subfolders
$ParamSplat = @{
     Path = $Path
     Recurse = $Recurse
}
ForEach ($Folder in (Get-ChildItem @ParamSplat | Where { $_.PSisContainer })) {
     AddObject $Folder
}

#Create the HTML for our report
$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<Title>
Folder Sizes for "$Path"
</Title>
"@

$TotSize = CalculateSize $TotSize

$Pre = "<h1>Folder Sizes for ""$Path""</h1><h2>Run on $(Get-Date -f 'MM/dd/yyyy hh:mm:ss tt')</h2>"
$Post = "<h2>Total Space Used In ""$($Path)"":  $TotalSize</h2></body></html>"

#Create the report and save it to a file
$HTML = $Report | Select 'Folder Name',Owner,'Created On','Last Updated',Size | Sort $SortBy -Descending:$Descending | ConvertTo-Html -PreContent $Pre -PostContent $Post -Head $Header | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File $ReportPath\FolderSizes.html

#Display the report in your default browser
& $ReportPath\FolderSizes.html
