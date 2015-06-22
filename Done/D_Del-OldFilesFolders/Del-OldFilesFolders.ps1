<#
	.SYNOPSIS
		Deletes all Empty Folders and Files older the 14 days from all three conference folders for AFN, AFS, 
		and BDE

	.DESCRIPTION
		This script is self-contained and all variables are static. Using these variables, it will run three 
		calls to the "deleteOldFiles" function that will check files that they are under 14 days old, and if they
		are older it will delete them. Then it will check for any empty folders and delete them. 

		NOTE: To prevent a folder from being deleted add a "Readme.txt" file to it. This folder will then be 
		skipped by the folder deletion command.

	.NOTES
		Name: Del-OldFilesFolders
		Author: Jonathan Durant
		Organization: USANATO BDE G6
		Version: 1.0
		Organization: USANATO BDE G6
		DateUpdated: 2014-01-21
		Name:   Jonathan Durant

	.EXAMPLE
		.\Del-OldFileFolders

		Description:

			Default Syntax for the Script: Deletes all files older than 14 days,a nd empty folders in 
			all three conference folders.

	.INPUTS
		None
	.OUTPUTS
		None
 
#>

# Initialize Variables
$limit = (Get-Date).AddDays(-15)
$pathAFN = "\\ktes2292apcs280\usanato_technical\AFN_Conf"
$pathAFS = "\\ktes2292apcs280\usanato_technical\AFS_Conf"
$pathBDE = "\\ktes2292apcs280\usanato_technical\BDE_Conf"
$AFNLog = "AFN_Delete_Log.txt"
$AFSLog	= "AFS_Delete_Log.txt"
$BDELog	= "BDE_Delete_Log.txt"


function deleteOldFiles {
	Param (
		[string]$pathRoot,
		[string]$logFileName
	)
	
	Write-Output (Get-Date) | Out-File -Append "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\$logFileName"
	
	# Delete files older than the $limit.
	Write-Output "**** Removed Files ****" | Out-File -Append "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\$logFileName"
	
	Get-ChildItem -Path $pathRoot -Recurse -Force | 
	Where-Object { !$_.PSIsContainer -and $_.LastAccessTime -lt $limit -and $_.Name -ne "Readme.txt" } | 
	ForEach-Object { 
		Write-Output "Deleted $($_.FullName)"
		Remove-Item $_.FullName -Force
		Start-Sleep -Seconds 1
		} | 
	Out-File -Append "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\$logFileName"

	# Delete any empty directories left behind after deleting the old files.
	Write-Output "**** Removed Folders ****" | Out-File -Append "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\$logFileName"
	
	Get-ChildItem -Path $pathRoot -Recurse -Force | 
	Select-Object FullName,PSISContainer |
	Sort-Object -descending -property FullName |
	Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force |	Where-Object { !$_.PSIsContainer }) -eq $null } | 
	ForEach-Object { 
		Write-Output "Deleted $($_.FullName)"
		Remove-Item $_.FullName -Force -Recurse
		Start-Sleep -Seconds 1
		} | 
	Out-File -Append "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\$logFileName"

	Write-Output "----------------------------------------------------" | Out-File -Append "\\ktes2292apcs280\usanato_technical\development\Scripts\Script_Logs\$logFileName"
}

# Main Call to all delete functions

deleteOldFiles $pathAFN $AFNLog
deleteOldFiles $pathAFS $AFSLog
deleteOldFiles $pathBDE $BDELog