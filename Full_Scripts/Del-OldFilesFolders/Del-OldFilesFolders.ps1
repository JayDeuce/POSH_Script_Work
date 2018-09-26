<#
	.SYNOPSIS
		Deletes all Empty Folders and Files older the 14 days from all three conference folders

	.DESCRIPTION
		This script is self-contained and all variables are static. Using these variables, it will run three
		calls to the "deleteOldFiles" function that will check files that they are under 14 days old, and if they
		are older it will delete them. Then it will check for any empty folders and delete them.

		NOTE: To prevent a folder from being deleted add a "Readme.txt" file to it. This folder will then be
		skipped by the folder deletion command.

	.NOTES
		Name: Del-OldFilesFolders
		Author: Jonathan Durant
		Version: 1.0
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
$path1 = "\\server1\folder\1_Conf"
$path2 = "\\server1\folder\2_Conf"
$path3 = "\\server1\folder\2_Conf"
$1Log = "1_Delete_Log.txt"
$2Log = "2_Delete_Log.txt"
$3Log = "3_Delete_Log.txt"


function deleteOldFiles {
     Param (
          [string]$pathRoot,
          [string]$logFileName
     )

     Write-Output (Get-Date) | Out-File -Append "\\server1\folder\$logFileName"

     # Delete files older than the $limit.
     Write-Output "**** Removed Files ****" | Out-File -Append "\\server1\folder\$logFileName"

     Get-ChildItem -Path $pathRoot -Recurse -Force |
          Where-Object { !$_.PSIsContainer -and $_.LastAccessTime -lt $limit -and $_.Name -ne "Readme.txt" } |
          ForEach-Object {
               Write-Output "Deleted $($_.FullName)"
               Remove-Item $_.FullName -Force
               Start-Sleep -Seconds 1
          } |
          Out-File -Append "\\server1\folder\$logFileName"

     # Delete any empty directories left behind after deleting the old files.
     Write-Output "**** Removed Folders ****" | Out-File -Append "\\server1\folder\$logFileName"

     Get-ChildItem -Path $pathRoot -Recurse -Force |
          Select-Object FullName, PSISContainer |
          Sort-Object -descending -property FullName |
          Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } |
          ForEach-Object {
               Write-Output "Deleted $($_.FullName)"
               Remove-Item $_.FullName -Force -Recurse
               Start-Sleep -Seconds 1
          } |
          Out-File -Append "\\server1\folder\$logFileName"

     Write-Output "----------------------------------------------------" | Out-File -Append "\\server1\folder\$logFileName"
}

# Main Call to all delete functions

deleteOldFiles $path1 $1Log
deleteOldFiles $path2 $2Log
deleteOldFiles $path3 $3Log