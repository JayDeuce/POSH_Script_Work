<#
	.SYNOPSIS
		Get-ACLDump

	.DESCRIPTION
		A detailed description of the function.

	.PARAMETER  ParameterA
		The description of the ParameterA parameter.

	.PARAMETER  ParameterB
		The description of the ParameterB parameter.

	.EXAMPLE
		Get-Something -ParameterA 'One value' -ParameterB 32

	.EXAMPLE
		Get-Something 'One value' 32

	.INPUTS
		System.String,System.Int32

	.OUTPUTS
		System.String

	.NOTES
		Additional information about the function go here.

	.LINK
		about_functions_advanced

	.LINK
		about_comment_based_help

#>

Clear-Host
$path =  $args[0]   
$outPutFile =  $args[1]  
$startDate = Get-Date
$newLine = "`r`n" #is a carrage return/line feed. 
 
#check input parameters
if([System.IO.Directory]::Exists($path) -eq $false){
    throw (new-object System.IO.DirectoryNotFoundException("Directory does not exist or is missing!"))
}
If($path.EndsWith("\"))
{
    $path = $path.Remove($path.Length-1, 1)
}
if ([System.String]::IsNullOrEmpty($outPutFile)){
    throw (new-object System.ApplicationException("OutputFile is missing!"))
}
 
#Build information for the header of the output file, if file exist it will be owerwritten!  
$header = "Start: " + $startDate + $newLine + "Output file: " + $outPutFile + $newLine + "ACL of the analyzed path: " + $path + $newLine
$mainPathAcl = get-ACL $path | Format-List
out-file -encoding ASCII -filePath $outPutFile -InputObject $header
out-file -encoding ASCII -filePath $outPutFile -append -InputObject $mainPathAcl
 
#depth first traverse
$myStack = new-object  System.Collections.Stack 
[System.IO.DirectoryInfo]$rootInfo = New-Object System.IO.DirectoryInfo($path)
$myStack.Push($rootInfo)
 
while ($myStack.Count -ne 0){
    $actualItem = $myStack.Pop();  #get last item
    #add children to stack
    if ($actualItem -is [System.IO.DirectoryInfo])
    {
        [System.IO.FileSystemInfo[]]$dirs2 = $actualItem.GetFileSystemInfos() | Sort-Object Name -Descending
        if ($dirs2){#check if it is null.
            Foreach ($dir1 in $dirs2) {  #add to the stack
                $myStack.Push($dir1)
            }
        }
         
        if($actualItem.Parent.FullName -eq $rootInfo.FullName){
            $appHeader = "" + $newLine + "------------------------" + $newLine + $actualItem.Name + $newLine + "------------------------"
              out-file -encoding ASCII -filePath $outPutFile -append -InputObject $appHeader   
        }
    }
     
    #dump acls if not inherited
    $aclActFile = Get-Acl -Path $actualItem.FullName
    $WriteFileHeader = $true; 
    Foreach ($Access in $aclActFile.Access) { 
        $Inherited = [string]$Access.IsInherited 
        if ($Inherited -eq "False") {
            #write File Header
              if ($WriteFileHeader) {
                $fileHeader = "File: " + $actualItem.FullName + $newLine + "SDDL: " + $aclActFile.Sddl
                out-file -encoding ASCII -filePath $outPutFile -append -InputObject  $fileHeader
                $WriteFileHeader = $false;
            }
            #write AccessControl in csv
            $output = "ACL:  " + $Access.AccessControlType + ", " + $Access.IdentityReference + ", " + $Access.FileSystemRights 
            out-file -encoding ASCII -filePath $outPutFile -append -InputObject $output
        }    
    }
}
 
#Footer
$endDate = Get-Date
$elapsedTime = $endDate - $startDate
$footer = "" + $newLine + "Run completed at: " + $endDate + $newLine + "Elapsed Time:" + $newLine + $elapsedTime + $newLine
out-file -encoding ASCII -filePath $outPutFile -append -InputObject $footer