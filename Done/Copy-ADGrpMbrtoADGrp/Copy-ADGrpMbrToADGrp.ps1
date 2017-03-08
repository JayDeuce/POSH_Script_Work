<#
	.SYNOPSIS
		Copies members from one Active Diretory Group to Another Active Directory Group

	.DESCRIPTION
		Using passed parameters for the source and destination Active Directory Groups; this script
		will copy the members of the source group to the destination Group. This script will use the
		parameters to find the disinguished name of the groups and use them to gather the proper member
		data for the copy.

		!! THIS SCRIPT MUST BE RUN WITH ACTIVE DIRECTORY ADMINISTRATOR RIGHTS !!

	.PARAMETER  SrcGrpIn
		This parameter is the Name (Not Disiguished name) in Active Directory of the Source AD Group. This
		can use the WildCard (*) symbol to search AD for a group. This will show a numbered list all groups
		found with the search term within the AD Structure, choose a number to set the parameter. The
		Wildcard can be set before an Alpha Character (*check) which will search for all groups ending in
		"Check", After an Alpha Character (Check*) which will find all groups beggining with "Check", or
		Before and After an Alpha Character (*Check*) which will find all groups with "Check" in the Name.

	.PARAMETER  DestGrpIn
		This parameter is the Name (Not Disiguished name) in Active Directory of the Destination AD Group

	.NOTES
		Name: Copy-ADGrpMbrToADGrp
		Author: Jonathan Durant
		Version: 1.0
		DateUpdated: 2014-03-09

	.EXAMPLE
		Copy-ADGrpMbrtoADGrp "SrcADGroup" "DestADGroup"

		Description:

			Default Syntax for the Script: Copies the members of the Source group "SrcADGroup" to the
			Destination Group "DestADGroup"

	.EXAMPLE
		Copy-ADGrpMbrtoADGrp "*Src*" "DestADGroup"

		Description:

			Wildcard Mode for the Script: Will find all AD Groups with "Src" in the name and list them
			for you choose one to set the source parameter, then copy the members to the "DestADGrpup"
			AD Group.

	.INPUTS
		None
	.OUTPUTS
		None

#>
[cmdletbinding()]
# Initialize Parameters
Param (
     [parameter(Mandatory=$true)][string]$srcGrpIn = "",
     [parameter(Mandatory=$true)][string]$destGrpIn = ""
)

#----Functions----
# function to find the Distiguished Name of an AD Object. Pass the Object type (Group, User, etc) and the Name of the Object
function Find-DN {
     [cmdletbinding()]
     Param
     (
          [string]$type,
          [string]$name
     )
     $root = [ADSI]''
     $searchObject = New-Object System.DirectoryServices.DirectorySearcher($root)
     $searchObject.filter = "(&(objectClass=$type) (CN=$name))"
     $adFind = $searchObject.findall()

     # just in case it finds 2 or more items it will ask which one first
     if ($adFind.count -gt 1) {
          $count = 0
          foreach	($i in $adFind) {
               Write-Host $count ": " $i.path
               $count += 1
          }
          $selection = Read-Host "Please select the correct group: "

          $result = $adFind[0].Path.Replace("LDAP://","")
          return $result
     }
     $result = $adFind[0].Path.Replace("LDAP://","")
     return $result
}

#----Main Script----

# Import Active Directory Module
Import-Module ActiveDirectory

# Set Variables
$sGroup = Find-DN "group" $srcGrpIn
$dGroup = Find-DN "group" $destGrpIn

# Get Members of the Source Group
$sMembers = Get-ADGroupMember -Identity $sGroup

# Loop through Source Group member Objects and copy the members from the Source Group to the Destination Group
foreach ($user in $sMembers) {
     Add-ADGroupMember -Identity $dGroup -Members $user.distinguishedname
}
