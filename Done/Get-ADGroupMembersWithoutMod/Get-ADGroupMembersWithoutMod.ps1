function Get-ADGroupMembers {
     <#
    .SYNOPSIS
	    Return all group members for specified groups.

    .FUNCTIONALITY
        Active Directory

    .DESCRIPTION
	    Return all group members for specified groups.   Requires .NET 3.5, does not require RSAT

    .PARAMETER Group
        One or more Security Groups to enumerate

    .PARAMETER Recurse
	    Whether to recurse groups.  Note that subgroups are NOT returned if this is true, only user accounts

        Default value is $True

    .EXAMPLE
        #Get all group members in Domain Admins or nested subgroups, only include samaccountname property
	    Get-ADGroupMembers "Domain Admins" | Select-Object -ExpandProperty samaccountname

    .EXAMPLE
        #Get members for objects returned by Get-ADGroupMembers
        Get-ADGroupMembers -group "Domain Admins" | Get-Member
    #>
     [cmdletbinding()]
     Param(
          [Parameter(Position=0, ValueFromPipeline=$true)]
          [string[]]$group = 'Domain Admins',
          [Parameter(Position=1)]
          [bool]$recurse = $false
     )

     Begin {
          # Add the .net type
          $netType = 'System.DirectoryServices.AccountManagement'
          Try {
               Add-Type -AssemblyName $netType -ErrorAction Stop
          }
          Catch {
               Throw "Could not load $netType`: Confirm .NET 3.5 or later is installed"
               Break
          }
          # Set up context type
          # Use the 'Machine' ContextType if you want to retrieve local group members
          # http://msdn.microsoft.com/en-us/library/system.directoryservices.accountmanagement.contexttype.aspx
          $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
     }
     Process {
          # List group members
          foreach($groupName in $group) {
               Try {
                    $groupInfo = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($contextType,$groupName)

                    # Display results or warn if no results
                    if($groupInfo) {
                         $groupInfo.GetMembers($recurse)
                    }
                    else {
                         Write-Warning "Could not find group '$groupName'"
                    }
               }
               Catch {
                    Write-Error "Could not obtain members for $groupName`: $_"
                    Continue
               }
          }
     }
     End{
          #cleanup
          $contextType = $groupInfo = $null
     }
}


#Get-ADGroupMembers "Work Admins"

Get-ADGroupMembers "Work Admins" | Select-Object -ExpandProperty Name

$test = Get-ADGroupMembers "Work Admins" | Select-Object -ExpandProperty name

if($test.Contains("Smith")) {
     Write-Host "Heck ya!!"
}
else {
     Write-Host "Hell NO!!"
}