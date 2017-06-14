<#
.SYNOPSIS
     Simple one line Description

.DESCRIPTION
     Full Description

.PARAMETER  Param1
     (Required, Default)

     Description

.PARAMETER  Param2
     (Required, Default)

     Description

.EXAMPLE
     .\SCRIPTNAME

     Description:

          Example description

.NOTES
     Name: SCRIPTNAME
     Author: Jonathan Durant
     Version: 1.0
     DateUpdated: DATE

.INPUTS
     NONE

.OUTPUTS
     NONE
#>

#======================================================

#region === Initialization ===


#region --- Advanced Binding ---

[cmdletbinding()]

#endregion --- Advanced Binding ---

#region --- Global Parameters From Command Line ---

Param (
     [Parameter(mandatory = $true, Position = 0)]
     [string]$GlblParam1 = '',
     [Parameter(mandatory = $false, Position = 1)]
     [string]$GlblParam2 = ''
)

#endregion --- Global Parameters From Command Line ---

#region --- Error Action Setting ---

$ErrorActionPreference = 'continue'

#endregion --- Error Action Setting ---

#region --- Global Variables ---

$Var1 = 'var1'
$Var2 = 'var2'

#endregion --- Global Variables ---


#endregion === Initialization ===

#======================================================

#region === Internal Functions ===

function Verb-Noun {
     # Internal Function One Description - Used internal to script only
     Param (
          [Parameter(mandatory = $true, Position = 0)]
          [string]$F1Param1 = ''
     )

     Begin {

     }
     Process {
          Try {
            Write-Host "Function1"
          }
          Catch {

          }
     }
     End {

     }
}

function Verb-Noun2 {
     # Internal Function Two Description - Used internal to script only
     Param (
          [Parameter(mandatory = $true, Position = 0)]
          [string]$F2Param1 = ''
     )

     Begin {

     }
     Process {
          Try {
                Write-Host "Function2"
          }
          Catch {

          }
     }
     End {

     }
}

#endregion === Internal Functions ===

#======================================================

#region === Main Execution ====

If($GlblParam1 -eq $true -or $GlblParam2 -eq $false){
     Verb-Noun -F1Param1 $Var1
}
Else {
     Verb-Noun2 -F2Param1 $Var2
}

#endregion === Main Execution ====