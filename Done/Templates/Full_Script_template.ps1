<#
.SYNOPSIS
    Simple one line Description

.DESCRIPTION
    Full Description

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.PARAMETER  Param1
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


#region --- Dot Source Libraries ---

. "c:\temp\script.ps1"

#endregion --- Dot Source Libraries ---

#region --- Error Action Setting ---

$ErrorActionPreference = "continue"

#endregion --- Error Action Setting ---

#region --- Global Variables ---

$Var1 = ""
$Var2 = ""

#endregion --- Global Variables ---

#region --- Advanced Binding ---

[cmdletbinding()]

#endregion --- Advanced Binding ---


#endregion === Initialization ===

#======================================================

#region === Internal Functions ===

function Verb-Noun {
    # Function Description
    Param (
        [Parameter(mandatory=$true,Position=0)]
            [string]$PARAM1 = ""
    )

    Begin {

    }

    Process {

        Try {

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

Param (
    [Parameter(mandatory=$true,Position=0)]
        [string]$PARAM1 = "",
    [Parameter(mandatory=$true,Position=1)]
        [string]$PARAM2 = ""
)

Begin {

}

Process {

    Try {

    }

    Catch {

    }   
}

End {

}

#endregion === Main Execution ====

