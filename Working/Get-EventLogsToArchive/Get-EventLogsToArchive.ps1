#requires -version 2
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

function New-ZipFile {

    # Creates a new Zip File with the passed parameter as the filename
    Param (
        [Parameter(mandatory=$true,Position=0)]
            [string]$ZipFileName
    )

    Begin {
    }

    Process {

        Try {
            set-content $ZipFileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
            (dir $ZipFileName).IsReadOnly = $false
        }
        Catch {
        }   
    }

    End {
    }
}

function Add-ZipFile {

    # Creates a new Zip File with the passed parameter as the filename
    Param (
        [Parameter(mandatory=$true,Position=0)]
            [string]$ZipFileName
    )

    Begin {
        if(-not (test-path($zipfilename))) {
            New-ZipFile $ZipfFileName
        }

    }

    Process {

        Try {
            $shellApplication = new-object -com shell.application

            $zipPackage = $shellApplication.NameSpace($zipfilename)
      
            foreach($file in $input)
            {
                $zipPackage.CopyHere($file.FullName)
                Start-Sleep -milliseconds 500
            }
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
        [string]$Servers=""
)

Begin {

}

Process {

    Try {
        Add-ZipFile -ZipFileName C:\temp\test.zip
    }

    Catch {

    }   
}

End {

}

#endregion === Main Execution ====

