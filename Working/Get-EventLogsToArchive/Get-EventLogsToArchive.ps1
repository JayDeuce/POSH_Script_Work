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
        [string]$zipFileName
    )
    Begin {
    }
    Process {
        Try {
            set-content $zipFileName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
            (dir $zipFileName).IsReadOnly = $false
        }
        Catch {
        }   
    }
    End {
        return
    }
}

function Add-ZipFile {    
    Param (
        [Parameter(mandatory=$true,Position=0)]
        [string]$zipFileName,
        [Parameter(mandatory=$true,Position=1)]
        [array]$fileName
    )

    Begin { 
        if(-not (test-path $zipfilename)) {
            New-ZipFile $zipFileName
        }
        $zipFilename = Resolve-Path $zipFilename
        $fileName = Resolve-Path $fileName    
    }

    Process {
        Try {
            
            $shellApplication = new-object -com shell.application
            $zipPackage = $shellApplication.NameSpace($zipFilename)
            
            foreach($file in $fileName){
                Write-Host $file
                $zipPackage.CopyHere($file)
                Start-sleep -milliseconds 500
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

