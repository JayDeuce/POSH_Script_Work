 # Function that indents to a level i
function Indent{
    Param([Int]$i)
    $Global:indent = $null
    For ($x=1; $x -le $i; $x++){
        $Global:indent += �&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�
    }
}

function Show-TreeItems{
    Param(
        [string]$location = "", 
        [string]$root
    )

    $Children = Get-ChildItem $location | Sort Name
    ForEach ($child in $children){
        Indent $i
        If ($child.PSIsContainer){
            
            Write-Output "<h3>$indent// $child \\</h3>"
            # Recurse through subdir
            $subPath = $child.FullName
            $i++
            Show-TreeItems $subPath $root
            $i--
        }
        else{
            $childPath = $child.FullName.Remove(0, $root.Length + 1)
            Write-Output "$indent<a href=`"$childPath`">$child</a><br />"
        }
    }

}


function New-TableOfContents($path)
{    
    $root = (Get-Item $path).FullName # Sets root of the main folder we are searching in, used to create relative paths within
    Write-Output "<html><head><title>Table of Contents</title><style>a:link{text-decoration:none;} a:visited{text-decoration:none} a:hover{text-decoration:underline;} a:active{text-decoration:underline;}</style></head><body><h1>eBook Listing</h1>"
    Show-TreeItems $path $root
    Write-Output "</body></html>"
}

New-TableOfContents "C:\Users\jwdur_000\Desktop" | Out-File "C:\Users\jwdur_000\Desktop\File_Index.html"
#New-TableOfContents "F:\Dropbox\Jon\eBooks\eBook_Unsorted" | Out-File "F:\Dropbox\Jon\eBooks\eBook_Unsorted\eBook_Unsorted.html"
