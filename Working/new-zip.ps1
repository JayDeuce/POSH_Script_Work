﻿param(
    [string]$zipfilename
)

set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
(dir $zipfilename).IsReadOnly = $false
