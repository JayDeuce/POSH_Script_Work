﻿# ---------------------------------------------------------------------------
# Author: Keith Hill
# Desc:   Prompt, colors and host window title updates suited to Windows XP
#         where folks tend to run with adminstrator privileges all the time.
# Date:   Nov 07, 2009
# Site:   http://pscx.codeplex.com
# Usage:  In your options hashtable place the following setting:
#
#         PromptTheme = 'WinXP'
# ---------------------------------------------------------------------------
#requires -Version 3
param([hashtable]$Theme)

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------
# Configure console colors.  $null signifies to use the console default color
if ($Pscx:IsAdmin -and ([System.Environment]::OSVersion.Version.Major -gt 5)) {
    $Theme.PromptForegroundColor = 'Red'
}

# ---------------------------------------------------------------------------
# Prompt ScriptBlock
# ---------------------------------------------------------------------------
$Theme.PromptScriptBlock = {
    param($nextCommandId) 
    
    # Determine what nesting level we are at (if any)
    $nestingLevel = ''
    if ($nestedpromptlevel -ge 1) {
        $nestingLevel = " [Nested:${nestedpromptlevel}]"
    }
    
    $promptChar = '>'
    if ($Pscx:IsAdmin) {
        $promptChar = '#'
    }
    
    # Output prompt string
    "${nextCommandId}${nestingLevel}$promptChar"
}

# ---------------------------------------------------------------------------
# Window Title Update ScriptBlock
# ---------------------------------------------------------------------------
$Theme.UpdateWindowTitleScriptBlock = {
    # Pretty much everyone runs as admin on XP and below so displaying the
    # admin status in the title bar is a waste of space. Only do this on Vista.
    $isVistaOrHigher = ([System.Environment]::OSVersion.Version.Major -gt 5)	

    $adminStatus = ''
    if ($Pscx:IsAdmin -and $isVistaOrHigher) { $adminStatus = 'Admin: ' }
        
    $location = Get-Location
    $version = $PSVersionTable.PSVersion
    
    $bitness = ''
    if ([IntPtr]::Size -eq 8) {
        $bitness = ' (x64)'
    }
    elseif ($Pscx:IsWow64Process) {
        $bitness = ' (x86)'
    }
    
    "$adminStatus$location - Windows PowerShell $version$bitness"
}

# SIG # Begin signature block
# MIIfUwYJKoZIhvcNAQcCoIIfRDCCH0ACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUKGzZWrx7/P0U7FApqRlzj4Bm
# /FagghqFMIIGajCCBVKgAwIBAgIQA5/t7ct5W43tMgyJGfA2iTANBgkqhkiG9w0B
# AQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVk
# IElEIENBLTEwHhcNMTMwNTIxMDAwMDAwWhcNMTQwNjA0MDAwMDAwWjBHMQswCQYD
# VQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRp
# bWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC6aUqBTW+lFBaqis1nvku/xmmPWBzgeegenVgmmNpc1Hyj+dsrjBI2w/z5ZAax
# u8KomAoXDeGV60C065ZtmL+mj3nPvIqSe22cGAZR2KUYUzIBJxlh6IRB38bw6Mr+
# d61f2J57jGBvhVxGvWvnD4DO5wPDfDHPt2VVxvvgmQjkc1r7l9rQTL60tsYPfyaS
# qbj8OO605DqkSNBM6qlGJ1vPkhGTnBan/tKtHyLFHqzBce+8StsBCUTfmBwtZ7qo
# igMzyVG19wJNCaRN/oBexddFw30IqgEzzDPYTzAW5P8iMi7rfjvw+R4y65Ul0vL+
# bVSEutXl1NHdG6+9WXuUhTABAgMBAAGjggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4Aw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAb8GA1UdIASC
# AbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYcaHR0cHM6Ly93
# d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkA
# IAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUA
# IABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAA
# bwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEA
# bgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIA
# ZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkA
# bABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUA
# ZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglg
# hkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStnAs0wHQYDVR0O
# BBYEFGMvyd95knu1I8q74aTuM37j4p36MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMDig
# NqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURD
# QS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3Nw
# LmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZIhvcNAQEFBQAD
# ggEBAKt0vUAATHYVJVc90xwD/31FyEUSZucoZWDY3zuz+g3BrDOP9IG5YfGd+5hV
# 195HQ7qAPfFIzD9nMFYfzvTQTIS9h6SexeEPqAZd0C9uXtwZ6PCH6uBOrz1sII5z
# b37WhxjghtOa/J7qjHLpQQ+4cbU4LPgpstUcop0b7F8quNw3IOHLu/DQbGyls8uf
# SvZU4yY0PS64wSsct/bDPf7RLR5Q9JTI+P3uc9tJtRv09f+lkME5FBvY7XEbapj7
# +kCaRKkpDlVeeLi3pIPDcAHwZkDlrnk04StNA6Et5ttUYhjt1QmLoqrWDMhPGr6Z
# JXhpmYnUWYne34jw02dedKWdpkQwggabMIIFg6ADAgECAhAK3lreshTkdg4UkQS9
# ucecMA0GCSqGSIb3DQEBBQUAMG8xCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xLjAsBgNVBAMTJURp
# Z2lDZXJ0IEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBLTEwHhcNMTMwOTEwMDAw
# MDAwWhcNMTYwOTE0MTIwMDAwWjBnMQswCQYDVQQGEwJVUzELMAkGA1UECBMCQ08x
# FTATBgNVBAcTDEZvcnQgQ29sbGluczEZMBcGA1UEChMQNkw2IFNvZnR3YXJlIExM
# QzEZMBcGA1UEAxMQNkw2IFNvZnR3YXJlIExMQzCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAI/YYNDd/Aw4AcjlGyyL+qjbxgXi1x6uw7Qmsjst/Z1yx0ES
# BQb29HmGeka3achcbRPgmBTt3Jn6427FDhvKOXhk7dPJ2mFxfv3NACa+Knvq/sz9
# xClrULvhpyOba8lOgXm5A9zWWBmUgYISVYz0jiS+/jl8x3yEEzplkTYrDsaiFiA0
# 9HSpKCqvdnhBjIL6MGJeS13rFXjlY5KlfwPJAV5txn4WM8/6cjGRDa550Cg7dygd
# SyDv7oDH7+AFqKakiE6Z+4yuBGhWQEBFnL9MZvlp3hkGK6Wlqy0Dfg3qkgqggcGx
# MS+CpdbfXF+pdCbSpuYu4FrCuDb+ae1TbyTiTBECAwEAAaOCAzkwggM1MB8GA1Ud
# IwQYMBaAFHtozimqwBe+SXrh5T/Wp/dFjzUyMB0GA1UdDgQWBBTpFzY/nfuGUb9f
# L83BlRNclRNsizAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMw
# cwYDVR0fBGwwajAzoDGgL4YtaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL2Fzc3Vy
# ZWQtY3MtMjAxMWEuY3JsMDOgMaAvhi1odHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# YXNzdXJlZC1jcy0yMDExYS5jcmwwggHEBgNVHSAEggG7MIIBtzCCAbMGCWCGSAGG
# /WwDATCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQuY29tL3Nz
# bC1jcHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5
# ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABl
# ACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAg
# AG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABh
# AG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwBy
# AGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBp
# AGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABl
# AGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wgYIG
# CCsGAQUFBwEBBHYwdDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQu
# Y29tMEwGCCsGAQUFBzAChkBodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURDb2RlU2lnbmluZ0NBLTEuY3J0MAwGA1UdEwEB/wQCMAAw
# DQYJKoZIhvcNAQEFBQADggEBAANu3/2PhW9plSTLJBR7SZBv4XqKxMzAJOw9GzNB
# uj4ihsyn/cRt1HV/ey7J9vM2mKZ5dZhU6rpb/cRnnKzEHDSSYnaogUDWbnBAw43P
# 6q6T9xKktrCpWhZRqbCRquix/VZN4dphqkdwpS//b/YnKnHi2da3MB1GqzQw6PQd
# mCWGHm+/CZWWI6GWZxdnRrDSkpMbkPYwdupQMVFFqQWWl/vJddLSM6bim0GD/XlU
# sz8hvYdOnOUT9g8+I3SegouqnrAOqu9Yj046iM29/6tkwyOCOKKeVl+uulpXnJRi
# nRkpczbl0OMMmIakVF1OTG/A/g2PPd6Xp4NDAWIKnsCdh64wggajMIIFi6ADAgEC
# AhAPqEkGFdcAoL4hdv3F7G29MA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMTAy
# MTExMjAwMDBaFw0yNjAyMTAxMjAwMDBaMG8xCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xLjAsBgNV
# BAMTJURpZ2lDZXJ0IEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBLTEwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCcfPmgjwrKiUtTmjzsGSJ/DMv3SETQ
# PyJumk/6zt/G0ySR/6hSk+dy+PFGhpTFqxf0eH/Ler6QJhx8Uy/lg+e7agUozKAX
# EUsYIPO3vfLcy7iGQEUfT/k5mNM7629ppFwBLrFm6aa43Abero1i/kQngqkDw/7m
# JguTSXHlOG1O/oBcZ3e11W9mZJRru4hJaNjR9H4hwebFHsnglrgJlflLnq7MMb1q
# WkKnxAVHfWAr2aFdvftWk+8b/HL53z4y/d0qLDJG2l5jvNC4y0wQNfxQX6xDRHz+
# hERQtIwqPXQM9HqLckvgVrUTtmPpP05JI+cGFvAlqwH4KEHmx9RkO12rAgMBAAGj
# ggNDMIIDPzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwggHD
# BgNVHSAEggG6MIIBtjCCAbIGCGCGSAGG/WwDMIIBpDA6BggrBgEFBQcCARYuaHR0
# cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUF
# BwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMG
# CCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2
# hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290
# Q0EuY3JsMB0GA1UdDgQWBBR7aM4pqsAXvkl64eU/1qf3RY81MjAfBgNVHSMEGDAW
# gBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAe3IdZP+I
# yDrBt+nnqcSHu9uUkteQWTP6K4feqFuAJT8Tj5uDG3xDxOaM3zk+wxXssNo7ISV7
# JMFyXbhHkYETRvqcP2pRON60Jcvwq9/FKAFUeRBGJNE4DyahYZBNur0o5j/xxKqb
# 9to1U0/J8j3TbNwj7aqgTWcJ8zqAPTz7NkyQ53ak3fI6v1Y1L6JMZejg1NrRx8iR
# ai0jTzc7GZQY1NWcEDzVsRwZ/4/Ia5ue+K6cmZZ40c2cURVbQiZyWo0KSiOSQOiG
# 3iLCkzrUm2im3yl/Brk8Dr2fxIacgkdCcTKGCZlyCXlLnXFp9UH/fzl3ZPGEjb6L
# HrJ9aKOlkLEM/zCCBs0wggW1oAMCAQICEAb9+QOWA63qAArrPye7uhswDQYJKoZI
# hvcNAQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNz
# dXJlZCBJRCBSb290IENBMB4XDTA2MTExMDAwMDAwMFoXDTIxMTExMDAwMDAwMFow
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBD
# QS0xMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6IItmfnKwkKVpYBz
# QHDSnlZUXKnE0kEGj8kz/E1FkVyBn+0snPgWWd+etSQVwpi5tHdJ3InECtqvy15r
# 7a2wcTHrzzpADEZNk+yLejYIA6sMNP4YSYL+x8cxSIB8HqIPkg5QycaH6zY/2DDD
# /6b3+6LNb3Mj/qxWBZDwMiEWicZwiPkFl32jx0PdAug7Pe2xQaPtP77blUjE7h6z
# 8rwMK5nQxl0SQoHhg26Ccz8mSxSQrllmCsSNvtLOBq6thG9IhJtPQLnxTPKvmPv2
# zkBdXPao8S+v7Iki8msYZbHBc63X8djPHgp0XEK4aH631XcKJ1Z8D2KkPzIUYJX9
# BwSiCQIDAQABo4IDejCCA3YwDgYDVR0PAQH/BAQDAgGGMDsGA1UdJQQ0MDIGCCsG
# AQUFBwMBBggrBgEFBQcDAgYIKwYBBQUHAwMGCCsGAQUFBwMEBggrBgEFBQcDCDCC
# AdIGA1UdIASCAckwggHFMIIBtAYKYIZIAYb9bAABBDCCAaQwOgYIKwYBBQUHAgEW
# Lmh0dHA6Ly93d3cuZGlnaWNlcnQuY29tL3NzbC1jcHMtcmVwb3NpdG9yeS5odG0w
# ggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgA
# aQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQA
# ZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcA
# aQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwA
# eQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkA
# YwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEA
# cgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIA
# eQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9bAMVMBIGA1UdEwEB/wQI
# MAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgw
# OqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYEFBUAEisTmLKZB+0e36K+
# Vw0rZwLNMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3
# DQEBBQUAA4IBAQBGUD7Jtygkpzgdtlspr1LPUukxR6tWXHvVDQtBs+/sdR90OPKy
# XGGinJXDUOSCuSPRujqGcq04eKx1XRcXNHJHhZRW0eu7NoR3zCSl8wQZVann4+er
# Ys37iy2QwsDStZS9Xk+xBdIOPRqpFFumhjFiqKgz5Js5p8T1zh14dpQlc+Qqq8+c
# dkvtX8JLFuRLcEwAiR78xXm8TBJX/l/hHrwCXaj++wc4Tw3GXZG5D2dFzdaD7eeS
# DY2xaYxP+1ngIw/Sqq4AfO6cQg7PkdcntxbuD8O9fAqg7iwIVYUiuOsYGk38KiGt
# STGDR5V3cdyxG0tLHBCcdxTBnU8vWpUIKRAmMYIEODCCBDQCAQEwgYMwbzELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNzdXJlZCBJRCBDb2RlIFNp
# Z25pbmcgQ0EtMQIQCt5a3rIU5HYOFJEEvbnHnDAJBgUrDgMCGgUAoHgwGAYKKwYB
# BAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAc
# BgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUrN+F
# eHtgt4NHExQmLfEj61HagU8wDQYJKoZIhvcNAQEBBQAEggEAS1BdRkwGbRdKg830
# 55dJZcGv7kwQKDDcq7DDPtpX343qOsjxg0Xt+PWJMYuaTwgI0cI0hmjyY1CT8HSK
# ntdQnWBRvg5dsb2LTdT7gcxODWche1GNiSGxgvg9jhUJaPz4S/iVHVretvaHhWjA
# FK5JG270Om7YJGLlbLchFdNf8JXeTV54nC3u3uwWS4rf9e4FJ4HGwoKcy+JrI+Ag
# QCnteiD7Z4kX4LELbH/8xNa5kmeV9nluvF1IgCo2x7TdgRvJXUGMr6g5fQl+qcQm
# 6SAR+qDYhDTweAKjfhV7feMewgNpqwCHNoADVZmpXkPHv8bz4YeCWIoBgbmgXmuO
# EVKgc6GCAg8wggILBgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQA5/t7ct5
# W43tMgyJGfA2iTAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEH
# ATAcBgkqhkiG9w0BCQUxDxcNMTMxMDE3MTQyNzM5WjAjBgkqhkiG9w0BCQQxFgQU
# LjQx/orAPHwsDSBnn4kPTwbJfVswDQYJKoZIhvcNAQEBBQAEggEAMMZQsXqeeZ73
# Z2laHdvhSaDkmkzoWW6V5sBIcSRu4ko1UyOVwjcrWvhl7WY2+4phrO0pkat8DYy8
# W2uHuXcQzGgHlPjX/kKIIo/S790zDPtJbCmOpveAthnO/NNjlIcqbpGEx9OpnJ4K
# P+T1B+1LNMu6jD5zhG3peurwmcFMvGMWqErBRRv4Nyo/XpWyUfGIhitWqTb5xUac
# D73o5ac9RpfuUxu3rnHrMwD7Ljc+0E5CA57nErazm/EPaN0R/Ohpsr8EcxOdD/v4
# c+LcwxwR5GOXjwliF57xu97vUY3LuJAZJCYT4EZD/PW3tJgJ7HTrRvj5CiH07vX4
# qMk+/f8Bhg==
# SIG # End signature block
