DIM fso
Set fso = CreateObject("Scripting.FilesystemObject")
On Error Resume Next
fso.CopyFile ".\Cisco AnyConnect Secure Mobility Client.lnk", "C:\Users\Public\Desktop\"
fso.CopyFile ".\drivemapper.lnk", "C:\Users\Public\Desktop\"
fso.CopyFile ".\drivemapper.exe", "C:\Program Files\"