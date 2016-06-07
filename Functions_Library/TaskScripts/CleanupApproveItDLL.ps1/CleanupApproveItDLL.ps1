$script:ErrorActionPreference="SilentlyContinue"
kill -processname WinWord -Force
kill -processname Excel -Force
kill -processname Outlook -Force
kill -processname INFOPATH -Force

remove-item -path "C:\users\*\appdata\Roaming\Microsoft\Word\Startup\*veItEx.dot" -Recurse -Force
remove-item -path "C:\users\*\appdata\Roaming\Microsoft\Excel\XLSTART\*veItEx.xla" -Recurse -Force
remove-item -path "C:\Users\*\AppData\Roaming\Microsoft\Templates\*mal*.dot*" -Recurse -Force
remove-item -path "C:\Users\*\AppData\Roaming\Microsoft\Excel\*cel*.xlb" -Recurse -Force
remove-item -path "C:\Users\*\AppData\Roaming\Microsoft\infopath\*.tbs" -Recurse -Force
remove-item -path "C:\Users\*\AppData\Local\InfoPath\Controls\{08E623D3-BEAD-4bd3-8401-EFF51FD754CE}.ict" -Recurse -Force

remove-item -path HKCU:\"SOFTWARE\VB and VBA Program Settings\ApproveIt MS Office" -recurse
remove-item -path HKCU:\"SOFTWARE\Microsoft\Office\Word\Addins\ADTAPILib.ApproveItAddin" -recurse
remove-item -path HKCU:\"SOFTWARE\Microsoft\Office\Excel\Addins\ADTAPILib.ApproveItAddin" -recurse
remove-item -path HKCU:\"SOFTWARE\Microsoft\Office\11.0\Word\Addins\ADTAPILib.ApproveItAddin" -recurse
remove-item -path HKCU:\"SOFTWARE\Microsoft\Office\12.0\Word\Addins\ADTAPILib.ApproveItAddin" -recurse
remove-item -path HKCU:\"SOFTWARE\Microsoft\Office\11.0\Excel\Addins\ADTAPILib.ApproveItAddin" -recurse
remove-item -path HKCU:\"SOFTWARE\Microsoft\Office\12.0\Excel\Addins\ADTAPILib.ApproveItAddin" -recurse
remove-item -path HKCU:\"Software\Classes\ApproveItDesignerAddIn" -recurse
remove-item -path HKCU:\"Software\Classes\CLSID\{97A21885-E335-4164-AD1C-8A3BF0F003E9}" -recurse
$null = new-psdrive -name HKCR -psprovider registry -root HKEY_CLASSES_ROOT
remove-item -path HKCR:\"CLSID\{08E623D3-BEAD-4bd3-8401-EFF51FD754CE}" -recurse
remove-item -path HKCU:\"Software\Microsoft\Office\InfoPath\Addins\ApproveItDesignerAddIn" -recurse
remove-item -path HKCU:\"Software\VB and VBA Program Settings\ApproveIt MS Office" -recurse
remove-item -path HKCU:\"Software\Silanis" -recurse