<#
    .SYNOPSIS
        Send object as string to Notepad App

    .DESCRIPTION
        This function takes the parameter object, converts
        it to a formatted string, and sends it to a notepad
        file for easier viewing and manipulation of output.
        (Can also be sent as pipline output by sending through
        Out-String first.)

    .PARAMETER object
        (Required, No default)

        THis parameter is the object data you wish to convert and send to notepad

        Examples: $object
                  (get-process)
                  "text in a sentence"
                  "Textword"

    .EXAMPLE
        .\Out-Notepad $object

        Takes the $object (Can be anything) and sends it to notepad

    .EXAMPLE
        .\Out-Notepad (Get-Process)

        Runs the command Get-process and sends it to notepad.

    .EXAMPLE
        $object | Out-String | .\Out-Notepad

        Takes the $object Variable, converts it to string and sends it to notepad.

    .EXAMPLE
        (Get-Process) | Out-String | .\Out-Notepad

        Runs the command Get-process, converts it to string and sends it to notepad.
    #>
param (
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true)]
    [Object][AllowEmptyString()]$Object
)

begin {
    [Int]$Width = 150
    $al = New-Object System.Collections.ArrayList
}

Process {
    $null = $al.Add($Object)
    $text = $al | Format-Table -AutoSize -Wrap | Out-String -Width $Width
}

end {
    $process = Start-Process notepad -PassThru
    $null = $process.WaitForInputIdle()

    $sig = '
              [DllImport("user32.dll", EntryPoint = "FindWindowEx")]public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
              [DllImport("User32.dll")]public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);
            '
    $type = Add-Type -MemberDefinition $sig -Name APISendMessage2 -PassThru
    $hwnd = $process.MainWindowHandle

    [IntPtr]$child = $type::FindWindowEx($hwnd, [IntPtr]::Zero, "Edit", $null)
    $null = $type::SendMessage($child, 0x000C, 0, $text)
}
