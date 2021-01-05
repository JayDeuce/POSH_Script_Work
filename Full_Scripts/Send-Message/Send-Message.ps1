<#
.SYNOPSIS
    Send a message to designated computers

.DESCRIPTION
    This script uses Windows Forms to present a GUI for sending messages to remote computers on a network using msg.exe.
    You can send to a single computer, default list choices, or multiple computers from a text file.

    There are no parameters, examples, or any command line controls as this is a GUI based Script. Run the script from
    the command line and a GUI will open for your use.

    You can also create a EXE file for this script using the "PS2EXE" project located here (Read comments for updates to
    source code to update PS2EXE to PowerShell 5.0):

        https://gallery.technet.microsoft.com/PS2EXE-Convert-PowerShell-9e4e07f1

    This script based on the Technet Script "MessageCenter 1.4" by Chris Carter. Get message center here for comparison:

        https://www.powershellgallery.com/packages/MessageCenter/1.4/DisplayScript

    This script is licensed by Chris Carter under the Creative Commons License found here:

        http://creativecommons.org/licenses/by-sa/4.0/

    REQUIREMENTS:

    !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.NOTES
    Name: Send-Message
    Author: Jonathan Durant (Based on Chris Carter's Script MesageCenter)
    Version: 1.5
    DateUpdated: 2016-06-27

.INPUTS
    NONE
.OUTPUTS
    NONE

#>

#----------------------------
#    Form Initialization
#----------------------------

#region Set Error Action Preference

$ErrorActionPreference = "SilentlyContinue"

#endregion $ErrorActionPreference = "SilentlyContinue"

#region Import the necessary Assemblies

[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

#endregion Import the necessary Assemblies

#region Generate Main Form Objects

# Main Form Object creation
$sendForm = New-Object System.Windows.Forms.Form

# Objects for Menu Bar
# Each object below builds the menu at the top of the form
$menu = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem # Root File Menu Text
$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem # Root Help Menu text
$menuFileOpen = New-Object System.Windows.Forms.ToolStripMenuItem # Open File Menu Option
$menuFileQuit = New-Object System.Windows.Forms.ToolStripMenuItem # Quit Menu Option
$menuHelpDirect = New-Object System.Windows.Forms.ToolStripMenuItem # Main Script Help Menu Option
$menuHelpView = New-Object System.Windows.Forms.ToolStripMenuItem # View Script Source Menu option
$separatorF = New-Object System.Windows.Forms.ToolStripSeparator # Seperator Line in the File Menu

# Add Other Objects to Form
<#
tabControl = The full tab object
MsgTab = The Send Message Tab
buttonClose = The Close button at the bottom of the form
buttonSend = The Send button at the bottom of the form
progressBar = The Progress bar object on the Progress bar form
progressLabel = Label to show which machine is processing on the prgoress bar form
grpListComp = The Send to Computers section of the tab
autoEntListComp = Automatic list choices for the computer entry section Title
labelAutoEntListComp = Instructions Label for the Auto List entry area
checkBoxBearIP = Check box for the Bear Default IP List
checkBoxPearIP = Check box for the Pear Default IP List
checkBoxFigsIP = Check box for the Figs Default IP List
checkBoxFireIP = Check box for the Fire Default IP List
buttonManEntClearComp = The Clear button in the Send to Computers sectio
buttonManEntListComp = The Open button in the Send to Computers section
textBoxManEntListComp = The Text box in the Send to Computers section
labelManEntListComp = The Text in the Send to Computers section
labelMsg = The text in the Type your Message section
richtextBoxMsg = The rich text box in the type your Message Section
openFileDialog1 = The open file dialog window and its options
InitialForWindowState = the initial settings of the form window when it opens
#>

$tabControl = New-Object System.Windows.Forms.TabControl
$MsgTab = New-Object System.Windows.Forms.TabPage
$buttonClose = New-Object System.Windows.Forms.Button
$buttonSend = New-Object System.Windows.Forms.Button
$processingLabel = New-Object System.Windows.Forms.Label
$grpListComp = New-Object System.Windows.Forms.GroupBox
$autoEntListComp = New-Object System.Windows.Forms.Label
$labelAutoEntListComp = New-Object System.Windows.Forms.Label
$checkBoxBearIP = New-Object System.Windows.Forms.CheckBox
$checkBoxPearIP = New-Object System.Windows.Forms.CheckBox
$checkBoxFigsIP = New-Object System.Windows.Forms.CheckBox
$checkBoxFireIP = New-Object System.Windows.Forms.CheckBox
$buttonManEntClearComp = New-Object System.Windows.Forms.Button
$buttonManEntListComp = New-Object System.Windows.Forms.Button
$textBoxManEntListComp = New-Object System.Windows.Forms.RichTextBox
$manEntListComp = New-Object System.Windows.Forms.Label
$labelManEntListComp = New-Object System.Windows.Forms.Label
$labelMsg = New-Object System.Windows.Forms.Label
$richTextBoxMsg = New-Object System.Windows.Forms.RichTextBox
$openFileDialog1 = New-Object System.Windows.Forms.OpenFileDialog
$initialFormWindowState = New-Object System.Windows.Forms.FormWindowState

#endregion Generate Main Form Objects

#region Global Variables

#region IP Lists

# Single IP Ranges used to create the IPs Variable
[array]$155IPlist = 1..254 | ForEach-Object {"192.168.155.$_"}
[array]$199IPlist = 1..254 | ForEach-Object {"192.161.199.$_"}
[array]$255IPlist = 1..254 | ForEach-Object {"192.171.205.$_"}
# Build IPs Variable
[array]$bearIPs = $155IPlist + $199IPlist + $255IPlist
[array]$figsIPs = 1..254 | ForEach-Object {"192.168.166.$_"}
[array]$pearsIPs = 1..254 | ForEach-Object {"192.168.244.$_"}
[array]$fireIPs = 1..254 | ForEach-Object {"192.168.250.$_"}
[array]$targetList

#endregion IP Lists

#endregion Global Variables

#region Script Functions

Function Create-MessageBox {
     # Creates Pop-Up Message Boxes to Inform Users of Output
     Param (
          [Parameter(Mandatory = $true)][string]$message,
          [Parameter(Mandatory = $true)][string]$title,
          [Parameter(Mandatory = $false)][System.Windows.Forms.MessageBoxButtons]$buttons = "OK",
          [Parameter(Mandatory = $false)][System.Windows.Forms.MessageBoxIcon]$icon = "Information"
     )
     [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
}

Function Check-IfAdmin {
     # Get the ID and security principal of the current user account
     $myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
     $myWindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
     # Get the security principal for the Administrator role
     $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
     # Check to see if we are currently running "as Administrator" and message user to relaunch as admin if not.
     if (!($myWindowsPrincipal.IsInRole($adminRole))) {
          Create-MessageBox -message "You are not running as an administrator, please right-click Send-Message.exe and 'RUN AS ADMINISTRATOR' to send a message!" -title "ERROR: NOT AN ADMIN" -icon Error
          exit
     }
}

Function Parse-TextFile ([string]$path, [System.Windows.Forms.RichTextBox]$textBox) {
     # Get text from file, combine into string and insert into text box
     $compNames = (Get-Content $path) -join ', '
     $TextBox.Text = $compNames
}

Function Parse-Input ([string]$workstations) {
     # Split inputed text into array of single entries delimiting on whitespace, commas, and semicolons
     [array]$arrComp = (($workstations -replace '[,;\s]', ' ').Trim()) -split '\s+'
     $Script:targetList += $arrComp
}

Function Set-TargetVariable ([array]$ipList, [System.Windows.Forms.CheckBox]$CheckBox) {
     if ($CheckBox.Checked) {
          $Script:targetList += $ipList
     }
     else {
          foreach ($ip in $ipList) {
               $targetListNew = $Script:targetList -ne $ip
               $Script:targetList = $targetListNew
          }
     }
}

Function Check-WMIStatus($workstation) {
     $wmiNameCheck = (Get-WmiObject Win32_OperatingSystem -ComputerName $workstation -ErrorVariable wmiErr).name
     if ($wmiErr -like "*Access is Denied*") {
          return "adminError"
     }
     elseif ($wmiNameCheck -eq $null) {
          return "null"
     }
}

Function Send-Message ([string]$message, [array]$workstations) {

     # Main Send Message Function
     # Initialize error counters and failed computer names variables
     [int]$errCount = 0
     [int]$errConnectCount = 0
     [int]$errWinCount = 0
     [int]$errAdminWinCount = 0
     [int]$errSendCount = 0
     [int]$posSendCount = 0
     [string]$resultsMsgBegin
     [string]$workGramNoun
     [string]$workGramVerb
     [array]$failConnectCompNames = @()
     [array]$failWinCompNames = @()
     [array]$failAdminWinCompNames = @()
     [array]$failSendCompNames = @()
     [array]$failConnectNames = @()
     [array]$failWinNames = @()
     [array]$failSendNames = @()
     [array]$failConnectNames = @()
     [array]$failWinNames = @()
     [array]$failSendNames = @()

     # Test to ensure there is message text
     if ($message) {
          if ($message.Length -ge 256) {
               Create-MessageBox -message "The entered message is too long to send through this sysem. Please rewrite the message so it is below 256 Characters." -title "Message too Long" -icon Error
          }
          else {
               # Test for at least one computer has been entered
               if ($workstations -or $checkBoxBearIP.checked -or $checkBoxPearIP.Checked -or $checkBoxFigsIP.Checked -or $checkBoxFireIP.Checked) {

                    # Iterate through array of computer names
                    foreach ($workstation in $workstations) {

                         $processingLabel.Show()
                         $processingLabel.Text = "Processing: $workstation"

                         if (!(test-connection -count 1 -Quiet -ComputerName $workstation)) {
                              # Test if computer is online; if not, add to error count, if so; continue to send message
                              $errCount += 1
                              $errConnectCount += 1
                              $failConnectCompNames += $workstation
                         }
                         elseif ((Check-WMIStatus($workstation)) -eq "adminError") {
                              $errCount += 1
                              $errAdminWinCount += 1
                              $failAdminWinCompNames += $workstation
                         }
                         elseif ((Check-WMIStatus($workstation)) -eq "null") {
                              $errCount += 1
                              $errWinCount += 1
                              $failWinCompNames += $workstation
                         }
                         else {
                              # Call msg.exe to send message to each computer in the array
                              # Message will stay up on Computer until clicked away by user or for 24hours whichever comes first
                              #Invoke-Expression "msg.exe * /time:86400 /server:$workstation `"$message`""
                              #powershell.exe -command "Invoke-Command -ScriptBlock {msg.exe * /time:86400 /server:$workstation `"$message`" }"
                              #$job = Start-Job -Name "job_$workstation" -ScriptBlock {Invoke-Expression "msg.exe * /time:86400 /server:$workstation `"$message`""}
                              $job = Start-Job -Name "job_$workstation" -ScriptBlock {param($sbWorkstation, $sbMessage) Invoke-Expression "msg.exe * /time:86400 /server:$sbWorkstation `"$sbMessage`""} -ArgumentList $workstation, $message

                              # Test exit code from msg.exe and add to failed to send error count and the computer name otherwise set a postive completion count
                              if ($job.State -eq 'Failed') {
                                   $errCount += 1
                                   $errSendCount += 1
                                   $failSendCompNames += $workstation
                              }
                              else {
                                   $posSendCount += 1
                              }
                         }
                    }

                    # Configure variables for Results message
                    $failConnectNames = $failConnectCompNames -join "`r`n"
                    $failWinNames = $failWinCompNames -join "`r`n"
                    $failAdminWinNames = $failAdminWinCompNames -join "`r`n"
                    $failSendNames = $failSendCompNames -join "`r`n"

                    # Set Plural or Singular nouns and verbs
                    if ($Script:targetList.count -eq 1) {
                         $workGramNoun = "workstation"
                         $workGramVerb = "is"
                    }
                    else {
                         $workGramNoun = "workstations"
                         $workGramVerb = "were"
                    }

                    # Set Positive Singular or Plural Message
                    if ($posSendCount -eq 1) {
                         $posResultsMsgBegin = "A Message was"
                    }
                    else {
                         $posResultsMsgBegin = "Messages were"
                    }

                    # Set Error Singular or Plural Message
                    if ($errCount -eq 1) {
                         $errResultsMsgBegin = "$errCount message was"
                    }
                    else {
                         $errResultsMsgBegin = "$errCount messages were"
                    }

                    # Generate Information Portion of Results Form
                    $infoResultsMessage = @"

                        - SEND-MESSAGE RESULTS -

The following are the results of sending the message to $($Script:targetlist.count) $workGramNoun`:

(If successful, the message pop-up will stay on the $workGramNoun screen for at least 24 hours or until read and clicked away by the user.)

                        *************************

"@

                    # Generate Positive portion of Results Form
                    $posResultsMessage = @"

                              - SUCCESSES -
                              -------------
$posResultsMsgBegin successfully sent to the $posSendCount $workGramNoun.

"@

                    # Generate Error portion of Results Form
                    $errResultsMessage = @"

                        *************************

                               - ERRORS -
                               ----------
$errResultsMsgBegin unable to be sent to the below $workGramNoun.

Please reveiw the below listing for the reason(s) for failure and the associated $workGramNoun

Failures List:
--------------

$errConnectCount $workGramNoun $workGramVerb not online and could not be connected to:

$failConnectNames
                        =========================

$errWinCount $workGramNoun $workGramVerb not Microsoft Windows, so the message was not sent:

$failWinNames
                        =========================

$errAdminWinCount $workGramNoun failed due to the account used not having Admin Access to the machine, so the message was not sent:

$failAdminWinNames
                        =========================

$errSendCount $workGramNoun failed to receive the message due to another issue:

$failSendNames
                         =========================
"@

                    # Contruct Results Message
                    If ($errCount -ge 1) {
                         $resultMessage = $infoResultsMessage + $posResultsMessage + $errResultsMessage
                    }
                    else {
                         $resultMessage = $infoResultsMessage + $posResultsMessage
                    }

                    #Pop up results Dialog
                    Create-ResultsForm($resultMessage)
               }
               else {
                    #Pop up for no computers entered
                    Create-MessageBox -message "You must enter at least one computer." -title "No Computers Chosen" -icon Error
               }
          }
     }
     else {
          # Pop up for no message entered
          Create-MessageBox -message "There is no message to send" -title "No Message" -icon Error
     }
}

Function Create-ViewSourceForm {
    # Add objects for Source Viewer
    $formSourceCode = New-Object System.Windows.Forms.Form
    $richTextBoxSource = New-Object System.Windows.Forms.RichTextBox
    # Form for viewing source code
    $System_Drawing_Size = New-Object System.Drawing.Size
    # Set Size of View Source Form
    $formSourceCode.WindowState = "Maximized"
    $formSourceCode.ClientSize = $System_Drawing_Size
    $formSourceCode.DataBindings.DefaultDataSourceUpdateMode = 0
    $formSourceCode.StartPosition = "CenterScreen"
    $formSourceCode.Name = "formSourceCode"
    # Set View Source Form Title
    $formSourceCode.Text = "Send Message Script Source View"
    $formSourceCode.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSCommandPath)
    # Add Rich Text Box for the Help Text to display in
    $richTextBoxSource.Anchor = 15
    $richTextBoxSource.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    # Set Location For the Rich Text Box
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 13
    $richTextBoxSource.Location = $System_Drawing_Point
    $richTextBoxSource.Name = "richTextBoxSource"
    $richTextBoxSource.Font = New-Object System.Drawing.Font("Consolas New",10)
    $System_Drawing_Size = New-Object System.Drawing.Size
    # Set Rich Text Box Size
    $System_Drawing_Size.Height = 401
    $System_Drawing_Size.Width = 638
    $richTextBoxSource.Size = $System_Drawing_Size
    $richTextBoxSource.DetectUrls = $False
    $richTextBoxSource.ReadOnly = $True
    # Get source from script file and add newline to each array item for formatting
    $richTextBoxSource.Text = Get-Content $PSCommandPath | ForEach-Object {$_ + "`n"}
    # Add Rich text Box to View Source Form
    $formSourceCode.Controls.Add($richTextBoxSource)
    # Show View Source Form
    $formSourceCode.Show() | Out-Null
}

Function Create-ResultsForm($resultsMsg) {
     # Form for viewing the Results information
     # Add objects for Results Form
     $formResults = New-Object System.Windows.Forms.Form
     $richTextBoxResults = New-Object System.Windows.Forms.RichTextBox
     $System_Drawing_Size = New-Object System.Drawing.Size
     # Set Size of Reuslts Form
     $System_Drawing_Size.Height = 600
     $System_Drawing_Size.Width = 620
     $formResults.ClientSize = $System_Drawing_Size
     $formResults.DataBindings.DefaultDataSourceUpdateMode = 0
     $formResults.StartPosition = "CenterScreen"
     $formResults.Name = "formResults"
     # Set View Source Form Title
     $formResults.Text = "Send-Message Results"
     $formResults.Icon = [System.IconExtractor]::Extract("imageres.dll", 76, $False)
     # Add Rich Text Box for the Results Text to display in
     $richTextBoxResults.Anchor = 15
     $richTextBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
     $System_Drawing_Point = New-Object System.Drawing.Point
     #Set Location For the Rich Text Box
     $System_Drawing_Point.X = 5
     $System_Drawing_Point.Y = 5
     $richTextBoxResults.Location = $System_Drawing_Point
     $richTextBoxResults.Name = "richTextBoxError"
     $richTextBoxResults.Font = New-Object System.Drawing.Font("Courier New", 10)
     $System_Drawing_Size = New-Object System.Drawing.Size
     # Set Rich Text Box Size
     $System_Drawing_Size.Height = 590
     $System_Drawing_Size.Width = 610
     $richTextBoxResults.Size = $System_Drawing_Size
     $richTextBoxResults.DetectUrls = $False
     $richTextBoxResults.ReadOnly = $True
     #Add Results Message Text
     $richTextBoxResults.Text = $resultsMsg
     #Add Rich text Box to View Source Form
     $formResults.Controls.Add($richTextBoxResults)
     # Show Results Form
     $formResults.Show() | Out-Null
}

Function Create-HelpForm {
     # Build Help Form
     $formDirections = New-Object System.Windows.Forms.Form
     $richTextBoxHelp = New-Object System.Windows.Forms.RichTextBox
     # Add Objects to Help form
     $formDirections.AutoScroll = $True
     $System_Drawing_Size = New-Object System.Drawing.Size
     # Set Size of help Form
     $System_Drawing_Size.Height = 670
     $System_Drawing_Size.Width = 494
     $formDirections.ClientSize = $System_Drawing_Size
     $formDirections.DataBindings.DefaultDataSourceUpdateMode = 0
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 700
     $System_Drawing_Size.Width = 680
     $formDirections.MaximumSize = $System_Drawing_Size
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 700
     $System_Drawing_Size.Width = 680
     $formDirections.MinimumSize = $System_Drawing_Size
     $formDirections.Name = "formDirections"
     $formDirections.StartPosition = 1
     # Set Help Form title
     $formDirections.Text = "Help"
     $formDirections.FormBorderStyle = "FixedSingle"
     $formDirections.Icon = [System.IconExtractor]::Extract("imageres.dll", 94, $False)
     $formDirections.MaximizeBox = $False
     # Add Rich Text Box for the Help Text to display in
     $richTextBoxHelp.Anchor = 15
     $richTextBoxHelp.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 240, 240)
     $richTextBoxHelp.BorderStyle = 0
     $richTextBoxHelp.DataBindings.DefaultDataSourceUpdateMode = 0
     $System_Drawing_Point = New-Object System.Drawing.Point
     # Set Loctation for Rich Text Box
     $System_Drawing_Point.X = 13
     $System_Drawing_Point.Y = 13
     $richTextBoxHelp.Location = $System_Drawing_Point
     $richTextBoxHelp.Name = "richTextBoxHelp"
     $richTextBoxHelp.Font = New-Object System.Drawing.Font("Courier New", 10)
     $richTextBoxHelp.ReadOnly = $True
     $richTextBoxHelp.SelectionProtected = $True
     $richTextBoxHelp.Cursor = [System.Windows.Forms.Cursors]::Default
     $System_Drawing_Size = New-Object System.Drawing.Size
     # Set Size for Rich Text
     $System_Drawing_Size.Height = 670
     $System_Drawing_Size.Width = 469
     $richTextBoxHelp.Size = $System_Drawing_Size
     $richTextBoxHelp.TabIndex = 0
     $richTextBoxHelp.TabStop = $False
     # Set Text For Help form into the Rich Text Box
     $richTextBoxHelp.Text = 'SEND MESSAGE HELP

INTRODUCTION

This script was designed to send popup messages to computers over the network. This used to be done with net send which was removed from Windows. In its place, this script uses msg.exe, which was designed to send popup messages to users logged in to the computer.

SENDING A MESSAGE TO A SINGLE COMPUTER

To send a message to a single computer, enter your message in the first text box under the Send Messages tab, then type the hostname of the computer you wish to send a message in the single-line text box below and then click “Send” or press <Enter>. A results window will display after the script completes.

SENDING A MESSAGE TO MULTIPLE COMPUTERS

To send a message to multiple computers, enter your message in the first text box under the Send Messages tab, then choose on of the default list group checkboxes, type the hostnames of the computers you wish to send a message (separated by commas, spaces, or semi-colons), or you may also click “Import File” and choose a text (*.txt) file with hostnames/IPs of computers already stored (separated by commas, spaces, semi-colons, or carriage returns). The text file will populate the text box and you may add to the list if you so choose. Click “Send” or press <Enter> when you are ready to send the messages. A results window will display after the script completes.

VIEWING SCRIPT SOURCE CODE

You may pull up the source code of the script by choosing Help -> View Script from the menu or by pressing Ctrl + E.'

     <# Actions when clicking of links in help document (NOT USED)
     $richTextBoxHelp.add_LinkClicked(
        {
            Invoke-Expression "start $($_.LinkText)"
        }
     ) #>
     # Add Rich Text box to Help Form
     $formDirections.Controls.Add($richTextBoxHelp)
     # Show Help Form
     $formDirections.Show() | Out-Null
}

#region ICONExtractor (ThirdParty Code)
# Wrapper for VB code calling the ExtractIconEX function from the Windows API
# for extracting icons from .dll, .exe, etc.
# Obtained verbatim from Kazun's post at -
# http://social.technet.microsoft.com/Forums/en/winserverpowershell/thread/16444c7a-ad61-44a7-8c6f-b8d619381a27
$codeIconExtract = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@

#Add Type to use wrapped Static function for icon extraction
Add-Type -TypeDefinition $codeIconExtract -ReferencedAssemblies System.Drawing

#endregion ICONExtractor (ThirdParty Code)

$OnLoadForm_StateCorrection = {
     # Correct the initial state of the form to prevent the .Net maximized form issue
     $sendForm.WindowState = $initialFormWindowState
}

#endregion Script Functions

#----------------------------
#         Form Build
#----------------------------

#region Build Main Form

# Build Main Background and base properties of the Form
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of form
$System_Drawing_Size.Height = 500
$System_Drawing_Size.Width = 453
$sendForm.ClientSize = $System_Drawing_Size
$sendForm.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets minimumm Size of form (Cannot be resized lower than this)
$System_Drawing_Size.Height = 530
$System_Drawing_Size.Width = 469
$sendForm.MinimumSize = $System_Drawing_Size
# Form Setttings
$sendForm.Icon = [System.IconExtractor]::Extract("imageres.dll", 15, $False)
$sendForm.StartPosition = "CenterScreen"
$sendForm.Name = "sendForm"
$sendForm.Text = "Workstation Messaging System"
$sendForm.AcceptButton = $buttonSend
$sendForm.CancelButton = $buttonClose

#-------------------------------------------

# Build Processing Bar, placed next to tab name label
$processingLabel.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Text
$System_Drawing_Point.X = 120
$System_Drawing_Point.Y = 22
$processingLabel.Location = $System_Drawing_Point
$processingLabel.Name = "processingLabel"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Text Area where text will be
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 319
$processingLabel.Size = $System_Drawing_Size
$processingLabel.BackColor = "Green"
$processingLabel.ForeColor = "White"
$processingLabel.TextAlign = "MiddleLeft"
$processingLabel.TabIndex = 0
# Sets Text to Show
$processingLabel.Text = ""
# Add text to Send Message Area
$processingLabel.Hide()
$sendForm.Controls.Add($processingLabel)

#endregion Build Main Form

#region Build Top Menu

#region Build File Menu

# Add a MenuStrip object for Quitting the form
$menuFileQuit.Text = "&Quit"
$menuFileQuit.ShortcutKeys = "Control, Q"
$menuFileQuit.add_Click( {
          $sendForm.Close()
     }
)
# Add a MenuStrip object for Importing a text file to import computers names from
$menuFileOpen.Text = "&Import List"
$menuFileOpen.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::O
# Actions for the Import Menu
$menuFileOpen.add_Click( {
          $btnChosen = $openFileDialog1.ShowDialog()
          # Sends the chosen text file to be converted to a string
          if ($btnChosen -eq "OK") {
               Parse-TextFile -Path $openFileDialog1.FileName -TextBox $textBoxManEntListComp
          }
     }
)
# Builds File Menu Dropdown
$menuFile.Text = "&File"
$menuFile.DropDownItems.AddRange(@($menuFileOpen, $separatorF, $menuFileQuit))

#endregion Build File Menu

#region Build Help Menu

# Builds Help Menu Option and its actions to open the Help window
$menuHelpDirect.Text = "Send Message &Help"
$menuHelpDirect.ShortcutKeys = "F1"
# Action for the help directions Menu
$menuHelpDirect.add_Click( {
          Create-HelpForm
     }
)
# Disabled due to not functoning when script is compiled to executable by PS2EXE (Working)
# Builds the View Script Source menu option and its action to open the View Source window
$menuHelpView.Text = "Vi&ew Script"
$menuHelpView.ShortcutKeys = "Control, E"
# Actions for the View Script Menu
$menuHelpView.add_Click(
    {
        Create-ViewSourceForm
    }
)
# Builds the Help Menu Dropdown
$menuHelp.Text = "&Help"
$menuHelp.DropDownItems.AddRange(@($menuHelpDirect, $menuHelpView))

#endregion Build Help Menu

# Add both Menus to to the form
$menu.Items.AddRange(@($menuFile, $menuHelp))
$sendForm.Controls.Add($menu)

#endregion Build Top Menu

#region Build Tab Within Form

#region Build Main Tab Area

# Build Tab Anchor base
$tabControl.Anchor = 15
$tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 27
$tabControl.Location = $System_Drawing_Point
$tabControl.Name = "tabControl"
$tabControl.SelectedIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets size of Tab Anchor
$System_Drawing_Size.Height = 458
$System_Drawing_Size.Width = 428
$tabControl.Size = $System_Drawing_Size
$tabControl.TabIndex = 4
# Add Anchor to Form
$sendForm.Controls.Add($tabControl)

#-------------------------------------------

# Build Tab area to hold objects
$MsgTab.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 22
$MsgTab.Location = $System_Drawing_Point
$MsgTab.Name = "MsgTab"
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
$System_Windows_Forms_Padding.All = 3
$System_Windows_Forms_Padding.Bottom = 3
$System_Windows_Forms_Padding.Left = 3
$System_Windows_Forms_Padding.Right = 3
$System_Windows_Forms_Padding.Top = 3
$MsgTab.Padding = $System_Windows_Forms_Padding
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets size of Area
$System_Drawing_Size.Height = 422
$System_Drawing_Size.Width = 400
$MsgTab.Size = $System_Drawing_Size
$MsgTab.TabIndex = 0
$MsgTab.Text = "Send A Message"
$MsgTab.UseVisualStyleBackColor = $True
# Add tab to Form
$tabControl.Controls.Add($MsgTab)

#endregion Build Main Tab Area

#region Build Main Tab Buttons

# Builds Close Button and Adds to the Main tab
$buttonClose.Anchor = 2
$buttonClose.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Button on the Form
$System_Drawing_Point.X = 198
$System_Drawing_Point.Y = 395
$buttonClose.Location = $System_Drawing_Point
$buttonClose.Name = "buttonClose"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Button
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$buttonClose.Size = $System_Drawing_Size
$buttonClose.TabIndex = 8
# Sets Text of Button
$buttonClose.Text = "Close"
$buttonClose.UseVisualStyleBackColor = $True
$buttonClose.add_Click( {
          $sendForm.Close()
     }
)
# Add Close button to Form
$MsgTab.Controls.Add($buttonClose)

#-------------------------------------------

# Builds Send Button and adds to Main Tab
$buttonSend.Anchor = 2
$buttonSend.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Button on the Form
$System_Drawing_Point.X = 117
$System_Drawing_Point.Y = 395
$buttonSend.Location = $System_Drawing_Point
$buttonSend.Name = "buttonSend"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Button
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$buttonSend.Size = $System_Drawing_Size
$buttonSend.TabIndex = 7
# Sets Text of Button
$buttonSend.Text = "Send"
$buttonSend.UseVisualStyleBackColor = $True
# Actions for the "Send" button
$buttonSend.add_Click( {
          #Sends the contents of the text box to be stripped and turned into an array and then sent
          if ($textBoxManEntListComp.TextLength -eq 0) {
               Send-Message -message $richTextBoxMsg.Text -workstations $Script:targetList
               $Script:targetList = @()
               $checkBoxBearIP.Checked = $False
               $checkBoxPearIP.Checked = $False
               $checkBoxFigsIP.Checked = $False
               $checkBoxFireIP.Checked = $False
               $textBoxManEntListComp.clear()
               $processingLabel.hide()
          }
          else {
               Parse-Input -message $richTextBoxMsg.Text -workstations $textBoxManEntListComp.Text
               Send-Message -message $richTextBoxMsg.Text -workstations $Script:targetList
               $Script:targetList = @()
               $checkBoxBearIP.Checked = $False
               $checkBoxPearIP.Checked = $False
               $checkBoxFigsIP.Checked = $False
               $checkBoxFireIP.Checked = $False
               $textBoxManEntListComp.Clear()
               $processingLabel.hide()
          }
     }
)
# Add Send button to Form
$MsgTab.Controls.Add($buttonSend)

#endregion Build Main Tab Buttons

#region Message Entry Area

# Adds Text to Message Entry Area
$labelMsg.Anchor = 13
$labelMsg.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Text
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 14
$labelMsg.Location = $System_Drawing_Point
$labelMsg.Name = "labelMsg"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Text Area where text will be
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 378
$labelMsg.Size = $System_Drawing_Size
$labelMsg.TabIndex = 0
# Sets Text to Show
$labelMsg.Text = "Type Your Message (Max 255 Character Limit)"
# Add text to Send Message Area
$MsgTab.Controls.Add($labelMsg)

# Adds Rich Text Box to Send Message Area
$richTextBoxMsg.Anchor = 13
$richTextBoxMsg.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Size of Rich Text Box
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 40
$richTextBoxMsg.Location = $System_Drawing_Point
$richTextBoxMsg.Name = "richTextBoxMsg"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Location of Rich Text Box
$System_Drawing_Size.Height = 110
$System_Drawing_Size.Width = 378
$richTextBoxMsg.Size = $System_Drawing_Size
$richTextBoxMsg.TabIndex = 1
# Sets Default Text of Rich Text Box if Desired
$richTextBoxMsg.Text = ""
# Adds Rich Text Box to Send Message Area
$MsgTab.Controls.Add($richTextBoxMsg)

#endregion Message Entry Area

#region Computer List Group Area

# Builds "Send to Computers" List Box Area as a group (Surrounded by Rectangle with Title)
$grpListComp.Anchor = 13
$grpListComp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of the Group Area
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 165
$grpListComp.Location = $System_Drawing_Point
$grpListComp.Name = "grpListComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of group Area
$System_Drawing_Size.Height = 235
$System_Drawing_Size.Width = 378
$grpListComp.Size = $System_Drawing_Size
$grpListComp.TabIndex = 5
$grpListComp.TabStop = $False
# Sets Title of Group Area
$grpListComp.Text = "List Workstations To Receive Message"
# Adds Group Area to Tab
$MsgTab.Controls.Add($grpListComp)

#region Auto Entry list Computers Group Area

# Adds Auto Entry Label to the List Computers group Area
$autoEntListComp.Anchor = 13
$autoEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
$Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Underline)
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Text
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 22
$autoEntListComp.Location = $System_Drawing_Point
$autoEntListComp.Name = "autoEntListComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Text Area where text will be
$System_Drawing_Size.Height = 14
$System_Drawing_Size.Width = 366
$autoEntListComp.Size = $System_Drawing_Size
$autoEntListComp.TabIndex = 0
$autoEntListComp.Font = $Font
# Sets Text to show
$autoEntListComp.Text = "Default Workstation Listings"
# Add Text to List Group Area
$grpListComp.Controls.Add($autoEntListComp)

#-------------------------------------------

# Adds Text instructions to the Auto Entry List Computers group Area
$labelAutoEntListComp.Anchor = 13
$labelAutoEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Text
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 41
$labelAutoEntListComp.Location = $System_Drawing_Point
$labelAutoEntListComp.Name = "labelAutoEntListComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Text Area where text will be
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 366
$labelAutoEntListComp.Size = $System_Drawing_Size
$labelAutoEntListComp.TabIndex = 0
# Sets Text to show
$labelAutoEntListComp.Text = "Choose a list(s) of Workstations to send the message to:"
# Add Text to List Group Area
$grpListComp.Controls.Add($labelAutoEntListComp)

#-------------------------------------------

# Adds Bear IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxBearIP.Anchor = 13
$checkBoxBearIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 14
$System_Drawing_Point.Y = 60
$checkBoxBearIP.Location = $System_Drawing_Point
$checkBoxBearIP.Name = "checkBoxBearIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 150
$checkBoxBearIP.Size = $System_Drawing_Size
$checkBoxBearIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxBearIP.Text = "Bear Workstations"
$checkBoxBearIP.UseVisualStyleBackColor = $True
$checkBoxBearIP.Add_Click( {
          Set-TargetVariable -ipList $bearIPs -CheckBox $checkBoxBearIP
     })

# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxBearIP)

#-------------------------------------------

# Adds Pear IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxPearIP.Anchor = 13
$checkBoxPearIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 14
$System_Drawing_Point.Y = 90
$checkBoxPearIP.Location = $System_Drawing_Point
$checkBoxPearIP.Name = "checkBoxPearIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 150
$checkBoxPearIP.Size = $System_Drawing_Size
$checkBoxPearIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxPearIP.Text = "Pear Workstations"
$checkBoxPearIP.UseVisualStyleBackColor = $True
$checkBoxPearIP.Add_Click( {
          Set-TargetVariable -ipList $pearsIPs -CheckBox $checkBoxPearIP
     })
# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxPearIP)

#-------------------------------------------

# Adds Figs IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxFigsIP.Anchor = 13
$checkBoxFigsIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 194
$System_Drawing_Point.Y = 60
$checkBoxFigsIP.Location = $System_Drawing_Point
$checkBoxFigsIP.Name = "checkBoxFigsIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 150
$checkBoxFigsIP.Size = $System_Drawing_Size
$checkBoxFigsIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxFigsIP.Text = "Figs Workstations"
$checkBoxFigsIP.UseVisualStyleBackColor = $True
$checkBoxFigsIP.Add_Click( {
          Set-TargetVariable -ipList $figsIPs -CheckBox $checkBoxFigsIP
     })
# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxFigsIP)

#-------------------------------------------

# Adds Fire IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxFireIP.Anchor = 13
$checkBoxFireIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 194
$System_Drawing_Point.Y = 90
$checkBoxFireIP.Location = $System_Drawing_Point
$checkBoxFireIP.Name = "checkBoxFireIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 200
$checkBoxFireIP.Size = $System_Drawing_Size
$checkBoxFireIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxFireIP.Text = "Fire Workstations"
$checkBoxFireIP.UseVisualStyleBackColor = $True
$checkBoxFireIP.Add_Click( {
          Set-TargetVariable -ipList $fireIPs -CheckBox $checkBoxFireIP
     })
# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxFireIP)

#endregion Auto Entry list Computers Group Area

#region Manual Entry List Computers Group Area

# Adds Manual Entry Label to the List Computers group Area
$manEntListComp.Anchor = 13
$manEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
$Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Underline)
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Text
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 125
$manEntListComp.Location = $System_Drawing_Point
$manEntListComp.Name = "labelListComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Text Area where text will be
$System_Drawing_Size.Height = 14
$System_Drawing_Size.Width = 366
$manEntListComp.Size = $System_Drawing_Size
$manEntListComp.TabIndex = 0
$manEntListComp.Font = $Font
# Sets Text to show
$manEntListComp.Text = "Manual Workstation Entry"
# Add Text to Area
$grpListComp.Controls.Add($manEntListComp)

#-------------------------------------------

# Adds Text instructions to the Manual Entry List Computers group Area
$labelManEntListComp.Anchor = 13
$labelManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Text
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 145
$labelManEntListComp.Location = $System_Drawing_Point
$labelManEntListComp.Name = "labelManEntListComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Text Area where text will be
$System_Drawing_Size.Height = 28
$System_Drawing_Size.Width = 366
$labelManEntListComp.Size = $System_Drawing_Size
$labelManEntListComp.TabIndex = 0
# Sets Text to show
$labelManEntListComp.Text = "Type the computer name(s)/IP(s) separated by spaces, commas, semi-colons, or import a list of names from a text file:"
# Add Text to Area
$grpListComp.Controls.Add($labelManEntListComp)

#-------------------------------------------

# Add Text Box to Manual Entry List Computers Group Area
$textBoxManEntListComp.Anchor = 13
$textBoxManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Text Box
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 181
$textBoxManEntListComp.Location = $System_Drawing_Point
$textBoxManEntListComp.Name = "textBoxManEntListComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Text Box
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 366
$textBoxManEntListComp.Size = $System_Drawing_Size
$textBoxManEntListComp.TabIndex = 1
# Adds Text Box to Area
$grpListComp.Controls.Add($textBoxManEntListComp)

#-------------------------------------------

# Add Import File Button to List Computers Group Area
$buttonManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Button
$System_Drawing_Point.X = 7
$System_Drawing_Point.Y = 205
$buttonManEntListComp.Location = $System_Drawing_Point
$buttonManEntListComp.Name = "buttonManEntListComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Button
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$buttonManEntListComp.Size = $System_Drawing_Size
$buttonManEntListComp.TabIndex = 2
# Sets Title of Button
$buttonManEntListComp.Text = "Import List"
$buttonManEntListComp.UseVisualStyleBackColor = $True
# Actions for the "Import File" Button
$buttonManEntListComp.add_Click( {
          $btnChosen = $openFileDialog1.ShowDialog()
          #Sends the chosen text file to be converted to a string
          if ($btnChosen -eq "OK") {
               Parse-TextFile -Path $openFileDialog1.FileName -TextBox $textBoxManEntListComp
          }
     }
)
# Adds Import File Button to Area
$grpListComp.Controls.Add($buttonManEntListComp)

#-------------------------------------------

# Build Clear Text Box Button to List Computers Group Area
$buttonManEntClearComp.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets Location of Button
$System_Drawing_Point.X = 89
$System_Drawing_Point.Y = 205
$buttonManEntClearComp.Location = $System_Drawing_Point
$buttonManEntClearComp.Name = "buttonManEntClearComp"
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of Buttons
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$buttonManEntClearComp.Size = $System_Drawing_Size
$buttonManEntClearComp.TabIndex = 3
# Sets Button Text
$buttonManEntClearComp.Text = "Clear Text"
$buttonManEntClearComp.UseVisualStyleBackColor = $True
# Actions for the Clear Button (Removes all text from the text box)
$buttonManEntClearComp.add_Click( {
          $textBoxManEntListComp.Clear()
     }
)
# Add Clear Button to Area
$grpListComp.Controls.Add($buttonManEntClearComp)

#endregion Manual Entry List Computers Group Area

#endregion Computer List Group Area

#endregion Build Tab Within Form

#----------------------------
#      Main Form Actions
#----------------------------

#region Form Actions

Check-IfAdmin

# Settings for the Open File Dialog when opening a text file
$openFileDialog1.Filter = "Text Files (*.txt) | *.txt"
$openFileDialog1.ShowHelp = $True

# Save the initial state of the form
$InitialFormWindowState = $sendForm.WindowState
# Init the OnLoad event to correct the initial state of the form
$sendForm.add_Load($OnLoadForm_StateCorrection)

# Display correctly outside of ISE
[System.Windows.Forms.Application]::EnableVisualStyles()

# Show the Form
$sendForm.Add_Shown( {$sendForm.Activate()})
$sendForm.ShowDialog($this) | Out-Null

#endregion Form Actions