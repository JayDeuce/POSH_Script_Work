#-----------------------------
#       Main Form Setup
#-----------------------------

#--> Set Error Action Preference
$ErrorActionPreference = "SilentlyContinue"

#--> Global Variables
[array]$targetList # Initialize Target List variable
$OnLoadForm_StateCorrection = {$functionForm.WindowState = $initialFormWindowState} # Correct the initial state of the form to prevent the .Net maximized form issue
$openFileDialog.Filter = "Text Files (*.txt) | *.txt" # Settings for the Open File Dialog when opening a text file

#--> Import the necessary Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

#--> Set Form Objects Variables
# Look here for all objects that windows forms can create here:
# https://docs.microsoft.com/en-us/dotnet/framework/winforms/controls/windows-forms-controls-by-function
# Add variable equal to a new object creation for the form object you are creating
$functionForm = New-Object System.Windows.Forms.Form # Main Form Variable
# Objects for Menu Bar - Each object below builds the menu at the top of the form
$menu = New-Object System.Windows.Forms.MenuStrip # Creates the Menu Strip the menu items sit in
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem # Root File Menu Text
$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem # Root Help Menu text
$menuFileOpen = New-Object System.Windows.Forms.ToolStripMenuItem # Open File Menu Option
$menuFileQuit = New-Object System.Windows.Forms.ToolStripMenuItem # Quit Menu Option
$menuHelpInfo = New-Object System.Windows.Forms.ToolStripMenuItem # Main Script Help Menu Option
$menuHelpView = New-Object System.Windows.Forms.ToolStripMenuItem # View Script Source Menu option
$separatorF = New-Object System.Windows.Forms.ToolStripSeparator # Seperator Line in the File Menu
# Add Other Objects to Form
$tabControl = New-Object System.Windows.Forms.TabControl # Initialize Tab Area
$firstTab = New-Object System.Windows.Forms.TabPage # Initialize  First Tab in the Tab Area, More can be added and then define in a new section below
$buttonClose = New-Object System.Windows.Forms.Button # Initialize Main Form Close Button
$buttonSend = New-Object System.Windows.Forms.Button # Initialize Main Form Send Button (Name shoudl be change to whatever the button text is named to)
$processingLabel = New-Object System.Windows.Forms.Label # Initialize Processing Label (Task Processing bar)
$textBoxManEntListComp = New-Object System.Windows.Forms.RichTextBox # Initialize Computer Manual Entry Text Box
$buttonManEntClearComp = New-Object System.Windows.Forms.Button # Initialize Computer Manual Entry Clear Button
$buttonManEntImportComp = New-Object System.Windows.Forms.Button # Initialize Computer Manual Entry Import Button
$labelManEntListComp = New-Object System.Windows.Forms.Label # Initialize Label for Computer Manual Entry text box
$descrLablManEntListComp = New-Object System.Windows.Forms.Label # Initialize Description for Computer manual Entry text box
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog # Initialize Open File Dialog
$initialFormWindowState = New-Object System.Windows.Forms.FormWindowState # Initialize Form Window State (Must be present)

#---------------------------------
#       Main Function Setup
#---------------------------------

Function Check-IfAdmin {
     # Check if the active user is an admin or not, Post requirements message if not and closes script
     $myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
     $myWindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
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
function Create-MainForm {
     #--> Build Main Form
     # Build Main Background and base properties of the Form
      # Sets Size of form
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 500
     $System_Drawing_Size.Width = 453
     $functionForm.ClientSize = $System_Drawing_Size
     $functionForm.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets minimumm Size of form (Cannot be resized lower than this)
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 530
     $System_Drawing_Size.Width = 469
     $functionForm.MinimumSize = $System_Drawing_Size
     # Form Setttings
     $functionForm.Icon = [System.IconExtractor]::Extract("imageres.dll", 63, $False)
     $functionForm.StartPosition = "CenterScreen"
     $functionForm.Name = "functionForm"
     $functionForm.Text = "Function Form Title"
     $functionForm.AcceptButton = $buttonSend
     $functionForm.CancelButton = $buttonClose

     #-------------------------------------------

     # Build Processing Bar, placed next to tab name label
     $processingLabel.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 120
     $System_Drawing_Point.Y = 22
     $processingLabel.Location = $System_Drawing_Point
     $processingLabel.Name = "processingLabel"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 18
     $System_Drawing_Size.Width = 319
     $processingLabel.Size = $System_Drawing_Size
     $processingLabel.BackColor = "Green"
     $processingLabel.ForeColor = "White"
     $processingLabel.TextAlign = "MiddleLeft"
     $processingLabel.TabIndex = 0
     # Sets Text to Show
     $processingLabel.Text = ""
     # Add text to Area
     $processingLabel.Hide()
     $functionForm.Controls.Add($processingLabel)

     #--> Build Top Menu

     #    --> Build File Menu

     # Add a MenuStrip object for Quitting the form
     $menuFileQuit.Text = "&Quit"
     $menuFileQuit.ShortcutKeys = "Control, Q"
     $menuFileQuit.add_Click( {
               $functionForm.Close()
          }
     )
     # Add a MenuStrip object for Importing a text file to import computers names from
     $menuFileOpen.Text = "&Import List"
     $menuFileOpen.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::O
     # Actions for the Import Menu
     $menuFileOpen.add_Click( {
               $btnChosen = $openFileDialog.ShowDialog()
               # Sends the chosen text file to be converted to a string
               if ($btnChosen -eq "OK") {
                    Parse-TextFile -Path $openFileDialog.FileName -TextBox $textBoxManEntListComp
               }
          }
     )
     # Builds File Menu Dropdown
     $menuFile.Text = "&File"
     $menuFile.DropDownItems.AddRange(@($menuFileOpen, $separatorF, $menuFileQuit))

     #    --> Build Help Menu

     # Builds Help Menu Option and its actions to open the Help window
     $menuHelpInfo.Text = "Function &Help"
     $menuHelpInfo.ShortcutKeys = "F1"
     # Action for the help directions Menu
     $menuHelpInfo.add_Click( {
               Create-HelpForm
          }
     )
     # Builds the View Script Source menu option and its action to open the View Source window
     $menuHelpView.Text = "Vi&ew Script"
     $menuHelpView.ShortcutKeys = "Control, E"
     # Actions for the View Script Menu
     $menuHelpView.add_Click( {
               Create-ViewSourceForm
          }
     )
     # Builds the Help Menu Dropdown
     $menuHelp.Text = "&Help"
     $menuHelp.DropDownItems.AddRange(@($menuHelpInfo, $menuHelpView))
     # Add both Menus to to the form
     $menu.Items.AddRange(@($menuFile, $menuHelp))
     $functionForm.Controls.Add($menu)

     #--> Build Main Tab Area

     # Build Tab Anchor base
     $tabControl.Anchor = 15
     $tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 7
     $System_Drawing_Point.Y = 27
     $tabControl.Location = $System_Drawing_Point
     $tabControl.Name = "tabControl"
     $tabControl.SelectedIndex = 0
     # Sets size of Tab Anchor
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 468
     $System_Drawing_Size.Width = 440
     $tabControl.Size = $System_Drawing_Size
     $tabControl.TabIndex = 4
     # Add Anchor to Form
     $functionForm.Controls.Add($tabControl)

     #-------------------------------------------

     # Build Tab area to hold objects
     $firstTab.DataBindings.DefaultDataSourceUpdateMode = 0
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 4
     $System_Drawing_Point.Y = 22
     $firstTab.Location = $System_Drawing_Point
     $firstTab.Name = "firstTab"
     $System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding
     $System_Windows_Forms_Padding.All = 3
     $System_Windows_Forms_Padding.Bottom = 3
     $System_Windows_Forms_Padding.Left = 3
     $System_Windows_Forms_Padding.Right = 3
     $System_Windows_Forms_Padding.Top = 3
     $firstTab.Padding = $System_Windows_Forms_Padding
     # Sets size of Area
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 422
     $System_Drawing_Size.Width = 400
     $firstTab.Size = $System_Drawing_Size
     $firstTab.TabIndex = 0
     $firstTab.Text = "Function Name"
     $firstTab.UseVisualStyleBackColor = $True
     # Add tab to Form
     $tabControl.Controls.Add($firstTab)

     # Builds Close Button and Adds to the Main tab
     $buttonClose.Anchor = 2
     $buttonClose.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button on the Form
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 198
     $System_Drawing_Point.Y = 395
     $buttonClose.Location = $System_Drawing_Point
     $buttonClose.Name = "buttonClose"
     # Sets Size of Button
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonClose.Size = $System_Drawing_Size
     $buttonClose.TabIndex = 8
     # Sets Text of Button
     $buttonClose.Text = "Close"
     $buttonClose.UseVisualStyleBackColor = $True
     # Sets action on click (Can be any function, cmdlet, .NET Method, logic control, etc.)
     $buttonClose.add_Click( {
               $functionForm.Close()
          }
     )
     # Add Close button to Form
     $firstTab.Controls.Add($buttonClose)

     #-------------------------------------------

     # Builds Send Button and adds to Main Tab
     $buttonSend.Anchor = 2
     $buttonSend.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button on the Form
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 117
     $System_Drawing_Point.Y = 395
     $buttonSend.Location = $System_Drawing_Point
     $buttonSend.Name = "buttonSend"
     # Sets Size of Button
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonSend.Size = $System_Drawing_Size
     $buttonSend.TabIndex = 7
     # Sets Text of Button
     $buttonSend.Text = "rename"
     $buttonSend.UseVisualStyleBackColor = $True
     # Actions for the "Send" button
     $buttonSend.add_Click( {
               #Sends the contents of the text box to be stripped and turned into an array
               Parse-Input $textBoxManEntListComp.Text
               $Script:targetList = @()
               $textBoxManEntListComp.clear()
               $processingLabel.hide()
          }
     )
     # Add Send button to Form
     $firstTab.Controls.Add($buttonSend)

     #--> Computers Manual Device Name/IP Area

     # Adds Manual Entry Label to the Computers Manual Device Name/IP Area
     $labelManEntListComp.Anchor = 13
     $labelManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
     $Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Underline)
     # Sets Location of Text
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 20
     $labelManEntListComp.Location = $System_Drawing_Point
     $labelManEntListComp.Name = "labelListComp"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 14
     $System_Drawing_Size.Width = 366
     $labelManEntListComp.Size = $System_Drawing_Size
     $labelManEntListComp.TabIndex = 0
     $labelManEntListComp.Font = $Font
     # Sets Text to show
     $labelManEntListComp.Text = "Device Name/IP Entry"
     # Add Text to Area
     $firstTab.Controls.Add($labelManEntListComp)

     #-------------------------------------------

     # Adds Text instructions to the Computers Manual Device Name/IP Area
     $descrLablManEntListComp.Anchor = 13
     $descrLablManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 40
     $descrLablManEntListComp.Location = $System_Drawing_Point
     $descrLablManEntListComp.Name = "descrLablManEntListComp"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 28
     $System_Drawing_Size.Width = 366
     $descrLablManEntListComp.Size = $System_Drawing_Size
     $descrLablManEntListComp.TabIndex = 0
     # Sets Text to show
     $descrLablManEntListComp.Text = "Type the workstation name(s)/IP(s) separated by spaces, commas, semi-colons, or import a list of names from a text file:"
     # Add Text to Area
     $firstTab.Controls.Add($descrLablManEntListComp)

     #-------------------------------------------

     # Add Text Box to the Manual Entry Computers Manual Device Name/IP Area
     $textBoxManEntListComp.Anchor = 13
     $textBoxManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text Box
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 70
     $textBoxManEntListComp.Location = $System_Drawing_Point
     $textBoxManEntListComp.Name = "textBoxManEntListComp"
     # Sets Size of Text Box
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 20
     $System_Drawing_Size.Width = 366
     $textBoxManEntListComp.Size = $System_Drawing_Size
     $textBoxManEntListComp.TabIndex = 1
     # Adds Text Box to Area
     $firstTab.Controls.Add($textBoxManEntListComp)

     #-------------------------------------------

     # Add Import File Button to Computers Manual Device Name/IP Area
     $buttonManEntImportComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 7
     $System_Drawing_Point.Y = 95
     $buttonManEntImportComp.Location = $System_Drawing_Point
     $buttonManEntImportComp.Name = "buttonManEntImportComp"
     # Sets Size of Button
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntImportComp.Size = $System_Drawing_Size
     $buttonManEntImportComp.TabIndex = 2
     # Sets Title of Button
     $buttonManEntImportComp.Text = "Import List"
     $buttonManEntImportComp.UseVisualStyleBackColor = $True
     # Actions for the "Import File" Button
     $buttonManEntImportComp.add_Click( {
               $btnChosen = $openFileDialog.ShowDialog()
               #Sends the chosen text file to be converted to a string
               if ($btnChosen -eq "OK") {
                    Parse-TextFile -Path $openFileDialog.FileName -TextBox $textBoxManEntListComp
               }
          }
     )
     # Adds Import File Button to Area
     $firstTab.Controls.Add($buttonManEntImportComp)

     #-------------------------------------------

     # Build Clear Text Box Button to the Computers Manual Device Name/IP Area
     $buttonManEntClearComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 89
     $System_Drawing_Point.Y = 95
     $buttonManEntClearComp.Location = $System_Drawing_Point
     $buttonManEntClearComp.Name = "buttonManEntClearComp"
     # Sets Size of Buttons
     $System_Drawing_Size = New-Object System.Drawing.Size
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
     $firstTab.Controls.Add($buttonManEntClearComp)

     #--> Additonal Objects

     # Add Additiontal objects you wish to add to the form, usually based on script parameters

     $functionForm.Add_Shown( {$functionForm.Activate()})
     $functionForm.ShowDialog($this) | Out-Null
}
Function Create-ViewSourceForm {
     # Builds View Source Form from the help menu "View Source"
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
     $formSourceCode.Text = "Install Update Script Source View"
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
     $richTextBoxSource.Font = New-Object System.Drawing.Font("Consolas New", 10)
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
Function Create-HelpForm {
     # Build Help Form
     $formHelp = New-Object System.Windows.Forms.Form
     $richTextBoxHelp = New-Object System.Windows.Forms.RichTextBox
     # Add Objects to Help form
     $formHelp.AutoScroll = $True
     # SetHelp Size of help Form
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 600
     $System_Drawing_Size.Width = 500
     $formHelp.ClientSize = $System_Drawing_Size
     $formHelp.DataBindings.DefaultDataSourceUpdateMode = 0
     $formHelp.MinimumSize = $System_Drawing_Size
     $formHelp.Name = "helpForm"
     $formHelp.StartPosition = 1
     # SetHelp Help Form title
     $formHelp.Text = "Help"
     $formHelp.FormBorderStyle = "FixedSingle"
     $formHelp.Icon = [System.IconExtractor]::Extract("imageres.dll", 94, $False)
     $formHelp.MaximizeBox = $False
     # Add Rich Text Box for the Help Text to display in
     $richTextBoxHelp.Anchor = 15
     $richTextBoxHelp.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 240, 240)
     $richTextBoxHelp.BorderStyle = 0
     $richTextBoxHelp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Set Loctation for Rich Text Box
     $System_Drawing_Point = New-Object System.Drawing.Point
     $System_Drawing_Point.X = 13
     $System_Drawing_Point.Y = 13
     $richTextBoxHelp.Location = $System_Drawing_Point
     $richTextBoxHelp.Name = "richTextBoxHelp"
     $richTextBoxHelp.Font = New-Object System.Drawing.Font("Courier New", 10)
     $richTextBoxHelp.ReadOnly = $True
     $richTextBoxHelp.SelectionProtected = $True
     $richTextBoxHelp.Cursor = [System.Windows.Forms.Cursors]::Default
     # Set Size for Rich Text
     $System_Drawing_Size = New-Object System.Drawing.Size
     $System_Drawing_Size.Height = 570
     $System_Drawing_Size.Width = 469
     $richTextBoxHelp.Size = $System_Drawing_Size
     $richTextBoxHelp.TabIndex = 0
     $richTextBoxHelp.TabStop = $False
     # Set Text For Help form into the Rich Text Box
     # Text must be between the '', and it must starts at the beggining of the line as show.
     # Type in the Help text you wish to show.
     $richTextBoxHelp.Text = 'INTRODUCTION

This script was designed to

TO A SINGLE COMPUTER


TO MULTIPLE COMPUTERS


VIEWING SCRIPT SOURCE CODE

You may pull up the source code of the script by choosing Help -> View Script from the menu or by pressing Ctrl + E.'

     # AddHelp Rich Text box to Help Form
     $formHelp.Controls.Add($richTextBoxHelp)
     # Show Help Form
     $formHelp.Show() | Out-Null
}

#--> ICONExtractor (ThirdParty Code)
# Wrapper for VB code calling the ExtractIconEX function from the Windows API
# for extracting icons from .dll, .exe, etc.
# Obtained verbatim from Kazun's post at -
# http://social.technet.microsoft.com/Forums/en/winserverpowershell/thread/16444c7a-ad61-44a7-8c6f-b8d619381a27
# ICON INDEX: https://diymediahome.org/windows-icons-reference-list-with-details-locations-images/
# If icon is image 64, enter 63, number is always one number below what is shown.
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

#-------------------------------
#       Form Main Actions
#-------------------------------

# Uncomment next line if Script must be run with Admin Rights.
#Check-IfAdmin

# Save the initial state of the form
$InitialFormWindowState = $functionForm.WindowState
# Init the OnLoad event to correct the initial state of the form
$functionForm.add_Load($OnLoadForm_StateCorrection)

# Display correctly outside of ISE
[System.Windows.Forms.Application]::EnableVisualStyles()

# Show the Form
Create-MainForm