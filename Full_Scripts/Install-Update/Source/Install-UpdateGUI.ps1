<#
.SYNOPSIS
     Pushes provided update to listed machines useing PSExec remote administration tool.

.DESCRIPTION
     Using passed parameters for the Computer name and and update/patch Name; this script
     will copy the update/patch to the designated Computer or Computers, run the installation of the update,
     and return the exit/error code of the windows updates service.

     !! THIS SCRIPT MUST BE RUN WITH ADMINISTRATOR RIGHTS !!

.NOTES
     Name: Install-Update (GUI)
     Author: Jonathan Durant
     Version: 1.0
     DateUpdated: 2017-06-20

.INPUTS
     Computer names, update/patch Names, file paths

.OUTPUTS
     Error/Exit codes to show completion status.
#>

#---------------------------------
#region Install-Update Form Setup
#-----------------------------
#--> Set Error Action Preference
$ErrorActionPreference = "SilentlyContinue"

#--> Global Variables
# Variables for Functions
[array]$computers = @() # Intialize Computer Namr(s) Variable
[string]$updatePatchName = "" # Intialize Update/Patch Name Variable
[string]$updatePatchPath = "C:\UpdatePush\Patches" # Initialize Update/Patch Folder Path Variable
[string]$psexecFilePath = "C:\UpdatePush\PSExec\PSExec.exe" # Initialize PSExec file path location Variable
[array]$exitCodeReport = @() # Initialize Exit Code Report Variable
[array]$exitCodeInfo = @() # Initialize PSExec (After install) Exit Code Gathering Variable
[string]$popupMessage = "" # Intialize Pop-up box Message variable
[string]$popupTitle = "" # Intialize Pop-up box Title variable
[string]$popupIcon = "" # Intialize Pop-up Box Icon Variable

#--> Import the Necessary .NET WInforms Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

# Display correctly outside of ISE
[System.Windows.Forms.Application]::EnableVisualStyles()

# Initialize Form Window State (Must be present)
$initialFormWindowState = New-Object System.Windows.Forms.FormWindowState # Initialize Form Window State (Must be present)

# Variables for use by all forms
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Windows_Forms_Padding = New-Object System.Windows.Forms.Padding

#--> ICONExtractor (ThirdParty Code)
# Wrapper for VB code calling the ExtractIconEX function from the Windows API for extracting icons from .dll, .exe, etc.
# Obtained http://social.technet.microsoft.com/Forums/en/winserverpowershell/thread/16444c7a-ad61-44a7-8c6f-b8d619381a27
# ICON INDEX: https://diymediahome.org/windows-icons-reference-list-with-details-locations-images/ (icon is image 64, enter 63, number is always one number below what is shown)
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
#---------------------------------
#endregion
#---------------------------------

#---------------------------------
#region  Install Update Function Setup
#---------------------------------
function Create-InfoMessageBox ([string]$message, [string]$title, [string]$icon) {

     [System.Windows.Forms.MessageBoxButtons]$buttons = "OK"
     [System.Windows.Forms.MessageBoxIcon]$icons = $icon

     [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icons) | Out-Null
}
function Check-IfAdmin {
     # Check if the active user is an admin or not, Post requirements message if not and closes script
     $myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
     $myWindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
     $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
     # Check to see if we are currently running "as Administrator" and message user to relaunch as admin if not.
     if (!($myWindowsPrincipal.IsInRole($adminRole))) {
          $popupMessage = "You are not running as an administrator, please rerun the program as ADMINISTRATOR!"
          $popupTitle = "ERROR: NOT AN ADMIN"
          $popupIcon = "Error"
          Create-InfoMessageBox -message $popupMessage -title $popupTitle -icon $popupIcon
          exit
     }
}
function Parse-TextFile ([string]$path, [System.Windows.Forms.RichTextBox]$textBox) {
     # Get text from file, combine into string and insert into text box
     $compNames = (Get-Content $path) -join ', '
     $TextBox.Text = $compNames
}
function Parse-Input ([string]$objects) {
     # Split inputed text into array of single entries delimiting on whitespace, commas, and semicolons
     $arrObj = (($objects -replace '[,;\s]', ' ').Trim()) -split '\s+'
     $Script:computers += $arrObj
}
function Install-Update ([array]$computers, [string]$updatePatchName, [string]$updatePatchPath, [String]$psexecFilePath) {

     # Set file extension variable based on the file passed for use by the switch function below
     $updatePatchFileExt = $updatePatchName.substring($updatePatchName.length - 3, 3)

     # Check File extension and decide what to do based on its value
     if ($updatePatchFileExt -eq "exe" -or $updatePatchFileExt -eq "msu" -or $updatePatchFileExt -eq "cab") {

          foreach ($computer in $computers) {

               # Initializes new PSObject for exit code gathering (Reset each loop)
               $Script:exitCodeInfo = New-Object -TypeName PSObject

               # Shows Progress Bar
               $processingLabel.Show()
               $processingLabel.Text = "Processing: $computer"

               if (!(test-connection -count 1 -Quiet -ComputerName $computer)) {
                    $Script:exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Computer Name' -Value $computer
                    $Script:exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Exit Code' -Value "System could not be contacted."
                    $Script:exitCodeReport += $Script:exitCodeInfo # Adds Exit Code to Exit Code Report variable
               }
               else {
                    if (!(Test-Path -Path "\\$computer\c$\Temp")) {
                         new-item -Path "\\$computer\c$\Temp" -ItemType directory | Out-Null
                    }
                    Copy-Item $updatePatchPath\$updatePatchName "\\$computer\c$\Temp"
                    switch ($updatePatchFileExt) {
                         exe {
                              & $psexecFilePath -s \\$computer "C:\Temp\$updatePatchName" /q /norestart
                              $exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Computer Name' -Value $computer
                              $exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Exit Code' -Value $LASTEXITCODE
                              $Script:exitCodeReport += $Script:exitCodeInfo # Adds Exit Code to Exit Code Report variable
                         }
                         msu {
                              & $psexecFilepath -s \\$computer wusa "C:\Temp\$updatePatchName" /quiet /norestart
                              $exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Computer Name' -Value $computer
                              $exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Exit Code' -Value $LASTEXITCODE
                              $Script:exitCodeReport += $Script:exitCodeInfo # Adds Exit Code to Exit Code Report variable
                         }
                         cab {
                              & $psexecFilePath -s \\$computer Dism.exe /Online /Add-Package /Quiet /NoRestart /packagepath:"C:\Temp\$updatePatchName"
                              $exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Computer Name' -Value $computer
                              $exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Exit Code' -Value $LASTEXITCODE
                              $Script:exitCodeReport += $Script:exitCodeInfo # Adds Exit Code to Exit Code Report variable
                         }
                    }
                    # Delete remote copy of update package
                    Remove-Item "\\$computer\c$\Temp\$updatePatchName"
               }
          }
          return
     }
     else {
          $Script:exitCodeInfo = New-Object -TypeName PSObject # Initializes new PSObject for exit code gathering (Just for this Else Statement)
          $exitCodeInfo | Add-Member -MemberType NoteProperty -Name 'Exit Code' -Value "$updatePatchName is of a unknown update filetype, please run manually on each computer."
          $Script:exitCodeReport += $exitCodeInfo # Adds Exit Code to Exit Code Report variable
          return
     }
}
function Create-MainForm {
     #--> Set Main Form Objects Variables
     # Look here for all objects that windows forms can create here:
     # https://docs.microsoft.com/en-us/dotnet/framework/winforms/controls/windows-forms-controls-by-function
     # Add variable equal to a new object created for the form object you are creating

     $installForm = New-Object System.Windows.Forms.Form # Main Form Variable
     # Objects for Menu Bar - Each object below builds the menu at the top of the form
     $menu = New-Object System.Windows.Forms.MenuStrip # Creates the Menu Strip the menu items sit in
     $menuFile = New-Object System.Windows.Forms.ToolStripMenuItem # Root File Menu Text
     $menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem # Root Help Menu text
     $menuFileOpen = New-Object System.Windows.Forms.ToolStripMenuItem # Open File Menu Option
     $menuFileQuit = New-Object System.Windows.Forms.ToolStripMenuItem # Quit Menu Option
     $menuHelpInfo = New-Object System.Windows.Forms.ToolStripMenuItem # Main Script Help Menu Option
     $menuHelpView = New-Object System.Windows.Forms.ToolStripMenuItem # View Script Source Menu option
     $separatorF = New-Object System.Windows.Forms.ToolStripSeparator # Seperator Line in the File Menu
     $separatorH = New-Object System.Windows.Forms.ToolStripSeparator # Seperator Line in the Help Menu
     # Add Other Objects to Main Form
     $tabControl = New-Object System.Windows.Forms.TabControl # Initialize Tab Area
     $firstTab = New-Object System.Windows.Forms.TabPage # Initialize First Tab in the Tab Area, More can be added and then define in a new section below
     $buttonClose = New-Object System.Windows.Forms.Button # Initialize Main Form Close Button
     $buttonInstall = New-Object System.Windows.Forms.Button # Initialize Main Form Send Button
     $processingLabel = New-Object System.Windows.Forms.Label # Initialize Processing Label (Task Processing bar)
     $textBoxManEntListComp = New-Object System.Windows.Forms.RichTextBox # Initialize Computer Manual Entry Text Box
     $buttonManEntClearComp = New-Object System.Windows.Forms.Button # Initialize Computer Manual Entry Clear Button
     $buttonManEntImportComp = New-Object System.Windows.Forms.Button # Initialize Computer Manual Entry Import Button
     $labelManEntListComp = New-Object System.Windows.Forms.Label # Initialize Label for Computer Manual Entry box
     $descrLablManEntListComp = New-Object System.Windows.Forms.Label # Initialize Description for Computer Manual Entry box
     $textBoxManEntListPatch = New-Object System.Windows.Forms.RichTextBox # Initialize Patch Name Entry Text Box
     $buttonManEntClearPatch = New-Object System.Windows.Forms.Button # Initialize Patch Name Entry Clear Button
     $buttonManEntFileNamePatch = New-Object System.Windows.Forms.Button # Initialize Patch Name Entry Import Button
     $labelManEntListPatch = New-Object System.Windows.Forms.Label # Initialize Label for Patch Name entry text box
     $descrLablManEntListPatch = New-Object System.Windows.Forms.Label # Initialize Description for Patch Name entry text box
     $textBoxManEntListPsexec = New-Object System.Windows.Forms.RichTextBox # Initialize Psexec Filepath Entry Text Box
     $buttonManEntClearPsexec = New-Object System.Windows.Forms.Button # Initialize Psexec Filepath Entry Clear Button
     $buttonManEntFilePathPsexec = New-Object System.Windows.Forms.Button # Initialize Psexec Filepath Entry Import Button
     $labelManEntListPsexec = New-Object System.Windows.Forms.Label # Initialize Label for Psexec Filepath entry text box
     $descrLablManEntListPsexec = New-Object System.Windows.Forms.Label # Initialize Description for Psexec Filepath entry text box
     $textBoxManEntListPPath = New-Object System.Windows.Forms.RichTextBox # Initialize Update/Patch Folder Location Entry Text Box
     $buttonManEntClearPPath = New-Object System.Windows.Forms.Button # Initialize PUpdate/Patch Folder Location Entry Clear Button
     $buttonManEntFolderPPath = New-Object System.Windows.Forms.Button # Initialize Update/Patch Folder Location Entry Import Button
     $labelManEntListPPath = New-Object System.Windows.Forms.Label # Initialize Label for Update/Patch Folder Location entry text box
     $descrLablManEntListPPath = New-Object System.Windows.Forms.Label # Initialize Description for Update/Patch Folder Location entry text box
     $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = "C:\"; Multiselect = $false; Title = "Choose File..."}  # Initialize Open File Dialog
     $openFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog # Initialize Open Folder Dialog
     $installFormFont = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Underline)

     #-->  Build Main Background and base properties of the Form
     # Sets Size of form
     $System_Drawing_Size.Height = 520
     $System_Drawing_Size.Width = 453
     $installForm.ClientSize = $System_Drawing_Size
     $installForm.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets minimumm Size of form (Cannot be resized lower than this)
     $System_Drawing_Size.Height = 560
     $System_Drawing_Size.Width = 469
     $installForm.MinimumSize = $System_Drawing_Size
     # Form Setttings
     $installForm.Icon = [System.IconExtractor]::Extract("imageres.dll", 63, $False)
     $installForm.StartPosition = "CenterScreen"
     $installForm.Name = "installForm"
     $installForm.Text = "Install-Update Function"
     $installForm.AcceptButton = $buttonInstall
     $installForm.CancelButton = $buttonClose

     #-------------------------------------------

     #-->  Build Processing Bar, placed next to tab name label
     $processingLabel.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 120
     $System_Drawing_Point.Y = 22
     $processingLabel.Location = $System_Drawing_Point
     $processingLabel.Name = "processingLabel"
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
     # Add text to Area
     $processingLabel.Hide()
     $installForm.Controls.Add($processingLabel)

     #-------------------------------------------

     #--> Build Top Menu
     #    --> Build File Menu
     # Add a MenuStrip object for Quitting the form
     $menuFileQuit.Text = "&Quit"
     $menuFileQuit.ShortcutKeys = "Control, Q"
     $menuFileQuit.add_Click( {
               $installForm.Close()
          }
     )
     # Add a MenuStrip object for Importing a text file to import computers names from
     $menuFileOpen.Text = "&Import Device List"
     $menuFileOpen.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::O
     # Actions for the Import Menu
     $menuFileOpen.add_Click( {
               $openFileDialog.Filter = "List Files (*.txt)|*.txt"
               $openFileDialog.ShowDialog()
               Parse-TextFile -Path $openFileDialog.FileName -TextBox $textBoxManEntListComp
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
     $menuHelp.DropDownItems.AddRange(@($menuHelpInfo, $separatorH, $menuHelpView))

     # Add both Menus to to the form
     $menu.Items.AddRange(@($menuFile, $menuHelp))
     $installForm.Controls.Add($menu)

     #-------------------------------------------

     #--> Build Main Tab Area
     # Build Tab Anchor base
     $tabControl.Anchor = 15
     $tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Tab
     $System_Drawing_Point.X = 7
     $System_Drawing_Point.Y = 27
     $tabControl.Location = $System_Drawing_Point
     $tabControl.Name = "tabControl"
     $tabControl.SelectedIndex = 0
     # Sets size of Tab Anchor
     $System_Drawing_Size.Height = 485
     $System_Drawing_Size.Width = 440
     $tabControl.Size = $System_Drawing_Size
     $tabControl.TabIndex = 4
     # Add Anchor to Form
     $installForm.Controls.Add($tabControl)

     # Build Tab area to hold objects
     $firstTab.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Tab Area
     $System_Drawing_Point.X = 4
     $System_Drawing_Point.Y = 22
     $firstTab.Location = $System_Drawing_Point
     $firstTab.Name = "firstTab"
     $System_Windows_Forms_Padding.All = 3
     $System_Windows_Forms_Padding.Bottom = 3
     $System_Windows_Forms_Padding.Left = 3
     $System_Windows_Forms_Padding.Right = 3
     $System_Windows_Forms_Padding.Top = 3
     $firstTab.Padding = $System_Windows_Forms_Padding
     # Sets size of Area
     $System_Drawing_Size.Height = 422
     $System_Drawing_Size.Width = 400
     $firstTab.Size = $System_Drawing_Size
     $firstTab.TabIndex = 0
     $firstTab.Text = "Install Update"
     $firstTab.UseVisualStyleBackColor = $True
     # Add tab to Form
     $tabControl.Controls.Add($firstTab)

     # Builds Close Button and Adds to the Main tab
     $buttonClose.Anchor = 2
     $buttonClose.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button on the Form
     $System_Drawing_Point.X = 198
     $System_Drawing_Point.Y = 395
     $buttonClose.Location = $System_Drawing_Point
     $buttonClose.Name = "buttonClose"
     # Sets Size of Button
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonClose.Size = $System_Drawing_Size
     $buttonClose.TabIndex = 8
     # Sets Text of Button
     $buttonClose.Text = "Close"
     $buttonClose.UseVisualStyleBackColor = $True
     # Sets action on click (Can be any function, cmdlet, .NET Method, logic control, etc.)
     $buttonClose.add_Click( {
               $installForm.Close()
          }
     )
     # Add Close button to Form
     $firstTab.Controls.Add($buttonClose)

     # Builds Install Button and adds to Main Tab
     $buttonInstall.Anchor = 2
     $buttonInstall.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button on the Form
     $System_Drawing_Point.X = 117
     $System_Drawing_Point.Y = 395
     $buttonInstall.Location = $System_Drawing_Point
     $buttonInstall.Name = "buttonInstall"
     # Sets Size of Button
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonInstall.Size = $System_Drawing_Size
     $buttonInstall.TabIndex = 7
     # Sets Text of Button
     $buttonInstall.Text = "Install"
     $buttonInstall.UseVisualStyleBackColor = $True
     # Actions for the "install" button
     $buttonInstall.add_Click( {
               # Pulls all info from Entry boxes, turns them in the proper output, and then runs install-update
               if ($textBoxManEntListComp.text -eq "") {
                    $popupMessage = "You need to input the computer names into the Device Name Entry List Box. It cannot be left blank."
                    $popupTitle = "Device Name Entry Box is Blank!"
                    $popupIcon = "Information"
                    Create-InfoMessageBox -message $popupMessage -title $popupTitle -icon $popupIcon
                    return
               }
               Else {
                    Parse-Input -objects $textBoxManEntListComp.Text
               }
               if ($textBoxManEntListPatch.text -eq "") {
                    $popupMessage = "You need to input the patch name into the Patch Name Entry List Box. It cannot be left blank."
                    $popupTitle = "Patch Entry Box is Blank!"
                    $popupIcon = "Information"
                    Create-InfoMessageBox -message $popupMessage -title $popupTitle -icon $popupIcon
                    return
               }
               Else {
                    $updatePatchName = $textBoxManEntListPatch.Text
               }
               if ($updatePatchPath -ne "") {
                    $updatePatchPath = $textBoxManEntListPPath.Text
               }
               if ($psexecFilePath -ne "") {
                    $psexecFilePath = $textBoxManEntListPsexec.Text
               }
               Install-Update -computers $computers -updatePatchName $updatePatchName -updatePatchPath $updatePatchPath -psexecFilePath $psexecFilePath
               Create-ResultsForm
               # Cleanup after Intsall Update completes
               $Script:computers = @()
               $Script:updatePatchName = ""
               $Script:updatePatchPath = "C:\UpdatePush\Patches"
               $Script:psexecFilePath = "C:\UpdatePush\PSExec\PSExec.exe"
               $Script:exitCodeReport = @()
               $textBoxManEntListComp.clear()
               $textBoxManEntListPatch.clear()
               $textBoxManEntListPPath.text = $updatePatchPath
               $textBoxManEntListPsexec.text = $psexecFilePath
               $processingLabel.hide()
          }
     )
     # Add Send button to Form
     $firstTab.Controls.Add($buttonInstall)

     #-------------------------------------------

     #--> Computers Manual Device Name/IP Entry Area

     # Adds Manual Entry Label to the Computers Manual Device Name/IP Entry Area
     $labelManEntListComp.Anchor = 13
     $labelManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0

     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 10
     $labelManEntListComp.Location = $System_Drawing_Point
     $labelManEntListComp.Name = "labelManEntListComp"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 14
     $System_Drawing_Size.Width = 366
     $labelManEntListComp.Size = $System_Drawing_Size
     $labelManEntListComp.TabIndex = 0
     $labelManEntListComp.Font = $installFormFont
     # Sets Text to show
     $labelManEntListComp.Text = "Device Name/IP Entry"
     # Add Text to Area
     $firstTab.Controls.Add($labelManEntListComp)

     # Adds Text instructions to the Computers Manual Device Name/IP Entry Area
     $descrLablManEntListComp.Anchor = 13
     $descrLablManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 30
     $descrLablManEntListComp.Location = $System_Drawing_Point
     $descrLablManEntListComp.Name = "descrLablManEntListComp"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 28
     $System_Drawing_Size.Width = 366
     $descrLablManEntListComp.Size = $System_Drawing_Size
     $descrLablManEntListComp.TabIndex = 0
     # Sets Text to show
     $descrLablManEntListComp.Text = "Type the workstation name(s)/IP(s) separated by spaces, commas, semi-colons, or import a list of names from a text file:"
     # Add Text to Area
     $firstTab.Controls.Add($descrLablManEntListComp)

     # Add Text Box to the Manual Entry Computers Manual Device Name/IP Entry Area
     $textBoxManEntListComp.Anchor = 13
     $textBoxManEntListComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text Box
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 60
     $textBoxManEntListComp.Location = $System_Drawing_Point
     $textBoxManEntListComp.Name = "textBoxManEntListComp"
     # Sets Size of Text Box
     $System_Drawing_Size.Height = 20
     $System_Drawing_Size.Width = 385
     $textBoxManEntListComp.Size = $System_Drawing_Size
     $textBoxManEntListComp.TabIndex = 1
     # Adds Text Box to Area
     $firstTab.Controls.Add($textBoxManEntListComp)

     # Add Import File Button to Computers Manual Device Name/IP Entry Area
     $buttonManEntImportComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 7
     $System_Drawing_Point.Y = 85
     $buttonManEntImportComp.Location = $System_Drawing_Point
     $buttonManEntImportComp.Name = "buttonManEntImportComp"
     # Sets Size of Button
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntImportComp.Size = $System_Drawing_Size
     $buttonManEntImportComp.TabIndex = 2
     # Sets Title of Button
     $buttonManEntImportComp.Text = "Import List"
     $buttonManEntImportComp.UseVisualStyleBackColor = $True
     # Actions for the "Import File" Button
     $buttonManEntImportComp.add_Click( {
               $openFileDialog.Filter = "List Files (*.txt)|*.txt"
               $openFileDialog.ShowDialog()
               Parse-TextFile -Path $openFileDialog.FileName -TextBox $textBoxManEntListComp
          }
     )
     # Adds Import File Button to Area
     $firstTab.Controls.Add($buttonManEntImportComp)

     # Build Clear Text Box Button to the Computers Manual Device Name/IP Entry Area
     $buttonManEntClearComp.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 89
     $System_Drawing_Point.Y = 85
     $buttonManEntClearComp.Location = $System_Drawing_Point
     $buttonManEntClearComp.Name = "buttonManEntClearComp"
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
     $firstTab.Controls.Add($buttonManEntClearComp)

     #-------------------------------------------

     #--> Update/Patch Filepath Entry Box Area

     # Adds Label to the Update/Patch Filepath Entry Box Area
     $labelManEntListPatch.Anchor = 13
     $labelManEntListPatch.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 115
     $labelManEntListPatch.Location = $System_Drawing_Point
     $labelManEntListPatch.Name = "labelManEntListPatch"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 14
     $System_Drawing_Size.Width = 366
     $labelManEntListPatch.Size = $System_Drawing_Size
     $labelManEntListPatch.TabIndex = 0
     $labelManEntListPatch.Font = $installFormFont
     # Sets Label Text
     $labelManEntListPatch.Text = "Patch Filename Entry"
     # Add Text to Area
     $firstTab.Controls.Add($labelManEntListPatch)

     # Adds Text instructions to the Update/Patch Filepath Entry Box Area
     $descrLablManEntListPatch.Anchor = 13
     $descrLablManEntListPatch.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 135
     $descrLablManEntListPatch.Location = $System_Drawing_Point
     $descrLablManEntListPatch.Name = "descrLablManEntListPatch"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 14
     $System_Drawing_Size.Width = 366
     $descrLablManEntListPatch.Size = $System_Drawing_Size
     $descrLablManEntListPatch.TabIndex = 0
     # Sets Text to show
     $descrLablManEntListPatch.Text = "Choose the File Location of the Update/Patch to Install:"
     # Add Text to Area
     $firstTab.Controls.Add($descrLablManEntListPatch)

     # Add Text Box to the Update/Patch Filepath Entry Box Area
     $textBoxManEntListPatch.Anchor = 13
     $textBoxManEntListPatch.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text Box
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 155
     $textBoxManEntListPatch.Location = $System_Drawing_Point
     $textBoxManEntListPatch.Name = "textBoxManEntListPatch"
     # Sets Size of Text Box
     $System_Drawing_Size.Height = 20
     $System_Drawing_Size.Width = 385
     $textBoxManEntListPatch.Size = $System_Drawing_Size
     $textBoxManEntListPatch.TabIndex = 1
     # Adds Text Box to Area
     $firstTab.Controls.Add($textBoxManEntListPatch)

     # Add Get Filename Button to Update/Patch Filepath Entry Box Area
     $buttonManEntFileNamePatch.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 7
     $System_Drawing_Point.Y = 180
     $buttonManEntFileNamePatch.Location = $System_Drawing_Point
     $buttonManEntFileNamePatch.Name = "buttonManEntFileNamePatch"
     # Sets Size of Button
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntFileNamePatch.Size = $System_Drawing_Size
     $buttonManEntFileNamePatch.TabIndex = 2
     # Sets Title of Button
     $buttonManEntFileNamePatch.Text = "Choose..."
     $buttonManEntFileNamePatch.UseVisualStyleBackColor = $True
     # Actions for the "Get Filename" Button
     $buttonManEntFileNamePatch.add_Click( {
               $openFileDialog.Filter = "Update Executables (*.exe;*.msu;*.cab)|*.exe;*.msu;*.cab"
               $openFileDialog.ShowDialog()
               $textBoxManEntListPatch.Text = $openFileDialog.SafeFileName
          }
     )
     # Adds Get Filename Button to Update/Patch Filepatch Entry Box Area
     $firstTab.Controls.Add($buttonManEntFileNamePatch)

     # Build Clear Text Box Button to the Update/Patch Filepath Entry Box Area
     $buttonManEntClearPatch.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 89
     $System_Drawing_Point.Y = 180
     $buttonManEntClearPatch.Location = $System_Drawing_Point
     $buttonManEntClearPatch.Name = "buttonManEntClearPatch"
     # Sets Size of Buttons
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntClearPatch.Size = $System_Drawing_Size
     $buttonManEntClearPatch.TabIndex = 3
     # Sets Button Text
     $buttonManEntClearPatch.Text = "Clear Text"
     $buttonManEntClearPatch.UseVisualStyleBackColor = $True
     # Actions for the Clear Button (Removes all text from the text box)
     $buttonManEntClearPatch.add_Click( {
               $textBoxManEntListPatch.Clear()
          }
     )
     # Add Clear Button to Area
     $firstTab.Controls.Add($buttonManEntClearPatch)

     #-------------------------------------------

     #--> Update/Patch Folder Location Entry Box Area

     # Adds Label to the Update/Patch Folder Location Entry Box Area
     $labelManEntListPPath.Anchor = 13
     $labelManEntListPPath.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 210
     $labelManEntListPPath.Location = $System_Drawing_Point
     $labelManEntListPPath.Name = "labelManEntListPPath"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 14
     $System_Drawing_Size.Width = 366
     $labelManEntListPPath.Size = $System_Drawing_Size
     $labelManEntListPPath.TabIndex = 0
     $labelManEntListPPath.Font = $installFormFont
     # Sets Text to show
     $labelManEntListPPath.Text = "Update/Patch Folder Location Entry (OPTIONAL)"
     # Add Text to Area
     $firstTab.Controls.Add($labelManEntListPPath)

     # Adds Text instructions to the Update/Patch Folder Location Entry Box Area
     $descrLablManEntListPPath.Anchor = 13
     $descrLablManEntListPPath.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 230
     $descrLablManEntListPPath.Location = $System_Drawing_Point
     $descrLablManEntListPPath.Name = "descrLablManEntListPPath"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 28
     $System_Drawing_Size.Width = 366
     $descrLablManEntListPPath.Size = $System_Drawing_Size
     $descrLablManEntListPPath.TabIndex = 0
     # Sets Text to show
     $descrLablManEntListPPath.Text = "Choose the folder location where the update/patch files are located.`n(Defaults to -> C:\UpdatePush\Patches):"
     # Add Text to Area
     $firstTab.Controls.Add($descrLablManEntListPPath)

     # Add Text Box to the Update/Patch Folder Location Entry Box Area
     $textBoxManEntListPPath.Anchor = 13
     $textBoxManEntListPPath.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text Box
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 260
     $textBoxManEntListPPath.Location = $System_Drawing_Point
     $textBoxManEntListPPath.Name = "textBoxManEntListPPath"
     # Sets Size of Text Box
     $System_Drawing_Size.Height = 20
     $System_Drawing_Size.Width = 385
     $textBoxManEntListPPath.Size = $System_Drawing_Size
     $textBoxManEntListPPath.TabIndex = 1
     # Set Default Entry to Default Global Variable
     $textBoxManEntListPPath.text = $updatePatchPath
     # Adds Text Box to Area
     $firstTab.Controls.Add($textBoxManEntListPPath)

     # Add Get Folder Button to Update/Patch Folder Location Entry Box Area
     $buttonManEntFolderPPath.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 7
     $System_Drawing_Point.Y = 285
     $buttonManEntFolderPPath.Location = $System_Drawing_Point
     $buttonManEntFolderPPath.Name = "buttonManEntFilePathPPath"
     # Sets Size of Button
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntFolderPPath.Size = $System_Drawing_Size
     $buttonManEntFolderPPath.TabIndex = 2
     # Sets Title of Button
     $buttonManEntFolderPPath.Text = "Choose..."
     $buttonManEntFolderPPath.UseVisualStyleBackColor = $True
     # Actions for the "Get Folder" Button
     $buttonManEntFolderPPath.add_Click( {
               $openFolderDialog.RootFolder = "c:\"
               $openFolderDialog.ShowNewFolderButton = $false
               $openFolderDialog.Description = "Choose Folder Location of Patch Folder..."
               $openFolderDialog.ShowDialog()
               $textBoxManEntListPPath.Text = $openFolderDialog.SelectedPath
          }
     )
     # Adds Import File Button to Area
     $firstTab.Controls.Add($buttonManEntFolderPPath)

     # Build Clear Text Box Button to the Update/Patch Folder Location Entry Box Area
     $buttonManEntClearPPath.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 89
     $System_Drawing_Point.Y = 285
     $buttonManEntClearPPath.Location = $System_Drawing_Point
     $buttonManEntClearPPath.Name = "buttonManEntClearPPath"
     # Sets Size of Buttons
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntClearPPath.Size = $System_Drawing_Size
     $buttonManEntClearPPath.TabIndex = 3
     # Sets Button Text
     $buttonManEntClearPPath.Text = "Clear Text"
     $buttonManEntClearPPath.UseVisualStyleBackColor = $True
     # Actions for the Clear Button (Removes all text from the text box)
     $buttonManEntClearPPath.add_Click( {
               $textBoxManEntListPPath.Clear()
          }
     )
     # Add Clear Button to Area
     $firstTab.Controls.Add($buttonManEntClearPPath)

     #-------------------------------------------

     #--> PSExec Filepath Entry Box Area

     # Adds Label to the PSExec Filepath Entry Box Area
     $labelManEntListPsexec.Anchor = 13
     $labelManEntListPsexec.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 315
     $labelManEntListPsexec.Location = $System_Drawing_Point
     $labelManEntListPsexec.Name = "labelManEntListPsexec"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 14
     $System_Drawing_Size.Width = 366
     $labelManEntListPsexec.Size = $System_Drawing_Size
     $labelManEntListPsexec.TabIndex = 0
     $labelManEntListPsexec.Font = $installFormFont
     # Sets Text to show
     $labelManEntListPsexec.Text = "Psexec Filename Entry (OPTIONAL)"
     # Add Text to Area
     $firstTab.Controls.Add($labelManEntListPsexec)

     # Adds Text instructions to the PSExec Filepath Entry Box Area
     $descrLablManEntListPsexec.Anchor = 13
     $descrLablManEntListPsexec.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 335
     $descrLablManEntListPsexec.Location = $System_Drawing_Point
     $descrLablManEntListPsexec.Name = "descrLablManEntListPsexec"
     # Sets Size of Text Area where text will be
     $System_Drawing_Size.Height = 28
     $System_Drawing_Size.Width = 366
     $descrLablManEntListPsexec.Size = $System_Drawing_Size
     $descrLablManEntListPsexec.TabIndex = 0
     # Sets Text to show
     $descrLablManEntListPsexec.Text = "Choose the filepath of the PSExec file `n(Defaults to -> C:\UpdatePush\PSExec\PSexec.exe):"
     # Add Text to Area
     $firstTab.Controls.Add($descrLablManEntListPsexec)

     # Add Text Box to the PSExec Filepath Entry Box Area
     $textBoxManEntListPsexec.Anchor = 13
     $textBoxManEntListPsexec.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Text Box
     $System_Drawing_Point.X = 6
     $System_Drawing_Point.Y = 365
     $textBoxManEntListPsexec.Location = $System_Drawing_Point
     $textBoxManEntListPsexec.Name = "textBoxManEntListPsexec"
     # Sets Size of Text Box
     $System_Drawing_Size.Height = 20
     $System_Drawing_Size.Width = 385
     $textBoxManEntListPsexec.Size = $System_Drawing_Size
     $textBoxManEntListPsexec.TabIndex = 1
     # Set Default Entry to Default Global Variable
     $textBoxManEntListPsexec.text = $psexecFilePath
     # Adds Text Box to Area
     $firstTab.Controls.Add($textBoxManEntListPsexec)

     # Add Import File Button to PSExec Filepath Entry Box Area
     $buttonManEntFilePathPsexec.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 7
     $System_Drawing_Point.Y = 390
     $buttonManEntFilePathPsexec.Location = $System_Drawing_Point
     $buttonManEntFilePathPsexec.Name = "buttonManEntFilePathPsexec"
     # Sets Size of Button
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntFilePathPsexec.Size = $System_Drawing_Size
     $buttonManEntFilePathPsexec.TabIndex = 2
     # Sets Title of Button
     $buttonManEntFilePathPsexec.Text = "Choose..."
     $buttonManEntFilePathPsexec.UseVisualStyleBackColor = $True
     # Actions for the "Import File" Button
     $buttonManEntFilePathPsexec.add_Click( {
               $openFileDialog.Filter = "PSExec Executables (*.exe)|*.exe"
               $openFileDialog.ShowDialog()
               $textBoxManEntListPsexec.Text = $openFileDialog.FileName
          }
     )
     # Adds Import File Button to Area
     $firstTab.Controls.Add($buttonManEntFilePathPsexec)

     # Build Clear Text Box Button to the PSExec Filepath Entry Box Area
     $buttonManEntClearPsexec.DataBindings.DefaultDataSourceUpdateMode = 0
     # Sets Location of Button
     $System_Drawing_Point.X = 89
     $System_Drawing_Point.Y = 390
     $buttonManEntClearPsexec.Location = $System_Drawing_Point
     $buttonManEntClearPsexec.Name = "buttonManEntClearPsexec"
     # Sets Size of Buttons
     $System_Drawing_Size.Height = 23
     $System_Drawing_Size.Width = 75
     $buttonManEntClearPsexec.Size = $System_Drawing_Size
     $buttonManEntClearPsexec.TabIndex = 3
     # Sets Button Text
     $buttonManEntClearPsexec.Text = "Clear Text"
     $buttonManEntClearPsexec.UseVisualStyleBackColor = $True
     # Actions for the Clear Button (Removes all text from the text box)
     $buttonManEntClearPsexec.add_Click( {
               $textBoxManEntListPsexec.Clear()
          }
     )
     # Add Clear Button to Area
     $firstTab.Controls.Add($buttonManEntClearPsexec)

     #-------------------------------------------

     #--> Additonal Objects

     # Add Additiontal objects you wish to add to the form, usually based on script parameters

     #-------------------------------------------

     #----------------------------------
     #  Initial Main Form runtime code
     #----------------------------------

     # Save the initial state of the form
     $Script:InitialFormWindowState = $installForm.WindowState

     # Init the OnLoad event to correct the initial state of the form
     # Corrects the initial state of the form to prevent the .Net maximized form issue
     $OnLoadForm_StateCorrection = $initialFormWindowState
     $installForm.add_Load($OnLoadForm_StateCorrection)

     # Start Main Form
     $installForm.Add_Shown( {$installForm.Activate()})
     $installForm.ShowDialog($this) | Out-Null

}
Function Create-ViewSourceForm {
     # Builds View Source Form from the help menu "View Source"
     # Add objects for Source Viewer
     $formSourceCode = New-Object System.Windows.Forms.Form
     $richTextBoxSource = New-Object System.Windows.Forms.RichTextBox
     # Form for viewing source code
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
     # Set Location For the Rich Text Box
     $System_Drawing_Point.X = 13
     $System_Drawing_Point.Y = 13
     $richTextBoxSource.Location = $System_Drawing_Point
     $richTextBoxSource.Name = "richTextBoxSource"
     $richTextBoxSource.Font = New-Object System.Drawing.Font("Consolas New", 10)
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
Function Create-ResultsForm {
     # Builds Results Form from the help menu "Results"
     # Add objects for Results Form
     $formResultsCode = New-Object System.Windows.Forms.Form
     $richTextBoxResults = New-Object System.Windows.Forms.RichTextBox
     # Form for viewing Results
     # Set Size of Results Form
     $System_Drawing_Size.Height = 600
     $System_Drawing_Size.Width = 725
     $formResultsCode.ClientSize = $System_Drawing_Size
     #$formResultsCode.DataBindings.DefaultDataSourceUpdateMode = 0
     $formResultsCode.StartPosition = "CenterScreen"
     $formResultsCode.Name = "formSourceCode"
     # Set View Results Title
     $formResultsCode.Text = "Install Update Results"
     $formResultsCode.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSCommandPath)
     # Add Rich Text Box for the Help Text to display in
     $richTextBoxResults.Anchor = 15
     $richTextBoxResults.DataBindings.DefaultDataSourceUpdateMode = 0
     # Set Location For the Rich Text Box
     $System_Drawing_Point.X = 13
     $System_Drawing_Point.Y = 13
     $richTextBoxResults.Location = $System_Drawing_Point
     $richTextBoxResults.Name = "richTextBoxResults"
     $richTextBoxResults.Font = New-Object System.Drawing.Font("lucida console", 10)
     # Set Rich Text Box Size
     $System_Drawing_Size.Height = 570
     $System_Drawing_Size.Width = 700
     $richTextBoxResults.Size = $System_Drawing_Size
     $richTextBoxResults.DetectUrls = $False
     $richTextBoxResults.ReadOnly = $True
     # Get source from script file and add newline to each array item for formatting
     $richTextBoxResults.Text = "Results of the Patch Install:`n" + "`n" + ($exitCodeReport | Format-Table -AutoSize | Out-String)
     # Add Rich text Box to Results Form
     $formResultsCode.Controls.Add($richTextBoxResults)
     # Show Results Form
     $formResultsCode.Show() | Out-Null
}
Function Create-HelpForm {
     # Build Help Form
     $formHelp = New-Object System.Windows.Forms.Form
     $richTextBoxHelp = New-Object System.Windows.Forms.RichTextBox
     # Add Objects to Help form
     $formHelp.AutoScroll = $True
     # SetHelp Size of help Form
     $System_Drawing_Size.Height = 600
     $System_Drawing_Size.Width = 500
     $formHelp.ClientSize = $System_Drawing_Size
     #$formHelp.DataBindings.DefaultDataSourceUpdateMode = 0
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
     $System_Drawing_Point.X = 13
     $System_Drawing_Point.Y = 13
     $richTextBoxHelp.Location = $System_Drawing_Point
     $richTextBoxHelp.Name = "richTextBoxHelp"
     $richTextBoxHelp.Font = New-Object System.Drawing.Font("Courier New", 10)
     $richTextBoxHelp.ReadOnly = $True
     $richTextBoxHelp.SelectionProtected = $True
     $richTextBoxHelp.Cursor = [System.Windows.Forms.Cursors]::Default
     # Set Size for Rich Text
     $System_Drawing_Size.Height = 570
     $System_Drawing_Size.Width = 469
     $richTextBoxHelp.Size = $System_Drawing_Size
     $richTextBoxHelp.TabIndex = 0
     $richTextBoxHelp.TabStop = $False
     # Set Text For Help form into the Rich Text Box
     # Text must be between the '', and it must starts at the beggining of the line as show.
     # Type in the Help text you wish to show.
     $richTextBoxHelp.Text = 'INTRODUCTION

This script was designed to Install Updates to remote computers using standard windows update installations tools and the PSExec remote service program. It will accept a list of computers or a single machine, and install the given update to those machines.

TO A SINGLE COMPUTER

Enter the workstation name or IP into the "Device Name/IP" text box, Choose the patch using the "Patch Filename Entry" Text Box open file dialog, and Press Install.

TO MULTIPLE COMPUTERS

Enter the workstations names or IPs into the "Device Name/IP" text box or import a list of machine names or IPs, Choose the patch using the "Patch Filename Entry" Text Box open file dialog, and Press Install.

VIEWING SCRIPT SOURCE CODE

You may pull up the source code of the script by choosing Help -> View Script from the menu or by pressing Ctrl + E.'

     # AddHelp Rich Text box to Help Form
     $formHelp.Controls.Add($richTextBoxHelp)
     # Show Help Form
     $formHelp.Show() | Out-Null
}

#---------------------------------
#endregion
#---------------------------------

#-------------------------------
#region  Install-Update Main Actions
#-------------------------------

# Uncomment next line if Script must be run with Admin Rights.
Check-IfAdmin

# Show the Form
Create-MainForm

#---------------------------------
#endregion
#---------------------------------