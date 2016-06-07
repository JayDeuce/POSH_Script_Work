<#
.DESCRIPTION
This script uses Windows Forms to present a GUI for sending messages to remote computers on a network using msg.exe.  You can send to a single computer, multiple computers entered manually, or multiple computers from a text file.
#>

#----------------------------
#    Form Initialization
#----------------------------

#region Import the necessary Assemblies

[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

#endregion Import the necessary Assemblies

#region Generate Main Form Objects

# Main Form Object creation
$form1 = New-Object System.Windows.Forms.Form

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

# Other Objects to Form
<# 
tabControl1 = The full tab object
MsgTab = The Send Message Tab
buttonClose = The Close button at the bottom of the form
buttonSend = The Send button at the bottom of the form
grpListComp = The Send to Computers section of the tab
buttonClearComp = The Clear button in the Send to Computers section
buttonManEntListComp = The Open button in the Send to Computers section
textBoxListComp = the Text box in the Send to Computers section
labelListComp = The Text in the Send to Computers section
labelMsg = The text in the Type your Message section
richtextBoxMsg = The rich text box in the type your Message Section
openFileDialog1 = The open file dialog window and its options
InitialForWindowState = the initial settings of the form winodw when it opens
#>
$tabControl1 = New-Object System.Windows.Forms.TabControl
$MsgTab = New-Object System.Windows.Forms.TabPage
$buttonClose = New-Object System.Windows.Forms.Button
$buttonSend = New-Object System.Windows.Forms.Button
$grpListComp = New-Object System.Windows.Forms.GroupBox
$autoEntListComp = New-Object System.Windows.Forms.Label
$labelAutoEntListComp = New-Object System.Windows.Forms.Label
$checkBoxDunhamIP = New-Object System.Windows.Forms.CheckBox
$checkBoxFillmoreIP = New-Object System.Windows.Forms.CheckBox
$checkBoxFIGIP = New-Object System.Windows.Forms.CheckBox
$checkBoxLEADIP = New-Object System.Windows.Forms.CheckBox
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

[array]$dunhamIPs = "160.151.198.1","160.151.198.2","160.151.198.3","160.151.198.4","160.151.198.5","160.151.198.6","160.151.198.7","160.151.198.8","160.151.198.9","160.151.198.10","160.151.198.11","160.151.198.12","160.151.198.13","160.151.198.14","160.151.198.15","160.151.198.16","160.151.198.17","160.151.198.18","160.151.198.19","160.151.198.20","160.151.198.21","160.151.198.22","160.151.198.23","160.151.198.24","160.151.198.25","160.151.198.26","160.151.198.27","160.151.198.28","160.151.198.29","160.151.198.30","160.151.198.31","160.151.198.32","160.151.198.33","160.151.198.34","160.151.198.35","160.151.198.36","160.151.198.37","160.151.198.38","160.151.198.39","160.151.198.40","160.151.198.41","160.151.198.42","160.151.198.43","160.151.198.44","160.151.198.45","160.151.198.46","160.151.198.47","160.151.198.48","160.151.198.49","160.151.198.50","160.151.198.51","160.151.198.52","160.151.198.53","160.151.198.54","160.151.198.55","160.151.198.56","160.151.198.57","160.151.198.58","160.151.198.59","160.151.198.60","160.151.198.61","160.151.198.62","160.151.198.63","160.151.198.64","160.151.198.65","160.151.198.66","160.151.198.67","160.151.198.68","160.151.198.69","160.151.198.70","160.151.198.71","160.151.198.72","160.151.198.73","160.151.198.74","160.151.198.75","160.151.198.76","160.151.198.77","160.151.198.78","160.151.198.79","160.151.198.80","160.151.198.81","160.151.198.82","160.151.198.83","160.151.198.84","160.151.198.85","160.151.198.86","160.151.198.87","160.151.198.88","160.151.198.89","160.151.198.90","160.151.198.91","160.151.198.92","160.151.198.93","160.151.198.94","160.151.198.95","160.151.198.96","160.151.198.97","160.151.198.98","160.151.198.99","160.151.198.100","160.151.198.101","160.151.198.102","160.151.198.103","160.151.198.104","160.151.198.105","160.151.198.106","160.151.198.107","160.151.198.108","160.151.198.109","160.151.198.110","160.151.198.111","160.151.198.112","160.151.198.113","160.151.198.114","160.151.198.115","160.151.198.116","160.151.198.117","160.151.198.118","160.151.198.119","160.151.198.120","160.151.198.121","160.151.198.122","160.151.198.123","160.151.198.124","160.151.198.125","160.151.198.126","160.151.198.127","160.151.198.128","160.151.198.129","160.151.198.130","160.151.198.131","160.151.198.132","160.151.198.133","160.151.198.134","160.151.198.135","160.151.198.136","160.151.198.137","160.151.198.138","160.151.198.139","160.151.198.140","160.151.198.141","160.151.198.142","160.151.198.143","160.151.198.144","160.151.198.145","160.151.198.146","160.151.198.147","160.151.198.148","160.151.198.149","160.151.198.150","160.151.198.151","160.151.198.152","160.151.198.153","160.151.198.154","160.151.198.155","160.151.198.156","160.151.198.157","160.151.198.158","160.151.198.159","160.151.198.160","160.151.198.161","160.151.198.162","160.151.198.163","160.151.198.164","160.151.198.165","160.151.198.166","160.151.198.167","160.151.198.168","160.151.198.169","160.151.198.170","160.151.198.171","160.151.198.172","160.151.198.173","160.151.198.174","160.151.198.175","160.151.198.176","160.151.198.177","160.151.198.178","160.151.198.179","160.151.198.180","160.151.198.181","160.151.198.182","160.151.198.183","160.151.198.184","160.151.198.185","160.151.198.186","160.151.198.187","160.151.198.188","160.151.198.189","160.151.198.190","160.151.198.191","160.151.198.192","160.151.198.193","160.151.198.194","160.151.198.195","160.151.198.196","160.151.198.197","160.151.198.198","160.151.198.199","160.151.198.200","160.151.198.201","160.151.198.202","160.151.198.203","160.151.198.204","160.151.198.205","160.151.198.206","160.151.198.207","160.151.198.208","160.151.198.209","160.151.198.210","160.151.198.211","160.151.198.212","160.151.198.213","160.151.198.214","160.151.198.215","160.151.198.216","160.151.198.217","160.151.198.218","160.151.198.219","160.151.198.220","160.151.198.221","160.151.198.222","160.151.198.223","160.151.198.224","160.151.198.225","160.151.198.226","160.151.198.227","160.151.198.228","160.151.198.229","160.151.198.230","160.151.198.231","160.151.198.232","160.151.198.233","160.151.198.234","160.151.198.235","160.151.198.236","160.151.198.237","160.151.198.238","160.151.198.239","160.151.198.240","160.151.198.241","160.151.198.242","160.151.198.243","160.151.198.244","160.151.198.245","160.151.198.246","160.151.198.247","160.151.198.248","160.151.198.249","160.151.198.250","160.151.198.251","160.151.198.252","160.151.198.253","160.151.198.254","160.151.199.1","160.151.199.2","160.151.199.3","160.151.199.4","160.151.199.5","160.151.199.6","160.151.199.7","160.151.199.8","160.151.199.9","160.151.199.10","160.151.199.11","160.151.199.12","160.151.199.13","160.151.199.14","160.151.199.15","160.151.199.16","160.151.199.17","160.151.199.18","160.151.199.19","160.151.199.20","160.151.199.21","160.151.199.22","160.151.199.23","160.151.199.24","160.151.199.25","160.151.199.26","160.151.199.27","160.151.199.28","160.151.199.29","160.151.199.30","160.151.199.31","160.151.199.32","160.151.199.33","160.151.199.34","160.151.199.35","160.151.199.36","160.151.199.37","160.151.199.38","160.151.199.39","160.151.199.40","160.151.199.41","160.151.199.42","160.151.199.43","160.151.199.44","160.151.199.45","160.151.199.46","160.151.199.47","160.151.199.48","160.151.199.49","160.151.199.50","160.151.199.51","160.151.199.52","160.151.199.53","160.151.199.54","160.151.199.55","160.151.199.56","160.151.199.57","160.151.199.58","160.151.199.59","160.151.199.60","160.151.199.61","160.151.199.62","160.151.199.63","160.151.199.64","160.151.199.65","160.151.199.66","160.151.199.67","160.151.199.68","160.151.199.69","160.151.199.70","160.151.199.71","160.151.199.72","160.151.199.73","160.151.199.74","160.151.199.75","160.151.199.76","160.151.199.77","160.151.199.78","160.151.199.79","160.151.199.80","160.151.199.81","160.151.199.82","160.151.199.83","160.151.199.84","160.151.199.85","160.151.199.86","160.151.199.87","160.151.199.88","160.151.199.89","160.151.199.90","160.151.199.91","160.151.199.92","160.151.199.93","160.151.199.94","160.151.199.95","160.151.199.96","160.151.199.97","160.151.199.98","160.151.199.99","160.151.199.100","160.151.199.101","160.151.199.102","160.151.199.103","160.151.199.104","160.151.199.105","160.151.199.106","160.151.199.107","160.151.199.108","160.151.199.109","160.151.199.110","160.151.199.111","160.151.199.112","160.151.199.113","160.151.199.114","160.151.199.115","160.151.199.116","160.151.199.117","160.151.199.118","160.151.199.119","160.151.199.120","160.151.199.121","160.151.199.122","160.151.199.123","160.151.199.124","160.151.199.125","160.151.199.126","160.151.199.127","160.151.199.128","160.151.199.129","160.151.199.130","160.151.199.131","160.151.199.132","160.151.199.133","160.151.199.134","160.151.199.135","160.151.199.136","160.151.199.137","160.151.199.138","160.151.199.139","160.151.199.140","160.151.199.141","160.151.199.142","160.151.199.143","160.151.199.144","160.151.199.145","160.151.199.146","160.151.199.147","160.151.199.148","160.151.199.149","160.151.199.150","160.151.199.151","160.151.199.152","160.151.199.153","160.151.199.154","160.151.199.155","160.151.199.156","160.151.199.157","160.151.199.158","160.151.199.159","160.151.199.160","160.151.199.161","160.151.199.162","160.151.199.163","160.151.199.164","160.151.199.165","160.151.199.166","160.151.199.167","160.151.199.168","160.151.199.169","160.151.199.170","160.151.199.171","160.151.199.172","160.151.199.173","160.151.199.174","160.151.199.175","160.151.199.176","160.151.199.177","160.151.199.178","160.151.199.179","160.151.199.180","160.151.199.181","160.151.199.182","160.151.199.183","160.151.199.184","160.151.199.185","160.151.199.186","160.151.199.187","160.151.199.188","160.151.199.189","160.151.199.190","160.151.199.191","160.151.199.192","160.151.199.193","160.151.199.194","160.151.199.195","160.151.199.196","160.151.199.197","160.151.199.198","160.151.199.199","160.151.199.200","160.151.199.201","160.151.199.202","160.151.199.203","160.151.199.204","160.151.199.205","160.151.199.206","160.151.199.207","160.151.199.208","160.151.199.209","160.151.199.210","160.151.199.211","160.151.199.212","160.151.199.213","160.151.199.214","160.151.199.215","160.151.199.216","160.151.199.217","160.151.199.218","160.151.199.219","160.151.199.220","160.151.199.221","160.151.199.222","160.151.199.223","160.151.199.224","160.151.199.225","160.151.199.226","160.151.199.227","160.151.199.228","160.151.199.229","160.151.199.230","160.151.199.231","160.151.199.232","160.151.199.233","160.151.199.234","160.151.199.235","160.151.199.236","160.151.199.237","160.151.199.238","160.151.199.239","160.151.199.240","160.151.199.241","160.151.199.242","160.151.199.243","160.151.199.244","160.151.199.245","160.151.199.246","160.151.199.247","160.151.199.248","160.151.199.249","160.151.199.250","160.151.199.251","160.151.199.252","160.151.199.253","160.151.199.254","160.151.205.129","160.151.205.130","160.151.205.131","160.151.205.132","160.151.205.133","160.151.205.134","160.151.205.135","160.151.205.136","160.151.205.137","160.151.205.138","160.151.205.139","160.151.205.140","160.151.205.141","160.151.205.142","160.151.205.143","160.151.205.144","160.151.205.145","160.151.205.146","160.151.205.147","160.151.205.148","160.151.205.149","160.151.205.150","160.151.205.151","160.151.205.152","160.151.205.153","160.151.205.154","160.151.205.155","160.151.205.156","160.151.205.157","160.151.205.158","160.151.205.159","160.151.205.160","160.151.205.161","160.151.205.162","160.151.205.163","160.151.205.164","160.151.205.165","160.151.205.166","160.151.205.167","160.151.205.168","160.151.205.169","160.151.205.170","160.151.205.171","160.151.205.172","160.151.205.173","160.151.205.174","160.151.205.175","160.151.205.176","160.151.205.177","160.151.205.178","160.151.205.179","160.151.205.180","160.151.205.181","160.151.205.182","160.151.205.183","160.151.205.184","160.151.205.185","160.151.205.186","160.151.205.187","160.151.205.188","160.151.205.189","160.151.205.190"
[array]$figIPs = "160.151.203.129","160.151.203.130","160.151.203.131","160.151.203.132","160.151.203.133","160.151.203.134","160.151.203.135","160.151.203.136","160.151.203.137","160.151.203.138","160.151.203.139","160.151.203.140","160.151.203.141","160.151.203.142","160.151.203.143","160.151.203.144","160.151.203.145","160.151.203.146","160.151.203.147","160.151.203.148","160.151.203.149","160.151.203.150","160.151.203.151","160.151.203.152","160.151.203.153","160.151.203.154","160.151.203.155","160.151.203.156","160.151.203.157","160.151.203.158","160.151.203.159","160.151.203.160","160.151.203.161","160.151.203.162","160.151.203.163","160.151.203.164","160.151.203.165","160.151.203.166","160.151.203.167","160.151.203.168","160.151.203.169","160.151.203.170","160.151.203.171","160.151.203.172","160.151.203.173","160.151.203.174","160.151.203.175","160.151.203.176","160.151.203.177","160.151.203.178","160.151.203.179","160.151.203.180","160.151.203.181","160.151.203.182","160.151.203.183","160.151.203.184","160.151.203.185","160.151.203.186","160.151.203.187","160.151.203.188","160.151.203.189","160.151.203.190","160.151.203.191","160.151.203.192","160.151.203.193","160.151.203.194","160.151.203.195","160.151.203.196","160.151.203.197","160.151.203.198","160.151.203.199","160.151.203.200","160.151.203.201","160.151.203.202","160.151.203.203","160.151.203.204","160.151.203.205","160.151.203.206","160.151.203.207","160.151.203.208","160.151.203.209","160.151.203.210","160.151.203.211","160.151.203.212","160.151.203.213","160.151.203.214","160.151.203.215","160.151.203.216","160.151.203.217","160.151.203.218","160.151.203.219","160.151.203.220","160.151.203.221","160.151.203.222","160.151.203.223","160.151.203.224","160.151.203.225","160.151.203.226","160.151.203.227","160.151.203.228","160.151.203.229","160.151.203.230","160.151.203.231","160.151.203.232","160.151.203.233","160.151.203.234","160.151.203.235","160.151.203.236","160.151.203.237","160.151.203.238","160.151.203.239","160.151.203.240","160.151.203.241","160.151.203.242","160.151.203.243","160.151.203.244","160.151.203.245","160.151.203.246","160.151.203.247","160.151.203.248","160.151.203.249","160.151.203.250","160.151.203.251","160.151.203.252","160.151.203.253","160.151.203.254"
[array]$fillmoreIPs = "160.151.202.129","160.151.202.130","160.151.202.131","160.151.202.132","160.151.202.133","160.151.202.134","160.151.202.135","160.151.202.136","160.151.202.137","160.151.202.138","160.151.202.139","160.151.202.140","160.151.202.141","160.151.202.142","160.151.202.143","160.151.202.144","160.151.202.145","160.151.202.146","160.151.202.147","160.151.202.148","160.151.202.149","160.151.202.150","160.151.202.151","160.151.202.152","160.151.202.153","160.151.202.154","160.151.202.155","160.151.202.156","160.151.202.157","160.151.202.158","160.151.202.159","160.151.202.160","160.151.202.161","160.151.202.162","160.151.202.163","160.151.202.164","160.151.202.165","160.151.202.166","160.151.202.167","160.151.202.168","160.151.202.169","160.151.202.170","160.151.202.171","160.151.202.172","160.151.202.173","160.151.202.174","160.151.202.175","160.151.202.176","160.151.202.177","160.151.202.178","160.151.202.179","160.151.202.180","160.151.202.181","160.151.202.182","160.151.202.183","160.151.202.184","160.151.202.185","160.151.202.186","160.151.202.187","160.151.202.188","160.151.202.189","160.151.202.190","160.151.202.191"
[array]$leadIPs = "160.151.202.193","160.151.202.194","160.151.202.195","160.151.202.196","160.151.202.197","160.151.202.198","160.151.202.199","160.151.202.200","160.151.202.201","160.151.202.202","160.151.202.203","160.151.202.204","160.151.202.205","160.151.202.206","160.151.202.207","160.151.202.208","160.151.202.209","160.151.202.210","160.151.202.211","160.151.202.212","160.151.202.213","160.151.202.214","160.151.202.215","160.151.202.216","160.151.202.217","160.151.202.218","160.151.202.219","160.151.202.220","160.151.202.221","160.151.202.222","160.151.202.223","160.151.202.224","160.151.202.225","160.151.202.226","160.151.202.227","160.151.202.228","160.151.202.229","160.151.202.230","160.151.202.231","160.151.202.232","160.151.202.233","160.151.202.234","160.151.202.235","160.151.202.236","160.151.202.237","160.151.202.238","160.151.202.239","160.151.202.240","160.151.202.241","160.151.202.242","160.151.202.243","160.151.202.244","160.151.202.245","160.151.202.246","160.151.202.247","160.151.202.248","160.151.202.249","160.151.202.250","160.151.202.251","160.151.202.252","160.151.202.253","160.151.202.254"
[array]$targetList

#endregion IP Lists

#endregion Global Variables

#region Script Functions

Function Create-MessageBox {
    # Creates Pop-Up Message Boxes to Inform Users of Output
    Param (
        [Parameter(Mandatory=$true)][string]$message,
        [Parameter(Mandatory=$true)][string]$title,
        [Parameter(Mandatory=$false)][System.Windows.Forms.MessageBoxButtons]$buttons="OK",
        [Parameter(Mandatory=$false)][System.Windows.Forms.MessageBoxIcon]$icon="Information"
    )    
    [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
}

Function Check-IfAdmin {
    # Get the ID and security principal of the current user account
    $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
    # Get the security principal for the Administrator role
    $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
    # Check to see if we are currently running "as Administrator" and relauch as admin if not
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

Function Parse-Input ([string]$message, [string]$computers) {     
    # Split inputed text into array of single entries delimiting on whitespace, commas, and semicolons
    [array]$arrComp = (($computers -replace '[,;\s]', ' ').Trim()) -split '\s+'
    # Pass message and array of computers to the Send-Message Function
    Send-Message -Message $message -Computers $arrComp
}

Function Send-Message ([string]$message, [array]$computers) {
    # Main Send Message Function
    # Initialize error counter and failed computer names
    $errCount = 0
    $failCompNames = @()

    # Test to ensure there is message text
    if ($message) {
        # Test for at least one computer has been entered
        if ($computers[0]) {
            # Iterate through array of computer names
            foreach ($computer in $computers) {
                if (!(test-connection -count 1 -Quiet -ComputerName $computer)) {
                    # Test if computer is online; if not, add to error count, if so; continue to send message
                    $errCount += 1
                    $failCompNames += $computer
                }
                else {      
                    # Call msg.exe to send message to each computer in the array
                    # Message will stay up on Computer until clicked away by user or for 24hours whichever comes first
                    Invoke-Expression "msg.exe * /time:86400 /server:$computer `"$message`""
                    #Write-Host $computer
                    #Test exit code from msg.exe and add an error count and the computer name
                    if ($LASTEXITCODE -ne 0) {
                        $errCount += 1
                        $failCompNames += $computer
                    }
                }
            }
            # Test if errors occured
            if ($errCount -ne 0) {                
                # Format the error message for conjugation and grammatical number
                if ($errCount -eq 1) {
                    $errMsgBegin = "$errCount message was"
                    $workGramNum = "workstation is"
                    $compGramNum = "computer"
                    $failNames = $failCompNames -join "`r`n"
                }
                else {
                    $errMsgBegin = "$errCount messages were"
                    $workGramNum = "workstations are"
                    $compGramNum = "computers"
                    $failNames = $failCompNames -join "`r`n"
                }
                #Generate error message
                $errMessage = "$errMsgBegin unable to be sent to the below $compGramNum`n`nIt is possible that the $workGramNum not online, or you do not have sufficient privileges to perform this operation.`n`nFailures List:`n`n$($failNames)"

                #Pop up Send Error Dialog
                #Create-MessageBox -Message $errMessage -Title "Messsage Send Error" -Icon Error
                Create-ErrorForm($errMessage)
            }
            else {
                #Pop up Send Success Dialog
                Create-MessageBox -Message "Message(s) sent successfully" -Title "Success"
            }
        }
        else {
            #Pop up for no computers entered
            Create-MessageBox -message "You must enter at least one computer." -title "No Computers Chosen" -icon Error
        }
    }
    else {
        #Pop up for no message entered
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
    $System_Drawing_Size.Height = 426
    $System_Drawing_Size.Width = 663
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

Function Create-ErrorForm($errorMsg){
    # Add objects for Error Form
    $formError = New-Object System.Windows.Forms.Form
    $richTextBoxError = New-Object System.Windows.Forms.RichTextBox
    # Form for viewing the error information
    $System_Drawing_Size = New-Object System.Drawing.Size
    # Set Size of View Source Form
    $System_Drawing_Size.Height = 200
    $System_Drawing_Size.Width = 290
    $formError.ClientSize = $System_Drawing_Size
    $formError.DataBindings.DefaultDataSourceUpdateMode = 0
    $formError.StartPosition = "CenterScreen"
    $formError.Name = "formError"
    # Set View Source Form Title
    $formError.Text = "Send Message Errors"
    $formError.Icon = [System.IconExtractor]::Extract("imageres.dll", 93, $False)
    # Add Rich Text Box for the Help Text to display in
    $richTextBoxError.Anchor = 15
    $richTextBoxError.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    #Set Location For the Rich Text Box
    $System_Drawing_Point.X = 5
    $System_Drawing_Point.Y = 5
    $richTextBoxError.Location = $System_Drawing_Point
    $richTextBoxError.Name = "richTextBoxError"
    $richTextBoxError.Font = New-Object System.Drawing.Font("Courier New",10)
    $System_Drawing_Size = New-Object System.Drawing.Size
    # Set Rich Text Box Size
    $System_Drawing_Size.Height = 190
    $System_Drawing_Size.Width = 280
    $richTextBoxError.Size = $System_Drawing_Size
    $richTextBoxError.DetectUrls = $False
    $richTextBoxError.ReadOnly = $True
    # Get source from script file and add newline to each array item for formatting
    $richTextBoxError.Text = $errorMsg
    #Add Rich text Box to View Source Form
    $formError.Controls.Add($richTextBoxError)
    # Show View Source Form
    $formError.Show() | Out-Null
}

Function Create-HelpForm {
    # Build Help Form
    $formDirections = New-Object System.Windows.Forms.Form
    $richTextBoxHelp = New-Object System.Windows.Forms.RichTextBox
    $initialFormWindowState = New-Object System.Windows.Forms.FormWindowState
    # Add Objects to Help form
    $formDirections.AutoScroll = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    # Set Size of help Form
    $System_Drawing_Size.Height = 720
    $System_Drawing_Size.Width = 464
    $formDirections.ClientSize = $System_Drawing_Size
    $formDirections.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 750
    $System_Drawing_Size.Width = 480
    $formDirections.MaximumSize = $System_Drawing_Size
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Height = 750
    $System_Drawing_Size.Width = 480
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
    $richTextBoxHelp.BackColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
    $richTextBoxHelp.BorderStyle = 0
    $richTextBoxHelp.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point
    # Set Loctation for Rich Text Box
    $System_Drawing_Point.X = 13
    $System_Drawing_Point.Y = 13
    $richTextBoxHelp.Location = $System_Drawing_Point
    $richTextBoxHelp.Name = "richTextBoxHelp"
    $richTextBoxHelp.Font = New-Object System.Drawing.Font("Courier New",10)
    $richTextBoxHelp.ReadOnly = $True
    $richTextBoxHelp.SelectionProtected = $True
    $richTextBoxHelp.Cursor = [System.Windows.Forms.Cursors]::Default
    $System_Drawing_Size = New-Object System.Drawing.Size
    # Set Size for Rich Text
    $System_Drawing_Size.Height = 720
    $System_Drawing_Size.Width = 439
    $richTextBoxHelp.Size = $System_Drawing_Size
    $richTextBoxHelp.TabIndex = 0
    $richTextBoxHelp.TabStop = $False
    # Set Text For Help form into the Rich Text Box
    $richTextBoxHelp.Text = 'SEND MESSAGE HELP

INTRODUCTION

This script was designed to send popup messages to computers over the network.  This used to be done with net send which was removed from Windows.  In its place, this script uses msg.exe, which was designed to send popup messages to users logged in to the computer.  

SENDING A MESSAGE TO A SINGLE COMPUTER

To send a message to a single computer, enter your message in the first text box under the Send Messages tab, then type the hostname of the computer you wish to send a message in the single-line text box below and then click “Send” or press <Enter>.  If your message is successful, you will be notified.

SENDING A MESSAGE TO MULTIPLE COMPUTERS

To send a message to multiple computers, enter your message in the first text box under the Send Messages tab, then type the hostnames of the computers you wish to send a message (separated by commas, spaces, or semi-colons).  You may also click “Open” and choose a text (*.txt) file with hostnames of computers already stored (separated by commas, spaces, semi-colons, or carriage returns).  The text file will populate the text box and you may add to the list if you so choose.  Click “Send” or press <Enter> when you are ready to send the messages.  If your messages are successful, you will be notified, or you will be given a list of hostnames that did not receive their messages.

VIEWING SCRIPT SOURCE CODE

You may pull up the source code of the script by choosing Help -> View Script from the menu or by pressing Ctrl + E.'

    # Actions when clicking of links in help document
    $richTextBoxHelp.add_LinkClicked(
        {
            Invoke-Expression "start $($_.LinkText)"
        }
    )       
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

$OnLoadForm_StateCorrection= { 
# Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $initialFormWindowState
}

#endregion Script Functions

#----------------------------
#         Form Build
#----------------------------

#region Build Main Form

$System_Drawing_Size = New-Object System.Drawing.Size
# Sets Size of form
$System_Drawing_Size.Height = 500 
$System_Drawing_Size.Width = 423
$form1.ClientSize = $System_Drawing_Size
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets minimumm Size of form (Cannto be  resize lower than this)
$System_Drawing_Size.Height = 530
$System_Drawing_Size.Width = 439
$form1.MinimumSize = $System_Drawing_Size
# Form Setttings
$form1.Icon = [System.IconExtractor]::Extract("imageres.dll", 15, $False)
$form1.StartPosition = "CenterScreen"
$form1.Name = "form1"
$form1.Text = "DAHC Workstation Messaging System"
$form1.AcceptButton = $buttonSend
$form1.CancelButton = $buttonClose

#endregion Build Main Form

#region Build Top Menu

#region Build File Menu

# Add a MenuStrip object for Quitting the form
$menuFileQuit.Text = "&Quit"
$menuFileQuit.ShortcutKeys ="Control, Q" 
$menuFileQuit.add_Click(
    {
        $form1.Close()
    }
)
# Add a MenuStrip object for Importing a text file to import computers names from
$menuFileOpen.Text = "&Import List"
$menuFileOpen.ShortcutKeys = [System.Windows.Forms.Keys]::Control, [System.Windows.Forms.Keys]::O
# Actions for the Import Menu
$menuFileOpen.add_Click(
    {
        $btnChosen = $openFileDialog1.ShowDialog()
        # Sends the chosen text file to be converted to a string
        if ($btnChosen -eq "OK") {
            Parse-TextFile -Path $openFileDialog1.FileName -TextBox $textBoxListComp
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
$menuHelpDirect.add_Click(
    {
        Create-HelpForm
    }
)
# Builds  the View Script Source menu option and its action to open the View Source window
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
$form1.Controls.Add($menu)

#endregion Build Top Menu

#region Build Tab Within Form

#region Build Main Tab Area

# Build Tab Anchor base
$tabControl1.Anchor = 15
$tabControl1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 27
$tabControl1.Location = $System_Drawing_Point
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
# Sets size of Tab Anchor
$System_Drawing_Size.Height = 458
$System_Drawing_Size.Width = 398
$tabControl1.Size = $System_Drawing_Size
$tabControl1.TabIndex = 4
# Add Anchor to Form
$form1.Controls.Add($tabControl1)

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
$System_Drawing_Size.Width =400
$MsgTab.Size = $System_Drawing_Size
$MsgTab.TabIndex = 0
$MsgTab.Text = "Send A Message"
$MsgTab.UseVisualStyleBackColor = $True
# Add tab to Form
$tabControl1.Controls.Add($MsgTab)

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
$buttonClose.add_Click(
    {
        $form1.Close()
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
$buttonSend.add_Click(
    {
        #Sends the contents of the text box to be stripped and turned into an array and then sent
        #Parse-Input -Message $richTextBoxMsg.Text -Computers ($textBoxListComp.Text)
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
$labelMsg.Text = "Type Your Message"
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
$grpListComp.Text = "Set Workstations To Receive Message"
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

# Adds Dunahm IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxDunhamIP.Anchor = 13
$checkBoxDunhamIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 14
$System_Drawing_Point.Y = 60
$checkBoxDunhamIP.Location = $System_Drawing_Point
$checkBoxDunhamIP.Name = "checkBoxDunhamIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 150
$checkBoxDunhamIP.Size = $System_Drawing_Size
$checkBoxDunhamIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxDunhamIP.Text = "Dunham Workstations"
$checkBoxDunhamIP.UseVisualStyleBackColor = $True
# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxDunhamIP)

#-------------------------------------------

# Adds Fillmore IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxFillmoreIP.Anchor = 13
$checkBoxFillmoreIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 14
$System_Drawing_Point.Y = 90
$checkBoxFillmoreIP.Location = $System_Drawing_Point
$checkBoxFillmoreIP.Name = "checkBoxFillmoreIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 150
$checkBoxFillmoreIP.Size = $System_Drawing_Size
$checkBoxFillmoreIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxFillmoreIP.Text = "Fillmore Workstations"
$checkBoxFillmoreIP.UseVisualStyleBackColor = $True
# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxFillmoreIP)

#-------------------------------------------

# Adds FIG IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxFIGIP.Anchor = 13
$checkBoxFIGIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 164
$System_Drawing_Point.Y = 60
$checkBoxFIGIP.Location = $System_Drawing_Point
$checkBoxFIGIP.Name = "checkBoxFIGIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 150
$checkBoxFIGIP.Size = $System_Drawing_Size
$checkBoxFIGIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxFIGIP.Text = "FIG Workstations"
$checkBoxFIGIP.UseVisualStyleBackColor = $True
# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxFIGIP)

#-------------------------------------------

# Adds Letterkenny IP Standard list Checkbox to Auto Entry List Computers Group Area
$checkBoxLEADIP.Anchor = 13
$checkBoxLEADIP.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
# Sets location of Checkbox
$System_Drawing_Point.X = 164
$System_Drawing_Point.Y = 90
$checkBoxLEADIP.Location = $System_Drawing_Point
$checkBoxLEADIP.Name = "checkBoxLEADIP"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 24
$System_Drawing_Size.Width = 200
$checkBoxLEADIP.Size = $System_Drawing_Size
$checkBoxLEADIP.TabIndex = 9
# Sets Text of Checkbox
$checkBoxLEADIP.Text = "Letterkenny Workstations"
$checkBoxLEADIP.UseVisualStyleBackColor = $True
# Adds Checkbox to Area
$grpListComp.Controls.Add($checkBoxLEADIP)

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
$labelManEntListComp.Text = "Type the computer name(s)/IP(s) separated by commas or import a list of names from a text file:"
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
$buttonManEntListComp.add_Click(
    {
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
$buttonManEntClearComp.add_Click(
    {
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

#Check-IfAdmin

# Settings for the Open File Dialog when opening a text file
$openFileDialog1.InitialDirectory = "$env:userprofile\Documents"
$openFileDialog1.Filter = "Text Files (*.txt) | *.txt"
$openFileDialog1.ShowHelp = $True

# Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
# Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)

# Display correctly outside of ISE
[System.Windows.Forms.Application]::EnableVisualStyles()

# Show the Form
$form1.Add_Shown({$form1.Activate()})
$form1.ShowDialog($this)| Out-Null

#endregion Form Actions