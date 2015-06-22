
# AccountModificationUtility Version 5.01
#   v5.02 Added checkbox for validation of only a single object on Tab2
#   v5.01 Turned off the "Correct Values" button unless the sort value is "Found - fail -"
#   v5.0 Added the DISA compare tab and made modifications to the querying of OU for users which increased speed.
#	v4 Added description field and set the highlighting for successful and failed changes on otherPager and Description.
#	v3 fixed logging issue, users not populating in OUs with only one user object, otherPager drop down menu.
#   v2.0 Cleaned up script, changed some of the object queries, modified the check for ActiveDirectory Module




$version = "UserModGUI - Version 5.02"


[string]$scriptFilePath = Split-Path -parent $MyInvocation.MyCommand.Definition
[string]$username = $env:USERNAME
[string]$Date = (Get-Date).tostring("MM.dd.yyy_HH.mm.tt")
[string]$computer = $env:COMPUTERNAME


#Used for GIF location
[string]$scriptDirectory = "\\KTES2292APCS280\ActiveDirectory\Reports\UserModGUI\Common"
[string]$gif = "$scriptDirectory\ETNOSC-sym.gif"
[string]$runLog = "$scriptDirectory\Ver5\RunLog.list"
[string]$errorLog =	"$scriptDirectory\Ver5\ErrorLog.list"


#Location of CSV containing the static DISA information
$StaticDisaCSV = "$scriptDirectory\Ver5\DisaStatic_6-21-2013.csv"
[string]$StaticDisaExportName = ($StaticDisaCSV.Split("\"))[(($StaticDisaCSV.Split("\")).count -1)]

#Validate ActiveDirectory Modules are available
$CheckForActiveDirectoryModule = Get-Module -ListAvailable ActiveDirectory
if(!$CheckForActiveDirectoryModule){

	Try{  
		"$Date  FAIL  Computer:  $computer   User:  $username  "  | Out-File -FilePath $runLog -Append -Encoding ASCII
	}
	Catch{
		#do nothing if it does not have access
	}

	$InstallADModulesMessage = @"
	
The ActiveDirectory Module for Windows PowerShell needs to be enabled.               	
	
To enable, run the following from an admin elevated command prompt on your workstation.	
	
	
Dism.exe /Online /Enable-Feature /FeatureName:RemoteServerAdministrationTools-Roles-AD-Powershell


A restart may be required after adding this feature before first use.
"@


$answer = [System.Windows.Forms.MessageBox]::Show($InstallADModulesMessage)

Exit 
}



Write-Host "Import Active Directory Module"
Import-Module ActiveDirectory
[string]$server = (Get-ADDomainController).hostname
[string]$userFriendlyName = (Get-ADUser $username).name


Try{  
	"$Date  PASS  Computer:  $computer   User:  $username  " + $userFriendlyName | Out-File -FilePath $runLog -Append -Encoding ASCII
}
Catch{
	#do nothing if it does not have access
}



$rootDN = "DC=EUR,DC=DS,DC=ARMY,DC=MIL"


Write-Host "Load Form"

#region Load Forms NameSpace
# load Forms NameSpace

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null


#endregion Load Forms NameSpace

#region Build Form


$StandardFont = New-Object System.Drawing.Font("Microsoft Sans Serif Regular",8,[System.Drawing.FontStyle]::Regular)
$LabelFont = New-Object System.Drawing.Font("Microsoft Sans Serif Regular",8,[System.Drawing.FontStyle]::Bold)


#Root Form
	$objForm = New-Object System.Windows.Forms.Form 
	$objForm.Text = $version 
	$objForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
	$objForm.AutoScaleDimensions = new-object System.Drawing.SizeF(6, 13)
	$objForm.Size = New-Object System.Drawing.Size(970,710) 
	$objForm.StartPosition = "CenterScreen"
	

#userCredentialsLabel
	$userCredentialsLabel = New-Object System.Windows.Forms.Label
	$userCredentialsLabel.Name = "userCredentialsLabel"
	$userCredentialsLabel.Text = "You are running this as: $Env:USERNAME                   " + "Domain Controller:  $server" 
	$userCredentialsLabel.Size = New-Object System.Drawing.Size(700, 13)
	$userCredentialsLabel.Location = new-object System.Drawing.Point(12, 30)
	$userCredentialsLabel.Font = "Microsoft Sans Serif Regular"
	$objForm.Controls.Add($userCredentialsLabel)
	
#etnoscImage
	$etnoscImage = New-Object System.Windows.Forms.PictureBox
	$etnoscImage.Name = "etnoscImage"
	$img = [System.Drawing.Image]::Fromfile($gif)
	$etnoscImage.Image = $img
	$etnoscImage.Size = New-Object System.Drawing.Size(180, 180)
	$etnoscImage.Location = New-Object System.Drawing.Point(755, 54)
	$objForm.Controls.Add($etnoscImage)
	
#browseObjectsTreeView
	$treeView = New-Object System.Windows.Forms.TreeView
	$treeView.Location = New-Object System.Drawing.Point(254, 54)
	$treeView.Name = "browseObjectsTreeView"
	$treeView.Size = New-Object System.Drawing.Size(490, 250)
	$objForm.Controls.Add($treeView)
	
#selectUserListBox
	$userListBox = New-Object System.Windows.Forms.ListBox
	$userListBox.Location = New-Object System.Drawing.Point(12, 54)
	$userListBox.Name = "$userListBox"
	$userListBox.Size = New-Object System.Drawing.Size(233, 613)
	$userListBox.ScrollAlwaysVisible = $true
	$objForm.Controls.Add($userListBox)
	
#region Tabs

    $TabForm = New-Object System.Windows.Forms.TabControl
    $TabForm.Location = New-Object System.Drawing.Point(250, 310)
    $TabForm.Size = New-Object System.Drawing.Size(690,347)
	$objForm.Controls.Add($TabForm)


	#region Tab1
	
    $Tab1 = New-Object System.Windows.Forms.TabPage("Single Object - Manual")               
    $TabForm.Controls.Add($Tab1)

		
	#selectAccountTypeLabel
		$selectAccountTypeLabel = New-Object System.Windows.Forms.Label
		$selectAccountTypeLabel.Name = "selectAccountTypeLabel"
		$selectAccountTypeLabel.Text = "Select Account Type"
		$selectAccountTypeLabel.Size = New-Object System.Drawing.Size(120, 13)
		$selectAccountTypeLabel.Location = new-object System.Drawing.Point(5, 15)
		$selectAccountTypeLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($selectAccountTypeLabel)
	#selectAccountTypeComboBox
		$selectAccountTypeComboBox = New-Object System.Windows.Forms.ComboBox
		$selectAccountTypeComboBox.Size = New-Object System.Drawing.Size(140, 21)
		$selectAccountTypeComboBox.Location = new-object System.Drawing.Point(130, 15)
		$selectAccountTypeComboBox.Items.Add("User") | Out-Null
		$selectAccountTypeComboBox.Items.Add("Service Account") | Out-Null
		$selectAccountTypeComboBox.Items.Add("Privileged Account") | Out-Null
		$selectAccountTypeComboBox.SelectedItem = "User"
		$selectAccountTypeComboBox.Font = "Microsoft Sans Serif Regular"
		$selectAccountTypeComboBox.add_SelectedValueChanged(
		{
			if($selectAccountTypeComboBox.SelectedItem -eq "Privileged Account")
			{
				DisplayPrivilegedAccountInstructions
			}
			elseif($selectAccountTypeComboBox.SelectedItem -eq "Service Account")
			{
				DisplayServiceAccountInstructions
			}
			elseif($selectAccountTypeComboBox.SelectedItem -eq "User")
			{
				DisplayUserAccountInstructions
			}		
		})
		$Tab1.Controls.Add($selectAccountTypeComboBox)
		
		
	#accountDisabledCheckBox
		$accountDisabledCheckBox = New-Object System.Windows.Forms.CheckBox
		$accountDisabledCheckBox.Name = "accountDisabledCheckBox"
		$accountDisabledCheckBox.Text = "Account Disabled"
		$accountDisabledCheckBox.Size = New-Object System.Drawing.Size(117, 17)
		$accountDisabledCheckBox.Location = New-Object System.Drawing.Point(320, 15)
		$accountDisabledCheckBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($accountDisabledCheckBox)
		
	#smartCardRequiredCheckBox
		$smartCardRequiredCheckBox = New-Object System.Windows.Forms.CheckBox
		$smartCardRequiredCheckBox.Name = "smartCardRequiredCheckBox"
		$smartCardRequiredCheckBox.Text = "SmartCard Required"
		$smartCardRequiredCheckBox.Size = New-Object System.Drawing.Size(150, 17)
		$smartCardRequiredCheckBox.Location = New-Object System.Drawing.Point(320, 35)	
		$smartCardRequiredCheckBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($smartCardRequiredCheckBox)


	#nameLabel
		$nameLabel = New-Object System.Windows.Forms.Label
		$nameLabel.Name = "nameLabel"
		$nameLabel.Text = "Name"
		$nameLabel.Size = New-Object System.Drawing.Size(125, 13)
		$nameLabel.Location = new-object System.Drawing.Point(5, 72)
		$nameLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($nameLabel)
	#nameTextBox
		$nameTextBox = New-Object System.Windows.Forms.TextBox
		$nameTextBox.Name = "nameTextBox"
		$nameTextBox.Size = New-Object System.Drawing.Size(208, 20)
		$nameTextBox.Location = new-object System.Drawing.Point(130, 72)
		$nameTextBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($nameTextBox)
	#nameInstructionsLabel
		$nameInstructionsLabel = New-Object System.Windows.Forms.Label
		$nameInstructionsLabel.Name = "nameInstructionsLabel"
	$nameInstructionsLabel.Text = "nameInstructionsLabel"
		$nameInstructionsLabel.Size = New-Object System.Drawing.Size(300, 13)
		$nameInstructionsLabel.Location = new-object System.Drawing.Point(364, 72)
		$nameInstructionsLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($nameInstructionsLabel)


	#displayNameLabel
		$displayNameLabel = New-Object System.Windows.Forms.Label
		$displayNameLabel.Name = "displayNameLabel"
		$displayNameLabel.Text = "Display Name"
		$displayNameLabel.Size = New-Object System.Drawing.Size(125, 13)
		$displayNameLabel.Location = new-object System.Drawing.Point(5, 98)
		$displayNameLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($displayNameLabel)
	#displayNameTextBox
		$displayNameTextBox = New-Object System.Windows.Forms.TextBox
		$displayNameTextBox.Name = "displayNameTextBox"
		$displayNameTextBox.Size = New-Object System.Drawing.Size(208, 20)
		$displayNameTextBox.Location = new-object System.Drawing.Point(130, 98)
		$displayNameTextBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($displayNameTextBox)
	#displayNameInstructionsLabel
		$displayNameInstructionsLabel = New-Object System.Windows.Forms.Label
		$displayNameInstructionsLabel.Name = "displayNameInstructionsLabel"
	$displayNameInstructionsLabel.Text = "displayNameInstructionsLabel"
		$displayNameInstructionsLabel.Size = New-Object System.Drawing.Size(300, 13)
		$displayNameInstructionsLabel.Location = new-object System.Drawing.Point(364, 98)
		$displayNameInstructionsLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($displayNameInstructionsLabel)


	#sAMAccountNameLabel
		$sAMAccountNameLabel = New-Object System.Windows.Forms.Label
		$sAMAccountNameLabel.Name = "sAMAccountNameLabel"
		$sAMAccountNameLabel.Text = "SAM Account Name"
		$sAMAccountNameLabel.Size = New-Object System.Drawing.Size(125, 13)
		$sAMAccountNameLabel.Location = new-object System.Drawing.Point(5, 124)
		$sAMAccountNameLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($sAMAccountNameLabel)
	#sAMAccountNameTextBox
		$sAMAccountNameTextBox = New-Object System.Windows.Forms.TextBox
		$sAMAccountNameTextBox.Name = "sAMAccountNameTextBox"
		$sAMAccountNameTextBox.Size = New-Object System.Drawing.Size(208, 20)
		$sAMAccountNameTextBox.Location = new-object System.Drawing.Point(130, 124)
		$sAMAccountNameTextBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($sAMAccountNameTextBox)
	#sAMAccountNameInstructionsLabel
		$sAMAccountNameInstructionsLabel = New-Object System.Windows.Forms.Label
		$sAMAccountNameInstructionsLabel.Name = "sAMAccountNameInstructionsLabel"
	$sAMAccountNameInstructionsLabel.Text = "sAMAccountNameInstructionsLabel"
		$sAMAccountNameInstructionsLabel.Size = New-Object System.Drawing.Size(300, 13)
		$sAMAccountNameInstructionsLabel.Location = new-object System.Drawing.Point(364, 124)
		$sAMAccountNameInstructionsLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($sAMAccountNameInstructionsLabel)


	#employeeIDLabel
		$employeeIDLabel = New-Object System.Windows.Forms.Label
		$employeeIDLabel.Name = "employeeIDLabel"
		$employeeIDLabel.Text = "EmployeeID"
		$employeeIDLabel.Size = New-Object System.Drawing.Size(125, 13)
		$employeeIDLabel.Location = new-object System.Drawing.Point(5, 150)
		$employeeIDLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($employeeIDLabel)
	#employeeIDTextBox
		$employeeIDTextBox = New-Object System.Windows.Forms.TextBox
		$employeeIDTextBox.Name = "employeeID"
		$employeeIDTextBox.Size = New-Object System.Drawing.Size(208, 20)
		$employeeIDTextBox.Location = new-object System.Drawing.Point(130, 150)
		$employeeIDTextBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($employeeIDTextBox)
	#employeeIDInstructionsLabel
		$employeeIDInstructionsLabel = New-Object System.Windows.Forms.Label
		$employeeIDInstructionsLabel.Name = "employeeIDInstructionsLabel"
	$employeeIDInstructionsLabel.Text = "employeeIDInstructionsLabel"
		$employeeIDInstructionsLabel.Size = New-Object System.Drawing.Size(300, 13)
		$employeeIDInstructionsLabel.Location = new-object System.Drawing.Point(364, 150)
		$employeeIDInstructionsLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($employeeIDInstructionsLabel)


	#userPrincipalNameLabel
		$userPrincipalNameLabel = New-Object System.Windows.Forms.Label
		$userPrincipalNameLabel.Name = "userPrincipalNameLabel"
		$userPrincipalNameLabel.Text = "User Principal Name"
		$userPrincipalNameLabel.Size = New-Object System.Drawing.Size(125, 13)
		$userPrincipalNameLabel.Location = new-object System.Drawing.Point(5, 176)
		$userPrincipalNameLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($userPrincipalNameLabel)
	#userPrincipalNameTextBox
		$userPrincipalNameTextBox = New-Object System.Windows.Forms.TextBox
		$userPrincipalNameTextBox.Name = "userPrincipalNameTextBox"
		$userPrincipalNameTextBox.Size = New-Object System.Drawing.Size(208, 20)
		$userPrincipalNameTextBox.Location = new-object System.Drawing.Point(130, 176)
		$userPrincipalNameTextBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($userPrincipalNameTextBox)
	#userPrincipalNameInstructionsLabel
		$userPrincipalNameInstructionsLabel = New-Object System.Windows.Forms.Label
		$userPrincipalNameInstructionsLabel.Name = "userPrincipalNameInstructionsLabel"
	$userPrincipalNameInstructionsLabel.Text = "userPrincipalNameInstructionsLabel"
		$userPrincipalNameInstructionsLabel.Size = New-Object System.Drawing.Size(300, 13)
		$userPrincipalNameInstructionsLabel.Location = new-object System.Drawing.Point(364, 176)
		$userPrincipalNameInstructionsLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($userPrincipalNameInstructionsLabel)


	#extensionAttribute14Label
		$extensionAttribute14Label = New-Object System.Windows.Forms.Label
		$extensionAttribute14Label.Name = "extensionAttribute14Label"
		$extensionAttribute14Label.Text = "Extension Attribute 14"
		$extensionAttribute14Label.Size = New-Object System.Drawing.Size(125, 13)
		$extensionAttribute14Label.Location = new-object System.Drawing.Point(5, 202)
		$extensionAttribute14Label.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($extensionAttribute14Label)
	#extensionAttribute14TextBox
		$extensionAttribute14TextBox = New-Object System.Windows.Forms.TextBox
		$extensionAttribute14TextBox.Name = "extensionAttribute14"
		$extensionAttribute14TextBox.Size = New-Object System.Drawing.Size(208, 20)
		$extensionAttribute14TextBox.Location = new-object System.Drawing.Point(130, 202)
		$extensionAttribute14TextBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($extensionAttribute14TextBox)
	#extensionAttribute14InstructionsLabel
		$extensionAttribute14InstructionsLabel = New-Object System.Windows.Forms.Label
		$extensionAttribute14InstructionsLabel.Name = "extensionAttribute14InstructionsLabel"
	$extensionAttribute14InstructionsLabel.Text = "extensionAttribute14InstructionsLabel"
		$extensionAttribute14InstructionsLabel.Size = New-Object System.Drawing.Size(300, 13)
		$extensionAttribute14InstructionsLabel.Location = new-object System.Drawing.Point(364, 202)
		$extensionAttribute14InstructionsLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($extensionAttribute14InstructionsLabel)


	#pagerOther
		$pagerOtherLabel = New-Object System.Windows.Forms.Label
		$pagerOtherLabel.Name = "pagerOtherLabel"
		$pagerOtherLabel.Text = "Pager Other"
		$pagerOtherLabel.Size = New-Object System.Drawing.Size(125, 13)
		$pagerOtherLabel.Location = new-object System.Drawing.Point(5, 228)
		$pagerOtherLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($pagerOtherLabel)
	#pagerOtherTextBox
		$pagerOtherTextBox = New-Object System.Windows.Forms.ComboBox
		$pagerOtherTextBox.Name = "pagerOtherTextBox"
		$pagerOtherTextBox.Size = New-Object System.Drawing.Size(208, 20)
		$pagerOtherTextBox.Location = new-object System.Drawing.Point(130, 228)
		$pagerOtherTextBox.Font = "Microsoft Sans Serif Regular"
	#	$pagerOtherTextBox.Enabled = $false
		$pagerOtherTextBox.Items.Add(" ") | Out-Null
		$pagerOtherTextBox.Items.Add("APP-ePO") | Out-Null
		$pagerOtherTextBox.Items.Add("APP-SCCM") | Out-Null
		$pagerOtherTextBox.Items.Add("APP-BAK") | Out-Null	
		$pagerOtherTextBox.Items.Add("APP-RET") | Out-Null	
		$Tab1.Controls.Add($pagerOtherTextBox)

		
	#nameMappingsLabel
		$nameMappingsLabel = New-Object System.Windows.Forms.Label
		$nameMappingsLabel.Name = "nameMappingsLabel"
		$nameMappingsLabel.Text = "AlternateNameMappings"
		$nameMappingsLabel.Size = New-Object System.Drawing.Size(125, 13)
		$nameMappingsLabel.Location = new-object System.Drawing.Point(5, 254)
		$nameMappingsLabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($nameMappingsLabel)
	#nameMappingsTextBox
		$nameMappingsTextBox = New-Object System.Windows.Forms.TextBox
		$nameMappingsTextBox.Name = "nameMappingsTextBox"
		$nameMappingsTextBox.Size = New-Object System.Drawing.Size(208, 20)
		$nameMappingsTextBox.Location = new-object System.Drawing.Point(130, 254)
		$nameMappingsTextBox.Font = "Microsoft Sans Serif Regular"
		$nameMappingsTextBox.Enabled = $false
		$Tab1.Controls.Add($nameMappingsTextBox)

	#userDescriptionlabel
		$userDescriptionlabel = New-Object System.Windows.Forms.Label
		$userDescriptionlabel.Name = "userDescriptionlabel"
		$userDescriptionlabel.Text = "Description"
		$userDescriptionlabel.Size = New-Object System.Drawing.Size(125, 13)
		$userDescriptionlabel.Location = new-object System.Drawing.Point(5, 280)
		$userDescriptionlabel.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($userDescriptionlabel)
	#userDescriptionTextBox
		$userDescriptionTextBox = New-Object System.Windows.Forms.TextBox
		$userDescriptionTextBox.Name = "userDescriptionTextBox"
		$userDescriptionTextBox.Size = New-Object System.Drawing.Size(550, 20)
		$userDescriptionTextBox.Location = new-object System.Drawing.Point(130, 280)
		$userDescriptionTextBox.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($userDescriptionTextBox)


	#commitChangesButton
		$commitChangesButton = New-Object System.Windows.Forms.Button
		$commitChangesButton.Name = "commitChangesButton"
		$commitChangesButton.Text = "Commit Changes"
		$commitChangesButton.Size = new-object System.Drawing.Size(120, 23)
		$commitChangesButton.Location = New-Object System.Drawing.Point(370, 240)
		$commitChangesButton.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($commitChangesButton)

	#goToKbButton
		$goToKbButton = New-Object System.Windows.Forms.Button
		$goToKbButton.Name = "goToKbButton"
		$goToKbButton.Text = "Instructions"
		$goToKbButton.Size = new-object System.Drawing.Size(120, 23)
		$goToKbButton.Location = New-Object System.Drawing.Point(500, 240)
		$goToKbButton.Font = "Microsoft Sans Serif Regular"
		$Tab1.Controls.Add($goToKbButton)

	#endregion Tab1
	

	#region Tab2
	
    $Tab2 = New-Object System.Windows.Forms.TabPage("Based on DISA Export - Automated")               
    $TabForm.Controls.Add($Tab2)


	#Compare
		$compareDISA = New-Object System.Windows.Forms.Button
		$compareDISA.Name = "Compare"
		$compareDISA.Text = "Run Compare"
		$compareDISA.Size = new-object System.Drawing.Size(120, 23)
		$compareDISA.Location = New-Object System.Drawing.Point(12, 15)
		$compareDISA.Font = "Microsoft Sans Serif Regular"
		$Tab2.Controls.Add($compareDISA)
		$compareDISA.BackColor = "LightGreen"
		
	#OnlyValidateSelectedValue
		$OnlyValidateSelectedValue = New-Object System.Windows.Forms.CheckBox
		$OnlyValidateSelectedValue.Name = "OnlyValidateSelectedValue"
		$OnlyValidateSelectedValue.Text = "Only Compare Selected Object"
		$OnlyValidateSelectedValue.Size = new-object System.Drawing.Size(220, 23)
		$OnlyValidateSelectedValue.Location = New-Object System.Drawing.Point(500, 5)
		$Tab2.Controls.Add($OnlyValidateSelectedValue)

	#2CorrectValues
		$2CorrectValues = New-Object System.Windows.Forms.Button
		$2CorrectValues.Name = "CorrectValues"
		$2CorrectValues.Text = "Correct Values"
		$2CorrectValues.Size = new-object System.Drawing.Size(120, 23)
		$2CorrectValues.Location = New-Object System.Drawing.Point(12, 260)
		$2CorrectValues.Font = "Microsoft Sans Serif Regular"
		$Tab2.Controls.Add($2CorrectValues)
		$2CorrectValues.BackColor = "Control"
		$2CorrectValues.Enabled = $false



	#2CurrentValuesLabel
		$2CurrentValuesLabel = New-Object System.Windows.Forms.Label
		$2CurrentValuesLabel.Name = "CurrentValuesLabel"
		$2CurrentValuesLabel.Text = "Current Value"
		$2CurrentValuesLabel.Size = new-object System.Drawing.Size(120, 23)
		$2CurrentValuesLabel.Location = New-Object System.Drawing.Point(200, 20)
		$2CurrentValuesLabel.Font = $LabelFont
		$Tab2.Controls.Add($2CurrentValuesLabel)
		
	#2CorrectValuesLabel
		$2CorrectValuesLabel = New-Object System.Windows.Forms.Label
		$2CorrectValuesLabel.Name = "CorrectValuesLabel"
		$2CorrectValuesLabel.Text = "Correct Value"
		$2CorrectValuesLabel.Size = new-object System.Drawing.Size(120, 23)
		$2CorrectValuesLabel.Location = New-Object System.Drawing.Point(400, 20)
		$2CorrectValuesLabel.Font = $LabelFont
		$Tab2.Controls.Add($2CorrectValuesLabel)	
		
	#2Name
		$2Name = New-Object System.Windows.Forms.Label
		$2Name.Name = "Name"
		$2Name.Text = "Name"
		$2Name.Size = new-object System.Drawing.Size(120, 23)
		$2Name.Location = New-Object System.Drawing.Point(20, 60)
		$2Name.Font = $StandardFont
		$Tab2.Controls.Add($2Name)	
	#2CurrentName
		$2CurrentName = New-Object System.Windows.Forms.Label
		$2CurrentName.Name = "CurrentName"
#		$2CurrentName.Text = "Name"
		$2CurrentName.Size = new-object System.Drawing.Size(200, 23)
		$2CurrentName.Location = New-Object System.Drawing.Point(200, 60)
		$2CurrentName.Font = $StandardFont
		$Tab2.Controls.Add($2CurrentName)	
	#2CorrectName
		$2CorrectName = New-Object System.Windows.Forms.Label
		$2CorrectName.Name = "CorrectName"
#		$2CorrectName.Text = "Name"
		$2CorrectName.Size = new-object System.Drawing.Size(200, 23)
		$2CorrectName.Location = New-Object System.Drawing.Point(400, 60)
		$2CorrectName.Font = $StandardFont
		$Tab2.Controls.Add($2CorrectName)	
		
		
	#2SamAccountName
		$2SamAccountName = New-Object System.Windows.Forms.Label
		$2SamAccountName.Name = "SamAccountName"
		$2SamAccountName.Text = "SamAccountName"
		$2SamAccountName.Size = new-object System.Drawing.Size(120, 23)
		$2SamAccountName.Location = New-Object System.Drawing.Point(20, 100)
		$2SamAccountName.Font = $StandardFont
		$Tab2.Controls.Add($2SamAccountName)
	#2CurrentSamAccountName
		$2CurrentSamAccountName = New-Object System.Windows.Forms.Label
		$2CurrentSamAccountName.Name = "CurrentSamAccountName"
#		$2CurrentSamAccountName.Text = "CurrentSamAccountName"
		$2CurrentSamAccountName.Size = new-object System.Drawing.Size(200, 23)
		$2CurrentSamAccountName.Location = New-Object System.Drawing.Point(200, 100)
		$2CurrentSamAccountName.Font = $StandardFont
		$Tab2.Controls.Add($2CurrentSamAccountName)
	#2CorrectAccountName
		$2CorrectSamAccountName = New-Object System.Windows.Forms.Label
		$2CorrectSamAccountName.Name = "2CorrectAccountName"
#		$2CorrectSamAccountName.Text = "2CorrectAccountName"
		$2CorrectSamAccountName.Size = new-object System.Drawing.Size(200, 23)
		$2CorrectSamAccountName.Location = New-Object System.Drawing.Point(400, 100)
		$2CorrectSamAccountName.Font = $StandardFont
		$Tab2.Controls.Add($2CorrectSamAccountName)
		

	#2EmployeeID
		$2EmployeeID = New-Object System.Windows.Forms.Label
		$2EmployeeID.Name = "EmployeeID"
		$2EmployeeID.Text = "EmployeeID"
		$2EmployeeID.Size = new-object System.Drawing.Size(120, 23)
		$2EmployeeID.Location = New-Object System.Drawing.Point(20, 140)
		$2EmployeeID.Font = $StandardFont
		$Tab2.Controls.Add($2EmployeeID)			
	#2CurrentEmployeeID
		$2CurrentEmployeeID = New-Object System.Windows.Forms.Label
		$2CurrentEmployeeID.Name = "2CurrentEmployeeID"
#		$2CurrentEmployeeID.Text = "2CurrentEmployeeID"
		$2CurrentEmployeeID.Size = new-object System.Drawing.Size(200, 23)
		$2CurrentEmployeeID.Location = New-Object System.Drawing.Point(200, 140)
		$2CurrentEmployeeID.Font = $StandardFont
		$Tab2.Controls.Add($2CurrentEmployeeID)	
	#2CorrectEmployeeID
		$2CorrectEmployeeID = New-Object System.Windows.Forms.Label
		$2CorrectEmployeeID.Name = "2CorrectEmployeeID"
#		$2CorrectEmployeeID.Text = "2CorrectEmployeeID"
		$2CorrectEmployeeID.Size = new-object System.Drawing.Size(200, 23)
		$2CorrectEmployeeID.Location = New-Object System.Drawing.Point(400, 140)
		$2CorrectEmployeeID.Font = $StandardFont
		$Tab2.Controls.Add($2CorrectEmployeeID)	

	#2UserPrincipalname
		$2UserPrincipalName = New-Object System.Windows.Forms.Label
		$2UserPrincipalName.Name = "UserPrincipalName"
		$2UserPrincipalName.Text = "UserPrincipalName"
		$2UserPrincipalName.Size = new-object System.Drawing.Size(120, 23)
		$2UserPrincipalName.Location = New-Object System.Drawing.Point(20, 180)
		$2UserPrincipalName.Font = $StandardFont
		$Tab2.Controls.Add($2UserPrincipalName)	
	#2CurrentUserPrincipalname
		$2CurrentUserPrincipalname = New-Object System.Windows.Forms.Label
		$2CurrentUserPrincipalname.Name = "2CurrentUserPrincipalname"
#		$2CurrentUserPrincipalname.Text = "2CurrentUserPrincipalname"
		$2CurrentUserPrincipalname.Size = new-object System.Drawing.Size(200, 23)
		$2CurrentUserPrincipalname.Location = New-Object System.Drawing.Point(200, 180)
		$2CurrentUserPrincipalname.Font = $StandardFont
		$Tab2.Controls.Add($2CurrentUserPrincipalname)	
	#2CorrectUserPrincipalname
		$2CorrectUserPrincipalname = New-Object System.Windows.Forms.Label
		$2CorrectUserPrincipalname.Name = "2CorrectUserPrincipalname"
#		$2CorrectUserPrincipalname.Text = "2CorrectUserPrincipalname"
		$2CorrectUserPrincipalname.Size = new-object System.Drawing.Size(200, 23)
		$2CorrectUserPrincipalname.Location = New-Object System.Drawing.Point(400, 180)
		$2CorrectUserPrincipalname.Font = $StandardFont
		$Tab2.Controls.Add($2CorrectUserPrincipalname)	
		
	#2ExtensionAttribute14
		$2ExtensionAttribute14 = New-Object System.Windows.Forms.Label
		$2ExtensionAttribute14.Name = "ExtensionAttribute14"
		$2ExtensionAttribute14.Text = "ExtensionAttribute14"
		$2ExtensionAttribute14.Size = new-object System.Drawing.Size(120, 23)
		$2ExtensionAttribute14.Location = New-Object System.Drawing.Point(20, 220)
		$2ExtensionAttribute14.Font = $StandardFont
		$Tab2.Controls.Add($2ExtensionAttribute14)			
	#2CurrentExtensionAttribute14
		$2CurrentExtensionAttribute14 = New-Object System.Windows.Forms.Label
		$2CurrentExtensionAttribute14.Name = "2CurrentExtensionAttribute14"
#		$2CurrentExtensionAttribute14.Text = "2CurrentExtensionAttribute14"
		$2CurrentExtensionAttribute14.Size = new-object System.Drawing.Size(200, 23)
		$2CurrentExtensionAttribute14.Location = New-Object System.Drawing.Point(200, 220)
		$2CurrentExtensionAttribute14.Font = $StandardFont
		$Tab2.Controls.Add($2CurrentExtensionAttribute14)	
	#2CorrectExtensionAttribute14
		$2CorrectExtensionAttribute14 = New-Object System.Windows.Forms.Label
		$2CorrectExtensionAttribute14.Name = "2CorrectExtensionAttribute14"
#		$2CorrectExtensionAttribute14.Text = "2CorrectExtensionAttribute14"
		$2CorrectExtensionAttribute14.Size = new-object System.Drawing.Size(200, 23)
		$2CorrectExtensionAttribute14.Location = New-Object System.Drawing.Point(400, 220)
		$2CorrectExtensionAttribute14.Font = $StandardFont
		$Tab2.Controls.Add($2CorrectExtensionAttribute14)					
		

		
	#2StatusBar
		$2StatusBar = New-Object System.Windows.Forms.Label
		$2StatusBar.Name = "2StatusBar"
		$2StatusBar.Text = ""
		$2StatusBar.Size = new-object System.Drawing.Size(500, 43)
		$2StatusBar.Location = New-Object System.Drawing.Point(220, 122)
		$2StatusBar.Font = $LabelFont
		$Tab2.Controls.Add($2StatusBar)	

	#22StatusBar
		$22StatusBar = New-Object System.Windows.Forms.Label
		$22StatusBar.Name = "22StatusBar"
		$22StatusBar.Text = ""
		$22StatusBar.Size = new-object System.Drawing.Size(400, 23)
		$22StatusBar.Location = New-Object System.Drawing.Point(220, 164)
		$22StatusBar.Font = $LabelFont
		$Tab2.Controls.Add($22StatusBar)	
		
	#2DISADataDetails
		$2DISADataDetails = New-Object System.Windows.Forms.Label
		$2DISADataDetails.Name = "2DISADataDetails"
		$2DISADataDetails.Text = "The compare data is based on a DISA export titled `"$StaticDisaExportName`""
		$2DISADataDetails.Size = new-object System.Drawing.Size(600, 23)
		$2DISADataDetails.Location = New-Object System.Drawing.Point(150, 260)
		$2DISADataDetails.Font = $LabelFont
		$Tab2.Controls.Add($2DISADataDetails)			
		

	#endregion Tab2

#endregion Tabs
	

	
#endregion Build Form		

		
#region Functions

#DisplayUserAccountInstructions
Function DisplayUserAccountInstructions{
	$nameInstructionsLabel.Text = "[Enterprise Username].[Persona Type Code]"
	$displayNameInstructionsLabel.Text = "[Lastname], [Firstname] [PTC] *Managed by EDS-Lite"
	$sAMAccountNameInstructionsLabel.Text = "[EDIPI].[Persona Type Code]"
	$employeeIDInstructionsLabel.Text = "[EDIPI] *Managed by EDS-Lite on the NIPRNet"
	$userPrincipalNameInstructionsLabel.Text = "[EDIPI]@mil"
	$extensionAttribute14InstructionsLabel.Text = "[Enterprise Email address of user]"
}

#DisplayPrivilegedAccountInstructions
Function DisplayPrivilegedAccountInstructions{
	$nameInstructionsLabel.Text = "[Enterprise Username].[Privileged Account Code]"
	$displayNameInstructionsLabel.Text = "[Enterprise Username].[Privileged Account Code]"
	$sAMAccountNameInstructionsLabel.Text = "[EDIPI from normal user account].[Privileged Account Code]"
	$employeeIDInstructionsLabel.Text = "[EDIPI from normal user account]"
	$userPrincipalNameInstructionsLabel.Text = "[ASCL or PIV Token Serial Number]@mil"
	$extensionAttribute14InstructionsLabel.Text = "[Enterprise Email address from normal user account]"
}

#DisplayServiceAccountInstructions
Function DisplayServiceAccountInstructions{
	$nameInstructionsLabel.Text = "[SVC].[Service Type][Organizational Code]"
	$displayNameInstructionsLabel.Text = "[SVC].[Service Type][Organizational Code]"
	$sAMAccountNameInstructionsLabel.Text = "[SVC].[Service Type][Organizational Code]"
	$employeeIDInstructionsLabel.Text = "Empty"
	$userPrincipalNameInstructionsLabel.Text = "Empty"
	$extensionAttribute14InstructionsLabel.Text = "[Enterprise Email address of account manager]"
}

#for tab 1
Function ResetTextBoxDefaults{
	$accountDisabledCheckBox.Checked = $false
	$smartCardRequiredCheckBox.Checked = $false
	$nameTextBox.Text = $null
	$displayNameTextBox.Text = $null
	$sAMAccountNameTextBox.Text = $null
	$employeeIDTextBox.Text = $null
	$userPrincipalNameTextBox.Text = $null
	$extensionAttribute14TextBox.Text = $null
	$pagerOtherTextBox.Text = $null
	$nameMappingsTextBox.Text = $null
	
	$nameLabel.BackColor = "Control"
	$displayNameLabel.BackColor = "Control"
	$sAMAccountNameLabel.BackColor = "Control"
	$employeeIDLabel.BackColor = "Control"
	$userPrincipalNameLabel.BackColor = "Control"
	$extensionAttribute14Label.BackColor = "Control"
	$pagerOtherLabel.BackColor = "Control"
	$userDescriptionlabel.BackColor = "Control"
	
	$commitChangesButton.BackColor = "Control"
}

#for tab 2
Function ResetTab2Labels{

	$2CurrentEmployeeID.Text = ""
	$2CurrentEmployeeID.BackColor = "Control"
	$2CurrentExtensionAttribute14.Text = ""
	$2CurrentExtensionAttribute14.BackColor = "Control"
	$2CurrentName.Text = ""
	$2CurrentName.BackColor = "Control"
	$2CurrentSamAccountName.Text = ""
	$2CurrentSamAccountName.BackColor = "Control"
	$2CurrentUserPrincipalname.Text = ""
	$2CurrentUserPrincipalname.BackColor = "Control"
	
	
	$2CorrectEmployeeID.Text = ""
	$2CorrectExtensionAttribute14.Text = ""
	$2CorrectName.Text = ""
	$2CorrectSamAccountName.Text = ""
	$2CorrectUserPrincipalName.Text = ""
	
	$commitChangesButton.BackColor = "Control"
	$2CorrectValues.Enabled = $false
	$2CorrectValues.BackColor = "Control"
}

#for tab 1
Function PopulateSelectedUsersValues{
	Try{
		$currentUser = Get-ADUser ($userListBox.SelectedItem).ObjectGUID -Properties * -Server $server
	}
	Catch{
		#do nothing
	}
	if($currentUser.Enabled -eq $false){$accountDisabledCheckBox.Checked = $true}
	if($currentUser.SmartcardLogonRequired -eq $true){$smartCardRequiredCheckBox.Checked = $True}
		
		
	$nameTextBox.Text = $currentUser.Name
	$displayNameTextBox.Text = $currentUser.DisplayName
	$sAMAccountNameTextBox.Text = $currentUser.SamAccountName
	$employeeIDTextBox.Text = $currentUser.EmployeeID
	$userPrincipalNameTextBox.Text = $currentUser.UserPrincipalName
	$extensionAttribute14TextBox.Text = $currentUser.extensionAttribute14
	$pagerOtherTextBox.Text = $currentUser.otherPager
	$userDescriptionTextBox.Text = $currentUser.description
		
		try
		{
			$altSecSplit = ($currentUser.altSecurityIdentities.Value).split(",") 
			$dodCA = ($altSecSplit[4]).Substring(3,9)  
			$nameMappingsTextBox.Text = $dodCA	
		}
		catch
		{
			#do nothing
		}
}

#for tab 2
Function PopulateTab2values {

	
	$DisaObject = $script:AutoItems | where {$_.SortValue -eq $userListBox.SelectedItem}

	$currentObject = Get-ADUser $DisaObject.ObjectGUID -Properties extensionAttribute14,employeeID
	
	$2CorrectValues.enabled = $false	
	$2CurrentEmployeeID.Text = $currentObject.employeeID
	$2CurrentExtensionAttribute14.Text = $currentObject.extensionAttribute14
	$2CurrentName.Text = $currentObject.Name
	$2CurrentSamAccountName.Text = $currentObject.SamAccountName
	$2CurrentUserPrincipalname.Text = $currentObject.UserPrincipalName

	if($DisaObject.SortValue -match "NotFound -"){
		$2CurrentEmployeeID.BackColor = "Yellow"
		$2CurrentExtensionAttribute14.BackColor = "Yellow"
		$2CurrentName.BackColor = "Yellow"
		$2CurrentSamAccountName.BackColor = "Yellow"
		$2CurrentUserPrincipalname.BackColor = "Yellow"
		return
	}
	
	if($DisaObject.SortValue -like "Found - fail -*"){
		$2CorrectValues.enabled = $true
	}

	
	$2CorrectEmployeeID.Text = $DisaObject.EDIPI
	$2CorrectExtensionAttribute14.Text = $DisaObject.DisaMail
	$2CorrectName.Text = $DisaObject.CorrectName
	$2CorrectSamAccountName.Text = $DisaObject.CorrectSamAccount
	$2CorrectUserPrincipalName.Text = $DisaObject.DisaUPN
	
	
	#validate values and color code
	if($currentObject.employeeID -ne $DisaObject.EDIPI){
		$2CurrentEmployeeID.BackColor = "PaleVioletRed"
	}
	if($currentObject.extensionAttribute14 -ne $DisaObject.DisaMail){
		$2CurrentExtensionAttribute14.BackColor = "PaleVioletRed"
	}
	if($currentObject.Name -ne $DisaObject.CorrectName){
		$2CurrentName.BackColor = "PaleVioletRed"
	}
	if($currentObject.SamAccountName -ne $DisaObject.CorrectSamAccount){
		$2CurrentSamAccountName.BackColor = "PaleVioletRed"
	}
	if($currentObject.UserPrincipalName -ne $DisaObject.DisaUPN){
		$2CurrentUserPrincipalname.BackColor = "PaleVioletRed"
	}
	
	$2CorrectValues.BackColor = "LightGreen"
}

Function Validate-EURUser ([Object]$userObject, [Object]$compareValues){

	Try{
		if($userObject.SamAccountName -ne $compareValues.CorrectSamAccount){return "fail"}
		if($userObject.Name -ne $compareValues.CorrectName){return "fail"}
		if($userObject.employeeID -ne (($compareValues.DisaUPN).substring(0,10))){return "fail"}
		if($userObject.extensionAttribute14 -ne $compareValues.DisaMail){return "fail"}
	}
	Catch{
		return "fail"
	}
	return "pass"
}

#for tab 2
Function Correct-Tab2Values {

	$CurrentObject = $script:AutoItems | where {$_.SortValue -eq $userListBox.SelectedItem}
	
	if($2CurrentEmployeeID.Text -ne $2CorrectEmployeeID.Text){
		Set-ADUser -Identity $currentObject.ObjectGUID -EmployeeID $2CorrectEmployeeID.Text -Server $server
	}
	if($2CurrentExtensionAttribute14.Text -ne $2CorrectExtensionAttribute14.Text){
		Set-ADObject -Identity $currentObject.ObjectGUID -Replace @{extensionattribute14 = ($2CorrectExtensionAttribute14.Text)} -Server $server
	}
	if($2CurrentName.Text -ne $2CorrectName.Text){
		Rename-ADObject -Identity $currentObject.ObjectGUID $2CorrectName.Text -Server $server
	}
	if($2CurrentSamAccountName.Text -ne $2CorrectSamAccountName.Text){
		Set-ADUser -Identity $currentObject.ObjectGUID -SamAccountName $2CorrectSamAccountName.Text -Server $server
	}
	if($2CurrentUserPrincipalname.Text -ne $2CorrectUserPrincipalname.Text){
		Set-ADUser -Identity $currentObject.ObjectGUID -UserPrincipalName $2CorrectUserPrincipalname.Text -Server $server 
	}

}

#endregion Functions


#Region Form Actions

$treeView.add_AfterSelect(
	{
	
		#Reset the tab values
		$2StatusBar.Text = ""
		$22StatusBar.Text = ""
		$Tab1.Enabled = $true
		ResetTextBoxDefaults
		ResetTab2Labels
		
		#Turn on compare button
		$compareDISA.BackColor = "LightGreen"
		$compareDISA.Enabled = $true
		
		$userListBox.Items.Clear()
		
		if ($treeView.SelectedNode.Tag -eq "Contains Users")
		{
			[array]$userObjects = Get-ADUser -filter * -SearchBase $treeView.SelectedNode.Name -SearchScope OneLevel -Server $server -Properties employeeID,extensionAttribute14
			$userListBox.Items.AddRange($userObjects)
			$userListBox.ValueMember = "name"
		}
		
		if ($treeView.SelectedNode.Checked -eq $false)
		{
			try
			{
			[array]$subOUs = $null; $subOUs = Get-ChildItem -Path ("AD:\" + $treeView.SelectedNode.Name) | Where-Object {$_.objectClass -eq "organizationalUnit"}  
			$treeView.SelectedNode.Checked = $true
			$value = $null
				foreach ($value in $subOUs)
				{
					$treeView.SelectedNode.Nodes.Add($value.name) | Out-Null
					foreach($node in $treeView.SelectedNode.Nodes)
					{
						if ($node.text -eq $value.name)
						{
							$node.Name = $value.distinguishedName
							[array]$usersInOU = Get-ADUser -filter * -SearchBase $value.distinguishedName -SearchScope OneLevel
							[int]$usersCount = $usersInOU.count	
							if($usersCount -ne 0)
							{
								$node.text = $node.Text + " ($usersCount)" 
								$node.Tag = "Contains Users"
							}
						}
					}			
				}
			}
			catch
			{	
				#[int]$usersCount = (Get-ADUser -filter * -SearchBase $treeView.SelectedNode.Name -SearchScope OneLevel).count
			}
		}

	})

$userListBox.add_SelectedValueChanged(
	{
		$2StatusBar.Text = ""
		$22StatusBar.Text = ""
		
		if($TabForm.SelectedTab.TabIndex -eq 0){
			ResetTextBoxDefaults
			PopulateSelectedUsersValues
		}
		if($TabForm.SelectedTab.TabIndex -eq 1){
		#The CompareDISA button should be disabled if user objects have been validated.	
			if($compareDISA.Enabled -eq $false){
				ResetTab2Labels
				PopulateTab2values
			}	
		}
		
	})

$commitChangesButton.add_MouseClick(
	{
		$index = $userListBox.SelectedIndex 
		$currentUserObject = Get-ADUser ($userListBox.SelectedItem).ObjectGUID -Properties * -Server $server
		[string]$flag1 = $null;[string]$flag2 = $null;[string]$flag3 = $null;[string]$flag4 = $null;[string]$flag5 = $null;[string]$flag6 = $null;[string]$flag7 = $null;[string]$flag8 = $null	
		
		#Account Enabled
			if(($accountDisabledCheckBox.Checked -eq $false) -and ($currentUserObject.Enabled -eq $false)){Set-ADUser $currentUserObject.DistinguishedName -Enabled $true -Server $server}
			if(($accountDisabledCheckBox.Checked -eq $true) -and ($currentUserObject.Enabled -eq $true)){Set-ADUser $currentUserObject.distinguishedName -Enabled $false -Server $server}
		
		#SmartCard Required
			if(($smartCardRequiredCheckBox.Checked -eq $false) -and ($currentUserObject.SmartcardLogonRequired -eq $true)){Set-ADUser $currentUserObject.distinguishedName -SmartcardLogonRequired $False -Server $server}
			if(($smartCardRequiredCheckBox.Checked -eq $true) -and ($currentUserObject.SmartcardLogonRequired -eq $False)){Set-ADUser $currentUserObject.distinguishedName -SmartcardLogonRequired $True -Server $server}
		
		#DisplayName
			try
			{
				if($displayNameTextBox.Text -ne $currentUserObject.DisplayName)
				{
					Set-ADUser $currentUserObject.DistinguishedName -DisplayName $displayNameTextBox.Text -Server $server 
					$flag1 = "worked"		
				}
			}
			catch
			{
				Write-Host "Unable to write DisplayName to " $currentUserObject.Name 
				$flag1 = "failed"
			}
		
		#sAMAccountName
			try
			{
				if($sAMAccountNameTextBox.Text -ne $currentUserObject.SamAccountName)
				{
					Set-ADUser $currentUserObject.DistinguishedName -SamAccountName $sAMAccountNameTextBox.Text -Server $server
					$flag2 = "worked"
				}
			}
			catch
			{
				Write-Host "Unable to write sAMAccountName to " $currentUserObject.Name
				$flag2 = "failed"
			}
		
		#UserPrincipalName	
			try
			{
				if($userPrincipalNameTextBox.Text -ne $currentUserObject.UserPrincipalName)
				{
					Set-ADUser $currentUserObject.DistinguishedName -UserPrincipalName $userPrincipalNameTextBox.Text -Server $server 
					$flag3 = "worked"
				}
			}
			catch
			{
				Write-Host "Unable to write UserPrincipalName to " $currentUserObject.Name
				$flag3 = "failed"
			}
		
		#EmployeeID
			try
			{
				if($employeeIDTextBox.Text -eq "")
				{
					try{Set-ADUser $currentUserObject.DistinguishedName -EmployeeID $null -Server $server}
					catch{$flag4 = "failed"}
				}
				elseif($employeeIDTextBox.Text -ne $currentUserObject.EmployeeID)
				{
					try{Set-ADUser $currentUserObject.DistinguishedName -EmployeeID $employeeIDTextBox.Text -Server $server; $flag4 = "worked"}
					catch{$flag4 = "failed"}			
				}
			}
			catch{Write-Host "Unable to write EmployeeID to " $currentUserObject.Name; $flag4 = "failed"}
		
		#ExtensionAttribute14
			try
			{
				[string]$userDN = $null; $userDN = $currentUserObject.distinguishedName
				if($extensionAttribute14TextBox.Text -eq "")
				{
					Set-ADObject -Identity $currentUserObject.distinguishedName -Clear extensionattribute14 -Server $server
					$flag5 = "failed"
				}
				elseif($extensionAttribute14TextBox.Text -ne $currentUserObject.extensionAttribute14)
				{
					Set-ADObject -Identity $currentUserObject.distinguishedName -Replace @{extensionattribute14 = ($extensionAttribute14TextBox.Text)} -Server $server
					$flag5 = "worked"
				}
			}	
			catch{Write-Host "Unable to write extensionAttribute14 to " $currentUserObject.Name; $flag5 = "failed"}	


		#otherPager
			
			Switch ($pagerOtherTextBox.Text)
			{
				" "			{Set-ADObject -Identity $currentUserObject.distinguishedName -clear otherPager -Server $server; $flag7 = "worked"} 
				"APP-ePO"	{Set-ADObject -Identity $currentUserObject.distinguishedName -Replace @{otherPager = "APP-ePO"} -Server $server; $flag7 = "worked"}
				"APP-SCCM"	{Set-ADObject -Identity $currentUserObject.distinguishedName -Replace @{otherPager = "APP-SCCM"} -Server $server; $flag7 = "worked"}
				"APP-BAK"  	{Set-ADObject -Identity $currentUserObject.distinguishedName -Replace @{otherPager = "APP-BAK"} -Server $server; $flag7 = "worked"}
				"APP-RET"	{Set-ADObject -Identity $currentUserObject.distinguishedName -Replace @{otherPager = "APP-RET"} -Server $server; $flag7 = "worked"}
			}

		#UserDescription	
			try
			{
				if($userDescriptionTextBox.Text -ne $currentUserObject.description)
				{
					Set-ADUser $currentUserObject.distinguishedName -Description $userDescriptionTextBox.Text -Server $server 
					$flag8 = "worked"
				}
			}
			catch
			{
				Write-Host "Unable to write description to " $currentUserObject.Name
				$flag8 = "failed"
			}


		#Name
		try{if($nameTextBox.Text -ne $currentUserObject.name){Rename-ADObject $currentUserObject.DistinguishedName $nameTextBox.Text -Server $server; $flag6 = "worked"}}catch{Write-Host "Unable to write name to " $currentUserObject.Name; $flag6 = "failed"}


		#Refresh ListBox
		#	$userListBox.Items.Clear()
		#	[array]$userObjects = $null; $userObjects = Get-ADUser -filter * -SearchBase $treeView.SelectedNode.Name -SearchScope OneLevel -Server $server -Properties CN,Name,DisplayName,sAMAccountName,EmployeeID,UserPrincipalName,extensionAttribute14,SmartcardLogonRequired,Enabled,otherPager,altSecurityIdentities,description
		#	$userListBox.Items.AddRange($userObjects)
		#	$userListBox.ValueMember = "name"
			
		#	$userListBox.SelectedIndex = $index
		#	PopulateSelectedUsersValues
			$commitChangesButton.BackColor = "SteelBlue"
			





			#displayName
				if ($flag1 -eq "worked"){$displayNameLabel.BackColor = "DarkSeaGreen"}
				elseif ($flag1 -eq "failed"){$displayNameLabel.BackColor = "Crimson"}
			#samAccountName
				if ($flag2 -eq "worked"){$sAMAccountNameLabel.BackColor = "DarkSeaGreen"}
				elseif ($flag2 -eq "failed"){$sAMAccountNameLabel.BackColor = "Crimson"}
			#userPrincipalName
				if ($flag3 -eq "worked"){$userPrincipalNameLabel.BackColor = "DarkSeaGreen"}
				elseif ($flag3 -eq "failed"){$userPrincipalNameLabel.BackColor = "Crimson"}
			#employeeID
				if ($flag4 -eq "worked"){$employeeIDLabel.BackColor = "DarkSeaGreen"}
				elseif ($flag4 -eq "failed"){$employeeIDLabel.BackColor = "Crimson"}
			#extensionAttribute14
				if ($flag5 -eq "worked"){$extensionAttribute14Label.BackColor = "DarkSeaGreen"}
				elseif ($flag5 -eq "failed"){$extensionAttribute14Label.BackColor = "Crimson"}
			#name
				if ($flag6 -eq "worked"){$nameLabel.BackColor = "DarkSeaGreen"}
				elseif ($flag6 -eq "failed"){$nameLabel.BackColor = "Crimson"}
			#pagerOther
				if ($flag7 -eq "worked"){$pagerOtherLabel.BackColor = "DarkSeaGreen"}
				elseif ($flag7 -eq "failed"){$pagerOtherLabel.BackColor = "Crimson"}
			#Description
				if ($flag8 -eq "worked"){$userDescriptionlabel.BackColor = "DarkSeaGreen"}
				elseif ($flag8 -eq "failed"){$userDescriptionlabel.BackColor = "Crimson"}	

	})

$2CorrectValues.add_MouseClick(
	{
		Correct-Tab2Values
		ResetTab2Labels
		PopulateTab2Values
	})	
	
$compareDISA.add_MouseClick(
	{
		$2StatusBar.Text = "Please wait while the values are validated."
		$22StatusBar.Text = "This can take a few seconds per user. There are " + $userListBox.Items.Count + " users selected."
		Start-Sleep -Seconds 2
		#Turn off compare button
		$compareDISA.BackColor = "Control"
		$compareDISA.Enabled = $false
		
		if($userListBox.Items -ne $null){
		
		$Tab1.Enabled = $false
			if($DISA -ne $null){
				Write-Host "Already Imported"
			}	
			else{
				$script:DISA = Import-Csv $StaticDisaCSV
				Write-Host "DISA Import"
				$DISA = $DISA | Sort-Object -Property "ObjectGUID"
			}
			[array]$script:AutoItems = $null
		
		#If only validating one object
		if($OnlyValidateSelectedValue.Checked -eq $true){
			$item = $null
			$item = $userListBox.SelectedItem
			if($item -eq $null){
				return
			}
			Write-Host "Only validating a single object $item"
			$foundItem = $DISA | where {$_.ObjectGUID -eq (($item).ObjectGUID).Guid}
			if($foundItem -ne $null){
				$result = Validate-EURUser ($item)($foundItem)
				$foundItem | Add-Member -MemberType NoteProperty -Force -Name SortValue -Value ("Found - $result - " + $item.Name)
				$script:AutoItems += $foundItem
			}
			else{
				$CurrentUser = New-Object PSObject
				$CurrentUser | Add-Member -MemberType NoteProperty -Name FlagType -Value ""
				$CurrentUser | Add-Member -MemberType NoteProperty -Name EurSamAccount -Value $item.SamAccountName
				$CurrentUser | Add-Member -MemberType NoteProperty -Name CorrectSamAccount -Value ""
				$CurrentUser | Add-Member -MemberType NoteProperty -Name EurName -Value $item.Name
				$CurrentUser | Add-Member -MemberType NoteProperty -Name CorrectName -Value ""
				$CurrentUser | Add-Member -MemberType NoteProperty -Name eurUPN -Value $item.userPrincipalName
				$CurrentUser | Add-Member -MemberType NoteProperty -Name ObjectPath -Value ""
				$CurrentUser | Add-Member -MemberType NoteProperty -Name extensionAttribute14 -Value ""
				$CurrentUser | Add-Member -MemberType NoteProperty -Name ObjectGUID -Value $item.ObjectGUID
				$CurrentUser | Add-Member -MemberType NoteProperty -Name SortValue -Value ("NotFound - " + $item.Name)
				$script:AutoItems += $currentUser			
			}
			
			$userListBox.Items.Clear()
			foreach($item in $script:AutoItems){
				$userListBox.Items.Add($item.SortValue)
			}				
			return
		}
			
				
				
		#if validating multiple objects		
		[int]$i = 1
			foreach($item in $userListBox.Items){
				Write-Host "$i  of " $userListBox.Items.Count " users have been validated."
				$foundItem = $DISA | where {$_.ObjectGUID -eq $item.ObjectGUID}
				if($foundItem -ne $null){
					$result = Validate-EURUser ($item)($foundItem)
					
					$foundItem | Add-Member -MemberType NoteProperty -Force -Name SortValue -Value ("Found - $result - " + $item.Name)
					$script:AutoItems += $foundItem
					
				}
				else{
				
					$CurrentUser = New-Object PSObject
					$CurrentUser | Add-Member -MemberType NoteProperty -Name FlagType -Value ""
					$CurrentUser | Add-Member -MemberType NoteProperty -Name EurSamAccount -Value $item.SamAccountName
					$CurrentUser | Add-Member -MemberType NoteProperty -Name CorrectSamAccount -Value ""
					$CurrentUser | Add-Member -MemberType NoteProperty -Name EurName -Value $item.Name
					$CurrentUser | Add-Member -MemberType NoteProperty -Name CorrectName -Value ""
					$CurrentUser | Add-Member -MemberType NoteProperty -Name eurUPN -Value $item.userPrincipalName
					$CurrentUser | Add-Member -MemberType NoteProperty -Name ObjectPath -Value ""
					$CurrentUser | Add-Member -MemberType NoteProperty -Name extensionAttribute14 -Value ""
					$CurrentUser | Add-Member -MemberType NoteProperty -Name ObjectGUID -Value $item.ObjectGUID
					$CurrentUser | Add-Member -MemberType NoteProperty -Name SortValue -Value ("NotFound - " + $item.Name)
					$script:AutoItems += $currentUser
				}
				$i++	
			}
			$script:AutoItems = $script:AutoItems | Sort-Object "SortValue" 
			$userListBox.Items.Clear()
			foreach($item in $script:AutoItems){
				$userListBox.Items.Add($item.SortValue)
			}	
			
			
		}
		else{
			Write-Host "Did not do anything since list is null"
		}
		
		$2StatusBar.Text = ""
		$22StatusBar.Text = "Please select the account you would like to correct from the left hand menu."
		
	})

$goToKbButton.add_MouseClick(
	{
		[System.Diagnostics.Process]::Start("https://5sigcmd.eep.army.mil/sites/E-TNOSC/ecc/AD/KnowledgeBase/Naming%20Conventions.aspx")
	})	
	
	
$objForm.add_FormClosing([System.Windows.Forms.FormClosingEventHandler]{
	Try{  
		"$Date  SHUT  Computer:  $computer   User:  $username  " + $userFriendlyName | Out-File -FilePath $runLog -Append -Encoding ASCII
	}
	Catch{
		#do nothing if it does not have access
	}	
	})
#endregion Form Actions
	
	
#Initial treeview value population
	[array]$subOUs = Get-ChildItem -Path "AD:\$rootDN" | Where-Object {$_.objectClass -eq "organizationalUnit"}  

	foreach ($value in $subOUs)
	{
		$treeView.Nodes.Add($value.name) | Out-Null	
		foreach($node in $treeView.Nodes)
		{		
			if ($node.text -eq $value.name)
			{
				$node.Name = $value.distinguishedName
			}
		}
	}
	
	
DisplayUserAccountInstructions
	
	
$objForm.ShowDialog() | Out-Null