# Printer Choice menu
$Script:printerName = ''; #Set choice Variable to null

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$form = New-Object Windows.Forms.Form
$form.Size = New-Object Drawing.Size @(260,150)
$form.StartPosition = "CenterScreen"
$form.Text = 'AWP Printer Choice:'

$Font = New-Object System.Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Regular)
$Form.Font = $Font

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Size(5,5)
$label.Size = New-Object System.Drawing.Size(240,20)
$label.Text = "Please Choose AWP Printer:"
$Form.Controls.Add($label)

$dButton = New-Object System.Windows.Forms.Button
$dButton.Location = New-Object System.Drawing.Size(20,30)
$dButton.Size = New-Object System.Drawing.Size(100,30)
$dButton.Text = "Location1"
$dButton.Add_Click({$Script:printerName = 'Location1'; $form.Close()})
$Form.Controls.Add($dButton)

$lButton = New-Object System.Windows.Forms.Button
$lButton.Location = New-Object System.Drawing.Size(120,30)
$lButton.Size = New-Object System.Drawing.Size(100,30)
$lButton.Text = "Location2"
$lButton.Add_Click({$Script:printerName = 'Location2'; $form.Close()})
$Form.Controls.Add($lButton)

$iButton = New-Object System.Windows.Forms.Button
$iButton.Location = New-Object System.Drawing.Size(20,60)
$iButton.Size = New-Object System.Drawing.Size(100,30)
$iButton.Text = "Indiantown"
$iButton.Add_Click({$Script:printerName = 'Indiantown'; $form.Close()})
$Form.Controls.Add($iButton)

$fButton = New-Object System.Windows.Forms.Button
$fButton.Location = New-Object System.Drawing.Size(120,60)
$fButton.Size = New-Object System.Drawing.Size(100,30)
$fButton.Text = "Location3"
$fButton.Add_Click({$Script:printerName = 'Location3'; $form.Close()})
$Form.Controls.Add($fButton)

$Form.Topmost = $True
$Form.Add_Shown({$Form.Activate()})
$result = $form.ShowDialog()

$credentials = Get-Credential

$net = New-Object -com WScript.Network
$user = $credentials.UserName
$pass = $credentials.GetNetworkCredential().Password
$net.mapnetworkdrive("Z:", "\\192.168.1.24\$Script:printerName", "true", $user, $pass)