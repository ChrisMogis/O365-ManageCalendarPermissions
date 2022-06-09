################################################################################################
# This script is used to manage rights on users' calendars on Office 365.                      #
# Editor : Christopher Mogis                                                                   #
# Date : 06/03/2022                                                                            #
# Version 1.1                                                                                  #
################################################################################################

Param(
[Parameter(Mandatory=$true)]
[ValidateSet("Info", "Add", "Remove")]
[String[]]
$Choose
)

#Set Powershell Execution policy
Set-ExecutionPolicy RemoteSigned -Force

#Install PowershellGet Module
Install-Module PowershellGet -Force

#Install ExchangeOnline Management Module
Install-Module -Name ExchangeOnlineManagement 

#Connect to exchange online
Connect-ExchangeOnline

#Targeted User Information
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Question'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$form.Icon = 'C:\Temp\favicon-image.ico'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(105,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the target Username in the space below:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

If ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
    $TargetUser = $textBox.Text
    $TargetUser
    }

If ($Choose -eq "Info")
    {
#List all calendar with all permissions
#Get-Mailbox -Identity $TargetUser | Get-MailboxFolderStatistics -FolderScope Calendar | ft Identity,Name
Get-MailboxFolderPermission -Identity "$($TargetUser):\Calendrier" | ft Identity,FolderName,User,AccessRights
    }

If ($Choose -eq "Add")
    {
#Delegated User Information
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Question'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$form.Icon = 'C:\Temp\favicon-image.ico'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(105,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Enter the name of the user receiving the rights:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

If ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
    $DelegatedUser = $textBox.Text
    $DelegatedUser
    }

#Delegated Right Information
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select Permission'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$form.Icon = 'C:\Temp\favicon-image.ico'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(105,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select the permission :'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

[void] $listBox.Items.Add('Owner')
[void] $listBox.Items.Add('PublishEditor')
[void] $listBox.Items.Add('Editor')
[void] $listBox.Items.Add('PublishingAuthor')
[void] $listBox.Items.Add('Author')
[void] $listBox.Items.Add('NonEditingAuthor')
[void] $listBox.Items.Add('Reviewer')
[void] $listBox.Items.Add('Contributor')
[void] $listBox.Items.Add('LimitedDetails')
[void] $listBox.Items.Add('AvailabilityOnly')

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Permission = $listBox.SelectedItem
    $Permission
}

Add-MailboxFolderPermission -Identity "$($TargetUser):\Calendrier" -User $DelegatedUser -AccessRights $Permission
    }

#Delete Calendar Permission
If ($Choose -eq "Remove")
{
#Delegated User Information
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Remove Permissions'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'
$form.Icon = 'C:\Temp\favicon-image.ico'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(105,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Enter the username to whom we remove the rights'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

If ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
    $RemoveUser = $textBox.Text
    $RemoveUser
    }

Remove-MailboxFolderPermission -Identity "$($TargetUser):\Calendrier" -User $RemoveUser -Confirm:$false
}

#Disconnect PSSession
Disconnect-ExchangeOnline -Confirm:$false
