#region Script Settings
#<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
#  <ScriptPackager>
#    <process>powershell.exe</process>
#    <arguments />
#    <extractdir>%TEMP%</extractdir>
#    <files />
#    <usedefaulticon>true</usedefaulticon>
#    <showinsystray>false</showinsystray>
#    <altcreds>false</altcreds>
#    <efs>true</efs>
#    <ntfs>true</ntfs>
#    <local>false</local>
#    <abortonfail>true</abortonfail>
#    <product />
#    <version>1.0.0.1</version>
#    <versionstring />
#    <comments />
#    <company />
#    <includeinterpreter>false</includeinterpreter>
#    <forcecomregistration>false</forcecomregistration>
#    <consolemode>false</consolemode>
#    <EnableChangelog>false</EnableChangelog>
#    <AutoBackup>false</AutoBackup>
#    <snapinforce>false</snapinforce>
#    <snapinshowprogress>false</snapinshowprogress>
#    <snapinautoadd>2</snapinautoadd>
#    <snapinpermanentpath />
#    <cpumode>1</cpumode>
#    <hidepsconsole>false</hidepsconsole>
#  </ScriptPackager>
#</ScriptSettings>
#endregion

# This Imports the O365 functions
Import-Module O365Functions
Import-Module MSOnline


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")  

$CredForm = New-Object System.WIndows.Forms.Form
$CredForm.Size = New-Object System.Drawing.Size(430,185)
$CredForm.Text = "Enter O365 Username and Password"
$ARCIco = New-Object System.Drawing.Icon("C:\Users\GarrettP.NA\Documents\ARC.ico")
$CredForm.Icon = $ARCIco


$SMBForm = New-Object System.Windows.Forms.Form
$SMBForm.Size = New-Object System.Drawing.Size(175,140)
$SMBForm.Text = "Shared Mailbox Permissions Tool"
$SMBForm.Icon = $ARCIco

$MailboxLicenseForm = New-Object System.Windows.Forms.Form
$MailboxLicenseForm.Size = New-Object System.Drawing.Size(385,170)
$MailboxLicenseForm.Text = "O365 Mailbox License Tool"
$MailboxLicenseForm.Icon = $ARCIco

############################################## Start FileDialog

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $env:USERPROFILE
$OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv"

############################################## End FileDialog

############################################## Start GroupBox 1

$SMBActionBox = New-Object System.Windows.Forms.GroupBox
$SMBActionBox.Location = New-Object System.Drawing.Size(5,5)
$SMBActionBox.size = New-Object System.Drawing.Size(110,100)
$SMBActionBox.text = "SMB Action:" 
$SMBForm.Controls.Add($SMBActionBox)
$SMBActionBox.Visible = $False

############################################## end groupBox 1

############################################## Start GroupBox 2

$LevelAccessBox = New-Object System.Windows.Forms.GroupBox
$LevelAccessBox.Location = New-Object System.Drawing.Size(5,5)
$LevelAccessBox.size = New-Object System.Drawing.Size(120,65)  
$LevelAccessBox.text = "Level of Access:" 
$SMBForm.Controls.Add($LevelAccessBox)
$LevelAccessBox.Visible = $False
############################################## end groupBox 2

############################################## Start GroupBox 3

$ActionBox = New-Object System.Windows.Forms.GroupBox
$ActionBox.Location = New-Object System.Drawing.Size(5,5)
$ActionBox.size = New-Object System.Drawing.Size(150,85)
$ActionBox.text = "Actions:" 
$SMBForm.Controls.Add($ActionBox)
$ActionBox.Visible = $True

############################################## end groupBox 3

############################################## Start GroupBox 4

$CredBox = New-Object System.Windows.Forms.GroupBox
$CredBox.Size = New-Object System.Drawing.Size(255,145)
$CredBox.text = "Enter Your Credentials :" 
$CredForm.Controls.Add($CredBox)

############################################## end groupBox 4

############################################## Start GroupBox for MailboxLicenseForm

$MailboxLicenseBox = New-Object System.Windows.Forms.GroupBox
$MailboxLicenseBox.Location = New-Object System.Drawing.Size(5,5)
$MailboxLicenseBox.Size = New-Object System.Drawing.Size(200, 115)
$MailboxLicenseForm.Controls.Add($MailboxLicenseBox)


############################################## End GroupBox for Form

$MailboxOutPutBox = New-Object System.Windows.Forms.GroupBox
$MailboxOutPutBox.Location = New-Object System.Drawing.Size(210,5)
$MailboxOutPutBox.Size = New-Object System.Drawing.Size(150,115)
$MailboxLicenseForm.Controls.Add($MailboxOutPutBox)

##############################################

$ConnectionboxOutPutBox = New-Object System.Windows.Forms.GroupBox
$ConnectionboxOutPutBox.Location = New-Object System.Drawing.Size(262,5)
$ConnectionboxOutPutBox.Size = New-Object System.Drawing.Size(150,140)
$CredForm.Controls.Add($ConnectionboxOutPutBox)
##############################################

$LevelAccessOutputBox = New-Object System.Windows.Forms.GroupBox
$LevelAccessOutputBox.Location = New-Object System.Drawing.Size(130,5)
$LevelAccessOutputBox.Size = New-Object System.Drawing.Size(150,180) 
$SMBForm.Controls.Add($LevelAccessOutputBox)
$LevelAccessOutputBox.Visible = $False

############################################## Start GroupBox for MobileDeviceStats

$mobileDeviceBox = New-Object System.Windows.Forms.GroupBox
$mobileDeviceBox.Size = New-Object System.Drawing.Size(400, 400)
$SMBForm.Controls.Add($mobileDeviceBox)
$mobileDeviceBox.Visible = $False

############################################# End

############################################# Start 
$MailBoxCreationBox = New-Object System.Windows.Forms.GroupBox
$MailBoxCreationBox.Location = New-Object System.Drawing.Size(5,5)
$MailBoxCreationBox.Size = New-Object System.Drawing.Size(200,170)
$MailBoxCreationBox.Text = "Shared Mailbox Creation"
$SMBForm.Controls.Add($MailBoxCreationBox)
$MailBoxCreationBox.Visible = $false

##############################################

$MailBoxCreationOutputBox = New-Object System.Windows.Forms.GroupBox
$MailBoxCreationOutputBox.Location = New-Object System.Drawing.Size(210,5)
$MailBoxCreationOutputBox.Size = New-Object System.Drawing.Size(150,170)
$SMBForm.Controls.Add($MailBoxCreationOutputBox)
$MailBoxCreationOutputBox.Visible = $false

############################################## Start Label 1
$UserNameLabel = New-Object System.WIndows.Forms.Label
$UserNameLabel.Location = New-Object System.Drawing.Size(5,20)
$UserNameLabel.Size = New-Object System.Drawing.Size(125,15)
$UserNameLabel.Text = "O365 Admin Username:"
$CredBox.Controls.Add($UserNameLabel)
############################################## End Label 1

############################################## Start Label 2
$UserPassLabel = New-Object System.Windows.Forms.Label
$UserPassLabel.Location = New-Object System.Drawing.Size(5,40)
$UserPassLabel.Size = New-Object System.Drawing.Size(125,15)
$UserPassLabel.Text = "O365 Admin Password:"
$CredBox.Controls.Add($UserPassLabel)
############################################## End Label 2

############################################## Start Label for MailboxLicenseBox

$MailboxLicenseLabel = New-Object System.Windows.Forms.Label
$MailboxLicenseLabel.Location = New-Object System.Drawing.Size(5, 20)
$MailboxLicenseLabel.Size = New-Object System.Drawing.Size(115, 15)
$MailboxLicenseLabel.Text = "Enter E-mail to check:"
$MailboxLicenseBox.Controls.Add($MailboxLicenseLabel)
############################################# End label for Mailbox License

############################################# Start Label for MobileDevice

$mobileDeviceLabel = New-Object System.Windows.Forms.Label
$mobileDeviceLabel.Location = New-Object System.Drawing.Size(5, 20)
$mobileDeviceLabel.Size = New-Object System.Drawing.Size(125, 15)
$mobileDeviceLabel.Text = "Enter E-mail to check Mobile Devices"
$mobileDeviceBox.Controls.Add($mobileDeviceLabel)


############################################# End Mobile Device

############################################ Label's for MailboxCreationBox

$mailboxcreationlastname = New-Object System.Windows.Forms.Label
$mailboxcreationlastname.Location = New-Object System.Drawing.Size(5,20)
$mailboxcreationlastname.Size = New-Object System.Drawing.Size(60,15)
$mailboxcreationlastname.Text = "LastName"
$MailBoxCreationBox.Controls.Add($mailboxcreationlastname)

$mailboxcreationLogonName = New-Object System.Windows.Forms.Label
$mailboxcreationLogonName.Location = New-Object System.Drawing.Size(5,50)
$mailboxcreationLogonName.Size = New-Object System.Drawing.Size(60,15)
$mailboxcreationLogonName.Text = "LogonName"
$MailBoxCreationBox.Controls.Add($mailboxcreationLogonName)

$mailboxcreationDomainName = New-Object System.Windows.Forms.Label
$mailboxcreationDomainName.Location = New-Object System.Drawing.Size(5,80)
$mailboxcreationDomainName.Size = New-Object System.Drawing.Size(75,15)
$mailboxcreationDomainName.Text = "DomainName"
$MailBoxCreationBox.Controls.Add($mailboxcreationDomainName)

$mailboxcreationManagerName = New-Object System.Windows.Forms.Label
$mailboxcreationManagerName.Location = New-Object System.Drawing.Size(5,110)
$mailboxcreationManagerName.Size = New-Object System.Drawing.Size(50,15)
$mailboxcreationManagerName.Text = "Manager"
$MailBoxCreationBox.Controls.Add($mailboxcreationManagerName)


############################################## Start TextBox 1
$UserNameTextBox = New-Object System.Windows.Forms.TextBox
$UserNameTextBox.Location = New-Object System.Drawing.Size(130,17)
$UserNameTextBox.Size = New-Object System.Drawing.Size(120,20)
$CredBox.Controls.Add($UserNameTextBox)
############################################## End TextBox 1

############################################## Start MaskedTextBox 1
$UserPassMaskedTextBox = New-Object System.Windows.Forms.MaskedTextBox
$UserPassMaskedTextBox.PasswordChar = '*'
$UserPassMaskedTextBox.Location = New-Object System.Drawing.Size(130,40)
$UserPassMaskedTextBox.Size = New-Object System.Drawing.Size(120,20)
$CredBox.Controls.Add($UserPassMaskedTextBox)
############################################## End MaskedTextBox 1

############################################## Start TextBox for Mailbox License Check
$MailboxLicenseTextBox = New-Object System.Windows.Forms.TextBox
$MailboxLicenseTextBox.Location = New-Object System.Drawing.Size(120, 17)
$MailboxLicenseTextBox.Size = New-Object System.Drawing.Size(75,20)
$MailboxLicenseBox.Controls.Add($MailboxLicenseTextBox)

############################################## End TextBox for Mailbox License Check

############################################## Start TextBox for Mobile Devices
$mobileDeviceTextBox = New-Object System.Windows.Forms.TextBox
$mobileDeviceTextBox.Location = New-Object System.Drawing.Size(130, 17)
$mobileDeviceTextBox.Size = New-Object System.Drawing.Size(120, 20)
$mobileDeviceBox.Controls.Add($mobileDeviceTextBox)
############################################## End TextBox for Mobile Devices

############################################# TextBoxes for MailboxCreationBox

$mailboxCreationLastNameTextBox = New-Object System.Windows.Forms.TextBox
$mailboxCreationLastNameTextBox.Location = New-Object System.Drawing.Size(65,18)
$mailboxCreationLastNameTextBox.Size = New-Object System.Drawing.Size(125,15)
$MailBoxCreationBox.Controls.Add($mailboxCreationLastNameTextBox)

$mailboxCreationLogonNameTextBox = New-Object System.Windows.Forms.TextBox
$mailboxCreationLogonNameTextBox.Location = New-Object System.Drawing.Size(65,47)
$mailboxCreationLogonNameTextBox.Size = New-Object System.Drawing.Size(125,15)
$MailBoxCreationBox.Controls.Add($mailboxCreationLogonNameTextBox)

$mailboxCreationDomainNameTextBox = New-Object System.Windows.Forms.TextBox
$mailboxCreationDomainNameTextBox.Location = New-Object System.Drawing.Size(80,77)
$mailboxCreationDomainNameTextBox.Size = New-Object System.Drawing.Size(110,15)
$MailBoxCreationBox.Controls.Add($mailboxCreationDomainNameTextBox)

$mailboxCreationManagerNameTextBox = New-Object System.Windows.Forms.TextBox
$mailboxCreationManagerNameTextBox.Location = New-Object System.Drawing.Size(55,107)
$mailboxCreationManagerNameTextBox.Size = New-Object System.Drawing.Size(135,15)
$MailBoxCreationBox.Controls.Add($mailboxCreationManagerNameTextBox)


############################################

$MailboxOutputInformation = New-Object System.Windows.Forms.TextBox
$MailboxOutputInformation.Location = New-Object System.Drawing.Size(0, 5)
$MailboxOutputInformation.Size = New-Object System.Drawing.Size(150,110)
$MailboxOutputInformation.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$MailboxOutputInformation.Multiline = $True
$MailboxOutputInformation.ReadOnly = $True
$MailboxOutPutBox.Controls.Add($MailboxOutputInformation)
############################################################################

$ConnectionOutputInformation = New-Object System.Windows.Forms.TextBox
$ConnectionOutputInformation.Location = New-Object System.Drawing.Size(0, 0)
$ConnectionOutputInformation.Size = New-Object System.Drawing.Size(150, 140)
$ConnectionOutputInformation.Multiline = $True
$ConnectionOutputInformation.ReadOnly = $True
$ConnectionboxOutPutBox.Controls.Add($ConnectionOutputInformation)
############################################################################
$LevelAccessOutputInformation = New-Object System.Windows.Forms.TextBox
$LevelAccessOutputInformation.Location = New-Object System.Drawing.Size(0,0)
$LevelAccessOutputInformation.Size = New-Object System.Drawing.Size(150,180)
$LevelAccessOutputInformation.Multiline = $True
$LevelAccessOutputInformation.ReadOnly = $True
$LevelAccessOutputBox.Controls.Add($LevelAccessOutputInformation)
$LevelAccessOutputInformation.Visible = $false
#######################################################

$MailboxCreationOutputInformation = New-Object System.Windows.Forms.TextBox
$MailboxCreationOutputInformation.Location = New-Object System.Drawing.Size(0,0)
$MailboxCreationOutputInformation.Size = New-Object System.Drawing.Size(150,170)
$MailboxCreationOutputInformation.Multiline = $True
$MailboxCreationOutputInformation.ReadOnly = $True
$MailboxCreationOutputBox.Controls.Add($MailboxCreationOutputInformation)
$MailboxCreationOutputInformation.Visible = $false


############################################## Start radio buttons for Group Box 1
$SMBActionRadioButton1 = New-Object System.Windows.Forms.RadioButton 
$SMBActionRadioButton1.Checked = $true
$SMBActionRadioButton1.Location = New-Object System.Drawing.Point(5,15)
$SMBActionRadioButton1.size = New-Object System.Drawing.Size(85,14)
$SMBActionRadioButton1.Text = "Add Access" 
$SMBActionBox.Controls.Add($SMBActionRadioButton1) 

$SMBActionRadioButton2 = New-Object System.Windows.Forms.RadioButton
$SMBActionRadioButton2.Location = New-Object System.Drawing.Point(5,30)
$SMBActionRadioButton2.size = New-Object System.Drawing.Size(104,14)
$SMBActionRadioButton2.Text = "Remove Access"
$SMBActionBox.Controls.Add($SMBActionRadioButton2)

$SMBActionRadioButton3 = New-Object System.Windows.Forms.RadioButton
$SMBActionRadioButton3.Location = New-Object System.Drawing.Point(5,45)
$SMBActionRadioButton3.Size = New-Object System.Drawing.Size(104,14)
$SMBActionRadioButton3.Text = "Create Mailbox"
$SMBActionBox.Controls.Add($SMBActionRadioButton3)

############################################## end radio buttons for Group Box 1

############################################## Start radio buttons for Group Box 2
$LevelAccessRadioButton1 = New-Object System.Windows.Forms.RadioButton 
$LevelAccessRadioButton1.Checked = $true
$LevelAccessRadioButton1.Location = New-Object System.Drawing.Point(5,15)
$LevelAccessRadioButton1.size = New-Object System.Drawing.Size(85,14) 
$LevelAccessRadioButton1.Text = "SendAs" 
$LevelAccessBox.Controls.Add($LevelAccessRadioButton1) 

$LevelAccessRadioButton2 = New-Object System.Windows.Forms.RadioButton
$LevelAccessRadioButton2.Location = New-Object System.Drawing.Point(5,30)
$LevelAccessRadioButton2.size = New-Object System.Drawing.Size(114,14)
$LevelAccessRadioButton2.Text = "Send on Behalf Of"
$LevelAccessBox.Controls.Add($LevelAccessRadioButton2)

$LevelAccessRadioButton3 = New-Object System.Windows.Forms.RadioButton
$LevelAccessRadioButton3.Location = New-Object System.Drawing.Point(5,45)
$LevelAccessRadioButton3.size = New-Object System.Drawing.Size(90,14) 
$LevelAccessRadioButton3.Text = "No SendAs" 
$LevelAccessBox.Controls.Add($LevelAccessRadioButton3) 

############################################## end radio buttons for Group Box 2

############################################## Start radio buttons for Group Box 3
$ActionRadioButton1 = New-Object System.Windows.Forms.RadioButton 
$ActionRadioButton1.Checked = $true
$ActionRadioButton1.Location = New-Object System.Drawing.Point(5,15)
$ActionRadioButton1.size = New-Object System.Drawing.Size(120,14)
$ActionRadioButton1.Text = "Shared Mailbox" 
$ActionBox.Controls.Add($ActionRadioButton1) 

$ActionRadioButton2 = New-Object System.Windows.Forms.RadioButton
$ActionRadioButton2.Location = New-Object System.Drawing.Point(5,30)
$ActionRadioButton2.size = New-Object System.Drawing.Size(142,14)
$ActionRadioButton2.Text = "Mobile Device Statistics"
$ActionBox.Controls.Add($ActionRadioButton2)

############################################## end radio buttons for Group Box 3

############################################## Start radio buttons for CredForm CredBox
$EmailLicenseButton = New-Object System.Windows.Forms.RadioButton 
$EmailLicenseButton.Location = New-Object System.Drawing.Point(64,65)
$EmailLicenseButton.size = New-Object System.Drawing.Size(137,14)
$EmailLicenseButton.Text = "Email Account License" 
$CredBox.Controls.Add($EmailLicenseButton)
$EmailLicenseButton.Checked = $True

###############################################

$SharedMailboxButton = New-Object System.Windows.Forms.RadioButton
$SharedMailboxButton.Location = New-Object System.Drawing.Point(64,80)
$SharedMailboxButton.size = New-Object System.Drawing.Size(150,14)
$SharedMailboxButton.Text = "Shared Mailbox Actions"
$CredBox.Controls.Add($SharedMailboxButton)

############################################## End radio buttons for CredForm CredBox


############################################## Start Functions

############################################## Start OpenFileDialog
function GetFile{
$OpenFileDialog.ShowDiaLog() | Out-Null
$global:CSVFile = $OpenFileDialog.filename

############################################## End OpenFileDialog
}

############################################## Start PermissionApply
function PermissionApply{

if($SMBActionRadioButton1.Checked -eq $true){
	
	if($LevelAccessRadioButton1.Checked -eq $true) {
			$LevelAccessOutputInformation.Text = AddAccessSendAs($CSVFile)
			if($LevelAccessOutputInformation.Text -match "SendAs")
				{
				$LevelAccessOutputInformation.BackColor = "DarkGreen"
				$LevelAccessOutputInformation.ForeColor = "White"
				}
	}
	elseif($LevelAccessRadioButton2.Checked -eq $true) {
			$LevelAccessOutputInformation.Text = AddAccessonBehalfOf($CSVFile)
			if($LevelAccessOutputInformation.Text -match "SendOnBehalfOf")
				{
				$LevelAccessOutputInformation.BackColor = "DarkGreen"
				$LevelAccessOutputInformation.ForeColor = "White"
				}
	}
	elseif($LevelAccessRadioButton2.Checked -eq $true) {
			$LevelAccessOutputInformation.Text = AddAccessNoSendAs($CSVFile)
			if($LevelAccessOutputInformation.Text -match "NoSendAs")
				{
				$LevelAccessOutputInformation.BackColor = "DarkGreen"
				$LevelAccessOutputInformation.ForeColor = "White"
				}
	}
}

elseif($SMBActionRadioButton2.Checked -eq $true) {
	$SMBActionBox.Visible = False
	if($LevelAccessRadioButton1.Checked -eq $true) {
			$LevelAccessOutputInformation.Text = RemoveAccessSendAs($CSVFile)
			if($LevelAccessOutputInformation.Text -match "SendAs")
				{
				$LevelAccessOutputInformation.BackColor = "DarkRed"
				$LevelAccessOutputInformation.ForeColor = "White"
				}
	}
	elseif($LevelAccessRadioButton2.Checked -eq $true) {
			$LevelAccessOutputInformation.Text = RemoveAccessonBehalfOf($CSVFile)
			if($LevelAccessOutputInformation.Text -match "SendOnBehalfOf")
				{
				$LevelAccessOutputInformation.BackColor = "DarkRed"
				$LevelAccessOutputInformation.ForeColor = "White"
				}
	}
	elseif($LevelAccessRadioButton3.Checked -eq $true) {
			$LevelAccessOutputInformation.Text = RemoveAccessNoSendAs($CSVFile)
			if($LevelAccessOutputInformation.Text -match "NoSendAs")
				{
				$LevelAccessOutputInformation.BackColor = "DarkRed"
				$LevelAccessOutputInformation.ForeColor = "White"
				}
	}
}



}

############################################# Start ActionSelect

function ActionSelect{
	if ($ActionRadioButton1.Checked -eq $True)
	{
	$ActionBox.Visible = $False
	$ActionButton.Visible = $False
	$SMBActionBox.Visible = $True
	$SMBActionButton.Visible = $True
	$SMBForm.Size = New-Object System.Drawing.Size(140,150)	
		
		
		
	}
	elseif ($ActionRadioButton2.Checked -eq $True)
	{
		$ActionBox.Visible = $false
		$ActionButton.Visible = $False
		$mobileDeviceBox.Visible = $True
	}
}

############################################## End ActionSelect


############################################## End PermissionApply

############################################## Start EmailActionSelect
function SMBActionSelect {
	if($SMBActionRadioButton1.Checked -eq $true){
		$ActionBox.Visible = $False
		$ActionButton.Visible = $False
		$SMBActionBox.Visible = $False
		$LevelAccessBox.Visible = $True
		$SMBActionButton.Visible = $False
		$GetCSVButton.Visible = $True
		$ApplyPermissionsButton.Visible = $True
		$LevelAccessOutputBox.Visible = $True
		$LevelAccessOutputInformation.Visible = $True
		$SMBForm.Size = New-Object System.Drawing.Size(300,225)
	}
	elseif($SMBActionRadioButton2.Checked -eq $true)
	{
		$ActionBox.Visible = $False
		$ActionButton.Visible = $False
		$SMBActionBox.Visible = $False
		$LevelAccessBox.Visible = $True
		$SMBActionButton.Visible = $False
		$GetCSVButton.Visible = $True
		$ApplyPermissionsButton.Visible = $True
		$LevelAccessOutputBox.Visible = $True
		$LevelAccessOutputInformation.Visible = $True
		$SMBForm.Size = New-Object System.Drawing.Size(300,225)

	}
			
	elseif($SMBActionRadioButton3.Checked -eq $true){
		$ActionBox.Visible = $False
		$SMBActionBox.Visible = $False
		$SMBActionButton.Visible = $False
		$MailBoxCreationBox.Visible = $True
		$MailBoxCreationOutputBox.Visible = $True
		$MailboxCreationOutputInformation.Visible = $true
		$SMBForm.Size = New-Object System.Drawing.Size(385,220)
	}

}
############################################## End EmailActionSelect

############################################## Start Connect to O365 and Close CredForm

function ConnectToO365
	{
	############################################## Convert PW to SecureString and make PSCredential Variable
	
	
	if($SharedMailboxButton.Checked -eq $True)
		{
		$ConnectionOutputInformation.Text = "Starting Connection to O365"
		$O365PW = $UserPassMaskedTextBox.Text | ConvertTo-SecureString -AsPlainText -Force
		$FullO365Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserNameTextBox.Text,$O365PW
		$ConnectionOutputInformation.Text = O365MailboxConnect($FullO365Creds)
		if($ConnectionOutputInformation.Text -match "Successfully connected to O365")
		{
		$ConnectionOutputInformation.ForeColor = "White"
		$connectionoutputinformation.BackColor = "DarkGreen"
		#$credForm.Close()
		}
	elseif($ConnectionOutputInformation.Text -match "ConnectionCheck not established, please enter credentials")
	{
		$ConnectionOutputInformation.ForeColor = "White"
		$ConnectionOutputInformation.BackColor = "DarkRed"
		$ConnectionOutputInformation.Text = O365MailboxConnect($FullO365Creds)
	}
		
		
		}
	elseif($EmailLicenseButton.Checked -eq $True)
		{
		$ConnectionOutputInformation.Text = "Starting Connection to O365"
		
		$O365PW = $UserPassMaskedTextBox.Text | ConvertTo-SecureString -AsPlainText -Force
		$FullO365Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserNameTextBox.Text,$O365PW
		
		$ConnectionOutputInformation.Text = O365MailboxConnect($FullO365Creds)
		if($ConnectionOutputInformation.Text -match "Successfully connected to O365")
			{
			$ConnectionOutputInformation.ForeColor = "White"
			$ConnectionOutputInformation.BackColor = "DarkGreen"
			}
		
		}
	}

############################################## End Connect to O365 and Close CredForm

##############################################  Start CheckLicense

function CheckLicense
{
	$checkUserLicense = $MailboxLicenseTextBox.Text
	$MailboxOutputInformation.Text = O365LicenseCheck($checkUserLicense)
	if($MailboxOutputInformation.Text -match "User Account Does not Exist please try again later" )
		{
		$MailboxOutputInformation.ForeColor = "White"
		$MailboxOutputInformation.BackColor = "DarkRed"
		}
	elseif($MailboxOutputInformation.Text -match "User Account Found")
		{
		$MailboxOutputInformation.ForeColor = "White"
		$mailboxoutputinformation.BackColor = "DarkGreen"
		if($MailboxOutputInformation.Text -match "Account not licensed")
			{
			$MailboxOutputInformation.ForeColor = "White"
			$MailboxOutputInformation.BackColor = "DarkRed"
			}
	}
	
}

############################################## End CheckLicense



############################################# Start GetMobileDeviceStats

function getMobileDevices
{
	$theUser = $mobileDeviceTextBox.Text
	GetMobileDeviceStats($theUser)
	
	
	
}

############################################ End GetMobileDeviceStats

########################################### createMailbox

function createMailbox
{
	$mailboxCreateLastName = ($mailboxCreationLastNameTextBox.Text).Trim()
	$mailboxCreateLogonName = ($mailboxCreationLogonNameTextBox.Text).Trim()
	$mailboxCreateDomainName = ($mailboxCreationDomainNameTextBox.Text).Trim()
	$mailboxCreateManagerName = ($mailboxCreationManagerNameTextBox.Text).Trim()
	$mailboxCreationArray = @($mailboxCreateLastName,$mailboxCreateLogonName,$mailboxCreateDomainName,$mailboxCreateManagerName,$CSVFile)
	$MailboxCreationOutputInformation.Text = mailboxCreate([array]$mailboxCreationArray)
	
}

function addtoDelegateGroup
	{
	$mailboxDelegateLastName = ($mailboxCreationLastNameTextBox.Text).Trim()
	$mailboxDelegateLogonName = ($mailboxCreationLogonNameTextBox.Text).Trim()
	$mailboxDelegateDomain = ($mailboxCreationDomainNameTextBox.Text).Trim()
	$mailboxDelegateManager = ($mailboxCreationManagerNameTextBox.Text).Trim()
	$mailboxDelegateArray = @($mailboxDelegateLastName,$mailboxDelegateLogonName,$mailboxDelegateDomain,$mailboxDelegateManager,$CSVFile)
	#$MailboxCreationOutputInformation.Text =  addUserstoDelegates([array]$mailboxDelegateArray)
	}



############################################## End Functions


############################################## Start Button 1

$GetCSVButton = New-Object System.Windows.Forms.Button 
$GetCSVButton.Location = New-Object System.Drawing.Size(25,72) 
$GetCSVButton.Size = New-Object System.Drawing.Size(78,40) 
$GetCSVButton.Text = "Get CSV File" 
$GetCSVButton.Add_Click({GetFile}) 
$SMBForm.Controls.Add($GetCSVButton)
$GetCSVButton.Visible = $False

############################################## End Button 1

############################################## Start Button 2

$ApplyPermissionsButton = New-Object System.Windows.Forms.Button 
$ApplyPermissionsButton.Location = New-Object System.Drawing.Size(25,120) 
$ApplyPermissionsButton.Size = New-Object System.Drawing.Size(78,50) 
$ApplyPermissionsButton.Text = "Apply Permissions" 
$ApplyPermissionsButton.Add_Click({PermissionApply}) 
$SMBForm.Controls.Add($ApplyPermissionsButton)
$ApplyPermissionsButton.Visible = $False

############################################## End Button 2

############################################## Start Button 3

$ActionButton = New-Object System.Windows.Forms.Button 
$ActionButton.Location = New-Object System.Drawing.Size(40,50) 
$ActionButton.Size = New-Object System.Drawing.Size(55,20) 
$ActionButton.Text = "Select" 
$ActionButton.Add_Click({ActionSelect}) 
$ActionBox.Controls.Add($ActionButton)


############################################## End Button 3

############################################## Start Button 4

$SMBActionButton = New-Object System.Windows.Forms.Button 
$SMBActionButton.Location = New-Object System.Drawing.Size(25,70) 
$SMBActionButton.Size = New-Object System.Drawing.Size(60,20) 
$SMBActionButton.Text = "Select" 
$SMBActionButton.Add_Click({SMBActionSelect}) 
$SMBActionBox.Controls.Add($SMBActionButton)
$SMBActionButton.Visible = $False

############################################## End Button 4

############################################## Start Button 5

$O365ConnectButton = New-Object System.Windows.Forms.Button
$O365ConnectButton.Location = New-Object System.Drawing.Size(80,100)
$O365ConnectButton.Size = New-Object System.Drawing.Size(55,35)
$O365ConnectButton.Text = "Connect"
$O365ConnectButton.Add_Click({ConnectToO365})
$credBox.Controls.Add($O365ConnectButton)
############################################## End Button 5

############################################## Start button for MailboxLicenseBox

$MailboxLicenseButton = New-Object System.Windows.Forms.Button
$MailboxLicenseBUtton.Location = New-Object System.Drawing.Size(75, 50)
$MailboxLicenseButton.Size = New-Object System.Drawing.Size(65, 50)
$MailboxLicenseButton.Text = "Check Account"
$MailboxLicenseButton.Add_Click({ checkLicense })
$MailboxLicenseBox.Controls.Add($MailboxLicenseButton)

############################################### End Button for mailbox License

############################################### Start Button for MobileDevice

$MobileDeviceButton = New-Object System.Windows.Forms.Button
$MobileDeviceButton.Location = New-Object System.Drawing.Size(5, 70)
$MobileDeviceButton.Size = New-Object System.Drawing.Size(78, 50)
$MobileDeviceButton.Text = "Check Account"
$MobileDeviceButton.Add_Click({ getMobileDevices })
$mobileDeviceBox.Controls.Add($MobileDeviceButton)

############################################## End Mobile Device Button

############################################# MailboxCreationButton

$MailboxCreationButton = New-Object System.Windows.Forms.Button
$MailboxCreationButton.Location = New-Object System.Drawing.Size(18,130)
$MailboxCreationButton.Size = New-Object System.Drawing.Size(50,35)
$MailboxCreationButton.Text = "Create"
$MailboxCreationButton.Add_Click({ createMailbox })
$MailBoxCreationBox.Controls.Add($MailboxCreationButton)

$mailboxImportUsers = New-Object System.Windows.Forms.Button
$mailboxImportUsers.Location = New-Object System.Drawing.Size(70,130)
$mailboxImportUsers.Size = New-Object System.Drawing.Size(50,35)
$mailboxImportUsers.Text = "Import Users"
$mailboxImportUsers.Add_Click({ GetFile })
$MailBoxCreationBox.Controls.Add($mailboxImportUsers)

$MailboxSecurityGroupButton = New-Object System.Windows.Forms.Button
$MailboxSecurityGroupButton.Location = New-Object System.Drawing.Size(122,130)
$MailboxSecurityGroupButton.Size = New-Object System.Drawing.Size(65,35)
$MailboxSecurityGroupButton.Text = "Add to Delegates"
$MailboxSecurityGroupButton.Add_Click({addtoDelegateGroup})
$MailBoxCreationBox.Controls.Add($MailboxSecurityGroupButton)

############################################## Display Forms
$CredForm.Add_Shown({$CredForm.Activate()})
[void] $CredForm.ShowDialog()

if ($SharedMailboxButton.Checked -eq $true)
{ 
	$SMBForm.Add_Shown({ $SMBForm.Activate() })
	[void]$SMBForm.ShowDialog()
}
elseif ($EmailLicenseButton.Checked -eq $true)
{
	$MailboxLicenseForm.Add_Shown({ $MailboxLicenseForm.Activate() })
	[void]$MailboxLicenseForm.ShowDialog()
}				
		
Remove-Module O365Functions