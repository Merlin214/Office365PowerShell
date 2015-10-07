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

# These functions are associated 

function O365MailboxConnectionCheck
{
	if ((Get-PSSession) -eq $null) 
	{
				Write-Output ConnectionCheck not established, please enter credentials
				Remove-Item -Path.\authInfo.txt
				O365MailboxConnect
	}
	else
	{
				Write-Output Successfully connected to O365
	}
}


function O365MailboxConnect([PSCredential]$O365Creds) {
Write-Output Starting Connection to O365
$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $O365Creds -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue 2> .\ConnectError.txt
$ConnectError = Get-Content -Path .\ConnectError.txt
Remove-Item .\ConnectError.txt
Import-PSSession $O365Session -InformationAction Ignore -WarningAction Ignore -ErrorAction Ignore
O365MailboxConnectionCheck
}


function AddAccessSendAs([string]$CSV){
#Set FullAccess permission
Import-Csv $CSV | ForEach-Object{

Add-MailboxPermission $_."Mailbox" -User $_."User" -AccessRights FullAccess -InheritanceType All -Confirm:$False -InformationAction Ignore -WarningAction Ignore -ErrorAction Ignore


#Set SendAs permission

Add-RecipientPermission $_."Mailbox" -AccessRights SendAs -Trustee $_."User" -Confirm:$False -WarningAction Ignore -ErrorAction Ignore

Write-Output "SendAs Access has been Granted to" $_."User" "for mailbox access to" $_."Mailbox"
}
}


function RemoveAccessSendAs([string]$CSV){

Import-Csv $CSV | ForEach-Object{

#Set FullAccess permission

Remove-MailboxPermission $_."Mailbox" -User $_."User" -AccessRights FullAccess -InheritanceType All -Confirm:$False -InformationAction Ignore -WarningAction Ignore -ErrorAction Ignore


#Set SendAs permission

Remove-RecipientPermission $_."Mailbox" -AccessRights SendAs -Trustee $_."User" -Confirm:$False -WarningAction Ignore -ErrorAction Ignore

Write-Output "SendAs Access has been Removed From" $_."User" "for mailbox access to" $_."Mailbox"
}
}

function AddAccessonBehalfOf([string]$CSV){

Import-Csv $CSV | ForEach-Object{

#Set FullAccess permission

Add-MailboxPermission $_."Mailbox" -User $_."User" -AccessRights FullAccess -InheritanceType All -Confirm:$False -InformationAction Ignore -WarningAction Ignore -ErrorAction Ignore


#Set SendOnBehlafOf permission

set-Mailbox $_."Mailbox" -GrantSendOnBehalfTo $_."User" -Confirm:$False -WarningAction Ignore -ErrorAction Ignore

Write-Output "SendOnBehalfOf Access has been Granted to" $_."User" "for mailbox access to" $_."Mailbox"
}
}

function RemoveAccessonBehalfOf([string]$CSV){

Import-Csv $CSV | ForEach-Object{

#Set FullAccess permission

Add-MailboxPermission $_."Mailbox" -User $_."User" -AccessRights FullAccess -InheritanceType All -Confirm:$False -InformationAction Ignore -WarningAction Ignore -ErrorAction Ignore


#Set SendOnBehlafOf permission

set-Mailbox $_."Mailbox" -GrantSendOnBehalfTo $_."User" -Confirm:$False -WarningAction Ignore -ErrorAction Ignore

Write-Output "SendOnBehalfOf Access has been Removed From" $_."User" "for mailbox access to" $_."Mailbox"
}
}

function AddAccessNoSendAs([string]$CSV){

Import-Csv $CSV | ForEach-Object{

#Set FullAccess permission

Add-MailboxPermission $_."Mailbox" -User $_."User" -AccessRights FullAccess -InheritanceType All -Confirm:$False -InformationAction Ignore -WarningAction Ignore -ErrorAction Ignore

Write-Output "NoSendAs Access has been Granted to" $_."User" "for mailbox access to" $_."Mailbox"
}
}

function RemoveAccessNoSendAs([string]$CSV)
	{
	
	Import-Csv $CSV | ForEach-Object {
		
		#Set FullAccess permission
		
		Remove-MailboxPermission $_. "Mailbox" -User $_. "User" -AccessRights FullAccess -InheritanceType All -Confirm:$False -InformationAction Ignore -WarningAction Ignore -ErrorAction Ignore
		
		Write-Output "No SendAs Access has been Removed From" $_. "User" "for mailbox access to" $_. "Mailbox"
		}
	}


function GetMobileDeviceStats([string]$theUser)
{
	Get-MobileDeviceStatistics -Mailbox $theUser
	
}

function mailboxCreate([array]$creationArray)
	{
	Write-Output $creationArray[2]
	#Creates connection to AD
	Add-PSSnapin Quest.ActiveRoles.ADManagement
	#connect-QADService -service 'rcodc3chi001.archq.ri.redcross.net'
	# Constants to modify multi-valued AD attributes.
	$ADS_PROPERTY_CLEAR = 1
	$ADS_PROPERTY_UPDATE = 2
	$ADS_PROPERTY_APPEND = 3
	$ADS_PROPERTY_DELETE = 4	
	$MgrFound = $false
	$ADForest = [System.DirectoryServices.ActiveDirectory.forest]::getcurrentforest()
	$GC = $ADForest.FindGlobalCatalog()
	$ADSearcher = $GC.GetDirectorySearcher()
	while($MgrFound -eq $false)
		{
		$ADsearcher.filter = "(&(objectCategory=user)(userprincipalname= $creationArray [3] ))"
		$ADSearcher.pagesize = 1000
		$ADSearcher.PropertiesToLoad.Add("distinguishedName")
		$Results = $ADSearcher.FindAll()
		if($Results -eq 1)
			{
			$MgrFound -eq $true
			}
		}
	
	foreach($u in $results)
		{
		$objItem = $u.Properties
		$objItem.distinguishedName
		$MgrDN = $objItem.distinguishedname
		}
	
	$UserCN = 'CN=' + $creationArray[1]
	$UserMail = $creationArray[1] + '@redcross.org'
	$UserTarget = $creationArray[1] + '@service.redcross.org'
	$UserUPN = $creationArray[1] + '@redcross.org'
	$GroupCN = 'CN=SMB ' + $creationArray[0] + ' Delegates'
	$GroupName = 'SMB ' + $creationArray[0] + ' Delegates'
	$GroupSamName = 'SMB' + $creationArray[0]
	$GroupMail = $GroupSamName + '@redcross.org'
	$GroupTarget = $GroupSamName + '@service.redcross.org'
	
	if($creationArray[2] -eq "ARCHQ")
		{
		Write-Output This is ARCHQ
		$objOU = [ADSI] "LDAP://OU=Shared Mailboxes,OU=Corporate,DC=archq,DC=ri,DC=redcross,DC=net"
		$objGrpOU = [ADSI] "LDAP://OU=Security Groups,OU=Groups,OU=Corporate,DC=archq,DC=ri,DC=redcross,DC=net"
		
		#Mailbox Creation
		$objUser = $objOU.Create("user",$UserCN)
		$objUser.SetInfo()
		
		$objUser.PUT("sAMAccountName",$creationArray[1])
		$objUser.PUT("sn",$creationArray[0])
		$objUser.PUT("displayName",$creationArray[1])
		$objUser.PUT("targetAddress", "SMTP:$UserTarget" )
		$objUser.PUT("mail",$UserMail)
		$objUser.PUT("description",'SMB')
		$objUser.PUT("userPrincipalName",$UserUPN)
		$objUser.putex(3,"proxyAddress",@("SMTP:$UserMail","smtp:$UserTarget"))
#		$objUser.proxyAddresses += "SMTP:$UserMail"
#		$objUser.proxyAddresses += "smtp:$UserTarget"
		$objUser.SetInfo()
		
		$objUser.psbase.Invoke("SetPassword","myGr88PwR3dCR0$ $1946 !")
		$objUser.psbase.InvokeSet("AccountDisabled","False")
		$objuser.psbase.CommitChanges()
		
		$flag = $objUser.userAccountControl.value -bxor 65536
		$objUser.userAccountControl = $flag
		$objUser.SetInfo()
		
		$objUser.PUT("manager",$MgrDN)
		$objUser.SetInfo()
	
		$objGroup = $objGrpOU.Create("group",$GroupCN)
		$objGroup.Put("sAMAccountName",$GroupSamName)
		$objGroup.SetInfo()

		$objGroup.PUT("mail",$GroupMail)
		$objGroup.PUT("description",$GroupName)
		$objGroup.putex(3,"proxyAddress",@("SMTP:$GroupMail","smtp:$GroupTarget"))

#		$objGroup.proxyAddresses += "SMTP: $GroupMail
#		$objGroup.proxyAddresses += "smtp: $GroupTarget
		$objGroup.PUT("GroupType","2147483656")
		$objGroup.SetInfo()

		$objGroup.PUT("managedBy",$MgrDN)
		$objGroup.SetInfo()
		Write-Output The User Account $creationArray[1] was created.
		Write-Output The group $GroupName was created.
		
		}
	elseif($creationArray[2] -eq "BIO")
		{
		Write-Output THis is BIO
		$objOU = [ADSI] "LDAP://OU=Shared Mailboxes,OU=Biomedical Field,DC=bio,DC=ri,DC=redcross,DC=net"
		$objGrpOU = [ADSI] "LDAP://OU=Security Groups,OU=Groups,OU=Biomedical Field,DC=bio,DC=ri,DC=redcross,DC=net"
			#Mailbox Creation
		$objUser = $objOU.Create("user",$UserCN)
		$objUser.SetInfo()
		
		$objUser.PUT("sAMAccountName",$Logonname)
		$objUser.PUT("sn",$lastname)
		$objUser.PUT("displayName",$lastname)
		$objUser.PUT( "targetAddress","SMTP:$UserTarget")
		$objUser.PUT("mail",$UserMail)
		$objUser.PUT("description",'SMB')
		$objUser.PUT("userPrincipalName",$UserUPN)
		$objUser.putex(3,"proxyAddress",@("SMTP:$UserMail","smtp:$UserTarget"))
#		$objUser.proxyAddresses += "SMTP:$UserMail
#		$objUser.proxyAddresses += "smtp:$UserTarget
		$objUser.SetInfo()
		
		$objUser.psbase.Invoke("SetPassword","myGr88PwR3dCR0$ $1946 !")
		$objUser.psbase.InvokeSet("AccountDisabled","False")
		$objuser.psbase.CommitChanges()
		
		$flag = $objUser.userAccountControl.value -bxor 65536
		$objUser.userAccountControl = $flag
		$objUser.SetInfo()
		
		$objUser.PUT("manager",$MgrDN)
		$objUser.SetInfo()
		
		$objGroup = $objGrpOU.Create("group",$GroupCN)
		$objGroup.Put("sAMAccountName",$GroupSamName)
		$objGroup.SetInfo()

		$objGroup.PUT("mail",$GroupMail)
		$objGroup.PUT("description",$GroupName)
		$objGroup.putex(3,"proxyAddress",@("SMTP:$GroupMail","smtp:$GroupTarget"))
#		$objGroup.proxyAddresses += "SMTP: $GroupMail
#		$objGroup.proxyAddresses += "smtp: $GroupTarget
		$objGroup.PUT("GroupType","2147483656")
		$objGroup.SetInfo()

		$objGroup.PUT("managedBy",$MgrDN)
		$objGroup.SetInfo()
		
		Write-Output The User Account $creationArray[1] was created.
		Write-Output The group $GroupName was created.
		}
	elseif($creationArray[2] -eq "RC")
		{
		Write-Output THis is RC
		$objOU = [ADSI] "LDAP://OU=Shared Mailboxes,OU=Chapters,DC=rc,DC=ri,DC=redcross,DC=net"
		$objGrpOU = [ADSI] "LDAP://OU=Security Groups,OU=Groups,OU=Chapters,DC=rc,DC=ri,DC=redcross,DC=net"
		#Mailbox Creation
		$objUser = $objOU.Create("user",$UserCN)
		$objUser.SetInfo()
		
		$objUser.PUT("sAMAccountName",$Logonname)
		$objUser.PUT("sn",$lastname)
		$objUser.PUT("displayName",$lastname)
		$objUser.PUT( "targetAddress","SMTP: $UserTarget")
		$objUser.PUT("mail",$UserMail)
		$objUser.PUT(" description ",'SMB')
		$objUser.PUT(" userPrincipalName ",$UserUPN )
		$objUser.putex(3,"proxyAddress",@("SMTP:$UserMail","smtp:$UserTarget"))
		
#		$objUser.proxyAddresses += "SMTP: $UserMail
#		$objUser.proxyAddresses += " smtp:$UserTarget
		$objUser.SetInfo()
		
		$objUser.psbase.Invoke("SetPassword","myGr88PwR3dCR0$ $1946 !")
		$objUser.psbase.InvokeSet("AccountDisabled","False")
		$objuser.psbase.CommitChanges()
		
		$flag = $objUser.userAccountControl.value -bxor 65536
		$objUser.userAccountControl = $flag
		$objUser.SetInfo()
		
		$objUser.PUT("manager",$MgrDN)
		$objUser.SetInfo()
		
		$objGroup = $objGrpOU.Create("group",$GroupCN)
		$objGroup.Put("sAMAccountName",$GroupSamName)
		$objGroup.SetInfo()
		
		$objGroup.PUT("mail",$GroupMail)
		$objGroup.PUT( "description",$GroupName )
		$objGroup.putex(3,"proxyAddress",@("SMTP:$GroupMail","smtp:$GroupTarget"))
#		$objGroup.proxyAddresses += " SMTP:$GroupMail
#		$objGroup.proxyAddresses += "smtp: $GroupTarget
		$objGroup.PUT("GroupType","2147483656")
		$objGroup.SetInfo()
			
		$objGroup.PUT(" managedBy ",$MgrDN)
		$objGroup.SetInfo()
		Write-Output The User Account $creationArray[1] was created.
		Write-Output The group $GroupName was created.
		}
}

function addUserstoDelegates([array]$delegateArray)
	{
#Add-PSSnapin Quest.ActiveRoles.ADManagement
#connect-QADService -service 'rcodc3chi001.archq.ri.redcross.net'
	if($delegateArray[2] -eq "ARCHQ")
		{
		#connect-QADService -service 'rcodc3chi001.archq.ri.redcross.net'
		Import-Csv $delegateArray[4] | ForEach-Object {
			
			#Add User to Delegates group
			
			Write-Output $_."User"
			}
		}
	elseif($delegateArray[2] -eq "BIO")
		{
		#connect-QADService -service 'rcodc3chi007.bio.ri.redcross.net'
		Import-Csv $delegateArray[4] | ForEach-Object {
			
			#Add User to Delegates group
			
			Write-Output $_. "User"
			}
		}
	elseif($delegateArray[2] -eq "RC")
		{
		#connect-QADService -service 'rcodc3chi003.rc.ri.redcross.net'
		Import-Csv $delegateArray[4] | ForEach-Object {
			
			#Add User to Delegates group
			
			Write-Output $_. "User"
			}
		}
	
	}

function O365LicenseCheck([string]$checkUser)
	{
	$standardUser = Get-MsolUser -UserPrincipalName Patrick.Garrett@redcross.org
	$currentUser = Get-MsolUser -UserPrincipalName $checkUser 2> .\tempUser.txt
	Start-Sleep -Seconds 1
	$theError = Get-Content -Path.\tempUser.txt
	if($theError -match "User Not Found." -eq "Get-MsolUser : User Not Found.  User: " + $checkUser + ".")
		{
		
		Write-Output User Account Does not Exist,please try again later
		}
	else
		{
		Write-Output User Account Found,continuing Script `r`n
		$LicenseStatus = $currentUser.IsLicensed
		if($LicenseStatus -eq $True)
			{
			Write-Output Account Already Licensed `r`n
			
			if($currentUser.UsageLocation -eq "US")
				{
				Write-Output Account UsageLocation currently set to: $currentUser.UsageLocation `r`n
				
				}
			$checkLicense = $currentUser.Licenses | Select-Object -ExpandProperty AccountSkuID
			if($checkLicense -eq "AmericanRedCross:STANDARDWOFFPACK")
				{
				Write-Output Account Currently Licensed to:Microsoft Office 365 Plan E2 `r`n
				
				}
			if((Compare-Object $currentUser.Licenses[0].ServiceStatus $standardUser.Licenses[0].ServiceStatus).length -eq 0)
				{
				Write-Output Account Currently licensed to all sub-plans `r`n
				
				}
			}
		else
			{
			Write-Output Account not licensed,assigning license
			set-MsolUser -userPrincipalName $currentUser.UserPrincipalName -UsageLocation US
			set-MsolUserLicense -UserPrincipalName $currentUser.UserPrincipalName -AddLicenses AmericanRedCross:STANDARDWOFFPACK
			Write-Output Refreshing User Data
			$currentUser = Get-MsolUser -UserPrincipalName $checkUser
			$recheckUser = $currentUser.UsageLocation
			$checkLicense = $currentUser.Licenses | Select-Object -ExpandProperty AccountSkuID
			$standardUser = Get-MsolUser -UserPrincipalName Patrick.Garrett@redcross.org
			if($recheckUser -eq "US")
				{
				Write-Output Account UsageLocation successfully set to:$recheckUser.UsageLocation
				}
			if($checkLicense -eq "AmericanRedCross:STANDARDWOFFPACK")
				{
				Write-Output Account Successfully Licensed to:Microsoft Office 365 Plan E2
				}
			Remove-Variable recheckUser
			if((Compare-Object $checkLicense.ServiceStatus $standardUser.Licenses[0].ServiceStatus).length -eq 0)
				{
				Write-Output Account Currently licensed to all sub-plans
				}
			}
		}
	Write-Output Purging all Variables
	Remove-Variable checkUser
	Remove-Variable -ErrorAction SilentlyContinue currentUser
	Remove-Variable -ErrorAction SilentlyContinue LicenseStatus
	Remove-Variable -ErrorAction SilentlyContinue checkLicense
	Remove-Variable -ErrorAction SilentlyContinue theError
	Remove-Variable -ErrorAction SilentlyContinue standardUser
	#Remove-Item.\tempUser.txt
	}