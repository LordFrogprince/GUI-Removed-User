######################################
#            Remove User             #
#                                    #
# Used to remove, move and set email #
#           delegation               #
#                                    #
#      Version 6.5 by Colin          #
######################################

# V3.0 This version queries AD for Exchange Servers via Group Memberships, Sets Delegation to the accounts Manager, or via another account via input.
# All AD accounts are verrified using LDAP queries and if a failure script stops. 
#
# V4.0 This version adds the query to microsoft online and removes the assigned license if any on the supplied user account.
# All outputs for delegation, status are now outputted via write-host for easier location of data
#
# V5.0 Adds GUI Interface
#
# V6.0 Changes the logic on the Email Delegation, Changes the marking account as disabled, and removes the Manager, Office Phone and sets Removed Description.
#
# v6.5 Adds the Checkbox for the Forward Email to Delegation


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '630,400'
$Form.text                       = "Remove User v6.5"
$Form.TopMost                    = $false

$UserID                          = New-Object system.Windows.Forms.TextBox
$UserID.multiline                = $false
$UserID.width                    = 177
$UserID.height                   = 20
$UserID.location                 = New-Object System.Drawing.Point(16,50)
$UserID.Font                     = 'Microsoft Sans Serif,10'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "User Account To Be Removed"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(16,27)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "@domain.com"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(199,54)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$EmailDelegation                 = New-Object system.Windows.Forms.Groupbox
$EmailDelegation.height          = 225
$EmailDelegation.width           = 270
$EmailDelegation.text            = "Email Delegation if Required"
$EmailDelegation.location        = New-Object System.Drawing.Point(14,89)

$MDelegation                     = New-Object system.Windows.Forms.CheckBox
$MDelegation.text                = "Managerial Delegation Requested"
$MDelegation.AutoSize            = $false
$MDelegation.width               = 220
$MDelegation.height              = 20
$MDelegation.location            = New-Object System.Drawing.Point(9,30)
$MDelegation.Font                = 'Microsoft Sans Serif,10'

$OUDelegation                    = New-Object system.Windows.Forms.CheckBox
$OUDelegation.text               = "Other User Delegation Requested"
$OUDelegation.AutoSize           = $false
$OUDelegation.width              = 220
$OUDelegation.height             = 20
$OUDelegation.location           = New-Object System.Drawing.Point(9,62)
$OUDelegation.Font               = 'Microsoft Sans Serif,10'

$UserDel                         = New-Object system.Windows.Forms.TextBox
$UserDel.multiline               = $false
$UserDel.width                   = 140
$UserDel.height                  = 20
$UserDel.location                = New-Object System.Drawing.Point(19,90)
$UserDel.Font                    = 'Microsoft Sans Serif,10'

$Forward                         = New-Object system.Windows.Forms.CheckBox
$Forward.text                    = "Forward Email to Delegation"
$Forward.AutoSize                = $false
$Forward.width                   = 220
$Forward.height                  = 20
$Forward.location                = New-Object System.Drawing.Point(45,120)
$Forward.Font                    = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "@domain.com"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(167,95)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$OK                              = New-Object system.Windows.Forms.Button
$OK.text                         = "Run"
$OK.width                        = 60
$OK.height                       = 30
$OK.location                     = New-Object System.Drawing.Point(224,339)
$OK.Font                         = 'Microsoft Sans Serif,10'

$Output                          = New-Object system.Windows.Forms.TextBox
$Output.multiline                = $true
$Output.width                    = 318
$Output.height                   = 300
$Output.location                 = New-Object System.Drawing.Point(299,17)
$Output.Font                     = 'Microsoft Sans Serif,10'
$output.ForeColor                = [Drawing.Color]::Red
$output.ScrollBars               = "Vertical"

$Exit                            = New-Object system.Windows.Forms.Button
$Exit.text                       = "Exit"
$Exit.width                      = 60
$Exit.height                     = 30
$Exit.location                   = New-Object System.Drawing.Point(557,339)
$Exit.Font                       = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($UserID,$Label1,$Label2,$EmailDelegation,$OK,$Output,$Exit))
$EmailDelegation.controls.AddRange(@($MDelegation,$OUDelegation,$UserDel,$Label3,$Forward))

#region gui events {
$UserID.Add_TextChanged({  })
$UserDel.Add_TextChanged({  })
$Output.Add_TextChanged({  })
$MDelegation.Add_CheckedChanged({  })
$Forward.Add_CheckedChanged({  })
$OUDelegation.Add_CheckedChanged({  })
$Exit.Add_Click({  })
$OK.Add_Click({  })
#endregion events }

#endregion GUI }

#Write your logic code here

#### Modules ######

    # Microsoft Online Install
        Import-Module MSOnline

#####  Variables #####

#Set Global Variables to False
    $global:userdm = $false
    $global:UserOD = $false

#Set User Delegation to False
    $UserDel.Enabled = $false

#### Script Logic ####

#Managerial Delegation Checkbox Logic
    $MDelegation.Add_CheckedChanged({
        if ($MDelegation.Checked -eq $true) {
                $OUDelegation.Checked = $false
                $global:Userdm = $true}
        else { $global:Userdm = $false}
                 
})
    
#User Delegation Checkbox Logic
    $OUDelegation.Add_CheckedChanged({
        if ($OUDelegation.Checked -eq $true) { 
            $UserDel.Enabled = $true 
            $MDelegation.Checked = $false
            $global:UserOD = $true}
        else { $userDel.Enabled = $false 
               $global:UserOD = $false} 
})

###### Running Script with Settings and Variables from Powershell Form
$OK.Add_Click({  
        # Set Date for and description for Removed Date
            $date = get-date -Format "yyyyMM"
            $Des = "$date"


        # Unload all Modules
             Remove-Module -Name MSOnline -ErrorAction Ignore
             Get-PSSession | Remove-PSSession               
	
    	#Set Error Action Preference - Stops the Script on any Error
    		$ErrorActionPreference = "Stop"
       	
		#Sets the Exchanger Server without Prompting or being Hardcoded
			$server = (Get-ADGroupMember "exchange Servers" | select-object -expandProperty samaccountname -First 1)
			$server = $server.TrimEnd('$')
			
		
		#Sets the Technician Credential to Remove the User License from Office 365
			$UserCredential = Get-Credential -Message "Office365 Login - Use username@domain.com"
            $output.AppendText("Connecting to Microsoft Office 365")
            $output.AppendText("`n`n")
			Connect-MsolService -Credential $UserCredential
		
		# Connect to Exchange Powershell
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$server/PowerShell/ -Authentication Kerberos
            $output.AppendText("Importing Exchange Powershell")
            $output.AppendText("`n`n")
			Import-PSSession $Session
		
		# Connect to Active Directory Powershell
			import-module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory\ActiveDirectory.psd1
		
	    # Set User Variable
			$User = $UserID.Text
			$delegation = $UserDel.Text
        
###### Account Verification ######
	
        # Managerial Account with LDAP Query to verify Account Exists  
		    if ($userdm -eq "True")
	{		$delegation = Get-ADUser $User -Properties Manager | Select-Object @{ n = "ManagerName"; e = { (Get-ADUser $_.Manager).SamAccountName } }
			$delegation = $delegation | select-object -expandProperty managername
			$UserVD = Get-ADUser -LDAPFilter "(sAMAccountName=$delegation)"
				If ($UserVD -eq $Null)
					{ $output.AppendText("User Manager Delegation Account Not Found in Active Directory")
                      $output.AppendText("`n`n")
					exit
					}
				Else
					{$output.AppendText("User Manager Delegation Account  Found in Active Directory")
                     $output.AppendText("`n`n") 
					  $userD = "True"}
	}
	
# Delegation to Non-Managerial Account with LDAP Query to verify Account Exists
		if ($UserOD -eq "True")
			{$UserVD= Get-ADUser -LDAPFilter "(sAMAccountName=$delegation)"
				If ($delegation -eq $Null)
					{ $output.AppendText("User Delegation Account does not exist in Active Directory")
                      $output.AppendText("`n`n")   
						exit
					}
				Else
					{ 	$output.AppendText( "User Delegation Account Found in Active Directory")
                        $output.AppendText("`n`n")  
			   			$userD = "True" }
	}
	
# LDAP Query for verification of User to be Removed
		$UserV = Get-ADUser -LDAPFilter "(sAMAccountName=$User)"
			If ($UserV -eq $Null)
				{ $output.AppendText("Removed User does not exist in Active Directory")
                  $output.AppendText("`n`n")  
				  Exit}
			Else { $output.AppendText("User Found in Active Directory")
                   $output.AppendText("`n`n")
	}
	
#Office 365 License Verification
        $License = (Get-MsolUser -UserPrincipalName ($user + "@domain.com") | Select-Object -ExpandProperty licenses )
		$License = ($License | select -First 1 | Select-Object -ExpandProperty AccountSkuId)
		$licenseString = ($License | Out-String)
		$LicenseMatch = ("Office365License") | Out-String

##### Remove User with Above Variable(s)

#Set Error Action Preference - Continues in this section since Verification is complete
        	$ErrorActionPreference = "Continue"
    
# Set Primary Group SID
    	    $group = get-adgroup "removed-users"
		    Add-ADGroupMember -Identity $group -Members $user
		    $groupSid = $group.sid	
            $GroupID = $groupSid.Value.Substring($groupSid.Value.LastIndexOf("-") + 1)


#Office 365 License Removal
		if ($licenseString -eq $LicenseMatch)
			{Set-MsolUserLicense -UserPrincipalName ($user + "@domain.com") -RemoveLicenses "$license"
			 $output.AppendText("Office 365 Licensed Removed")
             $output.AppendText("`n`n")}
		else { $output.AppendText("Office 365 Licensed not assigned")
               $output.AppendText("`n`n")}
	
#Convert Account to Shared and Disable AD Account
        $alias = Get-Mailbox -Identity $user | Select-Object -ExpandProperty "alias"       
		    if ($alias = $user )
                {Get-Mailbox -Identity $user | Set-Mailbox -Type Shared
		         $output.AppendText("Account Converted to Shared and Disabled") 
                 $output.AppendText("`n`n")}
            else { Disable-ADAccount -Identity $user
                   $output.AppendText("No exchange account found - Disabled AD Account") 
                   $output.AppendText("`n`n") }

	
#Add Mailbox Delegation
		if ($userd -eq "True")
			{	Add-MailboxPermission -Identity $user -user $delegation -AccessRights FullAccess -InheritanceType ALL
				$output.AppendText( "Account Delegated to $delegation") 
                 $output.AppendText("`n`n")   }

#Add Forward to Delegation
    if ($Forward.Checked -eq $true) {
            set-mailbox $user -ForwardingAddress ($delegation + "@domain.com") -DeliverToMailboxAndForward $True
            $output.AppendText("Account Forwarded to $Delegation")
            $output.AppendText("`n`n")   }
	
#Move AD account to Removed Users OU
	
		if ($userd -eq "True")
			{ 	get-aduser -identity $user | Move-ADObject -TargetPath "OU=TARGETDESTINATION"
				$output.AppendText("User Account Moved to Removed Users with a Shared Mailbox OU")
                $output.AppendText("`n`n")    }
		else
			{	get-aduser -identity $user | Move-ADObject -TargetPath "OU=TARGETDESTINATION"
				$output.AppendText("User Account Moved to Removed Users without Forward OU") 
                $output.AppendText("`n`n")}
	
#Set User to Removed Users OU Group Membership 
		Get-ADUser "$user" | Set-ADObject -Replace @{ primaryGroupID = "$GroupID" }
	
#10 Second wait is to allow time for completion of changing primary group
		Start-Sleep -s 10
	
#Remove all Membership except Removed Users OU
		$ADgroups = Get-ADPrincipalGroupMembership -Identity $user | where { $_.Name -ne "removed-users security group" }
		Remove-ADPrincipalGroupMembership -Identity $user -MemberOf $ADgroups -Confirm:$false

#Remove Manager, Description and Office Phone from AD Profile
        Set-ADUser $user -Manager $null -Description $Des -OfficePhone $null
 
#Write to output box that script is complete
    $output.AppendText("Script Completed for $user") 
    $output.AppendText("`n")        				
	

})





#Close the Form - Exit the Script
    $Exit.Add_Click({ $form.close()
})

# Must be at the end of the Script
[void]$Form.ShowDialog()
