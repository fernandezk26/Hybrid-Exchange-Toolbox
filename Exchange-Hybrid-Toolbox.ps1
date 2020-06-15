#this script has functions for the most common exchange tasks I have to perform at work. These are things that usually take much longer using the GUI
#You will have to edit it just a bit to fit your needs, this ranges from entering your connection URI's, and editing new user OU, name formatting, and 
#licensing info to match your company.

#Begins PowerShell Remote session in Exchange and O365
$cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "<your connection uri/PowerShell/>" -Authentication Kerberos -Credential $cred
Import-PSSession $Session -DisableNameChecking

#Connect to Exchange Online
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber

Connect-MsolService

#This Run-Toolbox function is here so that the script does not reconnect to all the powershell sessions once they are already open. So far I can only get this
#to work in powershell ISE. Regular powershell closes/stops after Connect-Msolservice. 
function Run-Toolbox {

function Create-User
{
    #Populate Variables used in the new user creation. This is just an example of if you were to create your users in the firstName.lastName format
     $UserName = Read-Host 'Please enter the first name in the following format: Bob Joe'
     $Password = Read-Host 'Please enter the temporary password' -AsSecureString
     $FirstName,$LastName = $UserName.split(' ')
     $DisplayName = $lastName + ', ' + $firstName
     $SamAccountName = $firstName + '.' + $lastName
     $domain = '@yourdomain.com'
     $Email = $FirstName + '.' + $LastName + $domain

      #Select time zone based on selection (many organizations break OU's into time zones)
        $timeZone = Switch(Read-Host 'Type "1" if Eastern, "2" if Central, "3" if Pacific') {
        1 {$zone = "Eastern" ; break}
        2 {$zone = "Central" ; break}
        3 {$zone = "Pacific"; break}
    }   
     $branch = Read-Host 'Please enter the branch number'
     $OU = 'yourdomain.local/Branch Office/' + $zone + '/' + $branch + '/Users'

    #Creates User AD account and O365 Mailbox in Exchange On Prem
    New-RemoteMailbox -Name $UserName -FirstName $FirstName -LastName $LastName -DisplayName $DisplayName -SamAccountName $SamAccountName -Confirm:$false -PrimarySmtpAddress $Email -Password $Password -UserPrincipalName $Email -OnPremisesOrganizationalUnit $OU -ResetPasswordOnNextLogon $true
}

function Enable-OOO
{$identity = Read-Host "Enter the username of the person you would like to turn OOO on for" 
 $message = Read-Host "Please enter the message you would like to add to the mailbox"
 Set-MailboxAutoReplyConfiguration -Identity $identity -AutoReplyState Enabled -InternalMessage "$message" -ExternalMessage "$message"}

function Disable-OOO
{$identity = Read-Host "Enter the username of the person you would like to turn OOO off for" 
 Set-MailboxAutoReplyConfiguration -Identity $identity -AutoReplyState Disabled}

#You will have to get your own license SKU's by running the Get-MsolAccountSku command. From there you can specify which licenses you would like to remove
#This applies to the EnableUserAccess and Disable-UserAccessfunctions

function ReEnable-UserAccess
{$identity = Read-Host "Enter the username of the persons mailbox you would like to convert to shared"
 Set-RemoteMailbox -Identity $identity -Type regular
 Set-MsolUserLicense -UserPrincipalName "$identity@yourdomain.com" -AddLicenses "<Your License SKU here>"
 Set-MailboxAutoReplyConfiguration -Identity $identity -AutoReplyState Disabled
 }

function Disable-UserAccess
{$identity = Read-Host "Enter the username of the persons mailbox you would like to convert to shared"
 Set-RemoteMailbox -Identity $identity -Type shared
 Set-MsolUserLicense -UserPrincipalName "$identity@yourdomain.com" -RemoveLicenses "<Your License SKU here>"
 
function Approve-MobileDevice
{$identity = Read-Host "Enter username, such as bob.joe"
Get-MobileDevice -Mailbox $identity | fl FriendlyName, Identity, DeviceAccessState, DeviceID 
"Copy the the DeviceId of the quarantined phone for the next part"
$deviceID = Read-host "Copy the the DeviceId of the quarantined phone and paste DeviceID here"
Set-CASMailbox -identity $identity -ActiveSyncAllowedDeviceIDs @{add= $deviceID}
}
 }


Switch(Read-Host 'Select "1" if you would like to create a new on-Prem O365 mailbox,
       "2" to enable OOO for a user, 
       "3" to disable OOO for a user, `
       "4" to disable user access (convert mailbox to shared and remove E3 license), 
       "5" to re-enable user access (Convert mailbox to regular and assign E3 license), 
       "6" to Approve a Mobile device in quarantine
       "7" to exit') {

   1{Create-User}
   2{Enable-OOO}
   3{Disable-OOO}
   4{Disable-UserAccess}
   5{ReEnable-UserAccess}
   6{Approve-MobileDevice}
   7{exit}
}
}



#Test and add later
7{Delegate-Mailbox}

function Delegate-Mailbox {
$identity = Read-Host "Enter the username of the user whos mailbox you want to delegate
$trustee = Read-Host "Enter the username of the trustee (person who is getting access)
Add-RecipientPermission -Identity $identity -Trustee $trustee -AccessRights SendAs
}

