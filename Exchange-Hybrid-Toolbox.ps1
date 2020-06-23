#this script has functions for the most common exchange tasks I have to perform at work. These are things that usually take much longer using the GUI
#You will have to edit it just a bit to fit your needs, this ranges from entering your connection URI's, and editing new user OU, name formatting, and 
#licensing info to match your company.

#import-module C:\path_to_script on your regular powershell Window to get this script to work.
Write-Output "Type the following Command to connect to Exchange servers and MSOnline: Connect-ToServers."
Write-Output "Type the following command to start the Exchange toolbox: Run-Toolbox"

#Begins PowerShell Remote session in Exchange and O365
$cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "<your connection uri/PowerShell/>" -Authentication Kerberos -Credential $cred
Import-PSSession $Session -DisableNameChecking

#Connect to Exchange Online
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber

Connect-MsolService -Credential $cred

#This Run-Toolbox function is here so that the script does not reconnect to all the powershell sessions once they are already open. So far I can only get this
#to work in powershell ISE. Regular powershell closes/stops after Connect-Msolservice. 
function Run-Toolbox {

function Create-User
{
    #Populate Variables used in the new user creation, you will have to edit some of this for your own company needs. We use this naming convention as do many
    #companies I have worked at, but yours may be different. For the OU, we use time zones, etc. Yours may be different.
     $UserName = Read-Host 'Please enter the first name in the following format: Bob Joe'
     $Password = Read-Host 'Please enter the temporary password' -AsSecureString
     $FirstName,$LastName = $UserName.split(' ')
     $DisplayName = $lastName + ', ' + $firstName
     $SamAccountName = $firstName + '.' + $lastName
     $domain = '@yourdomain.com'
     $Email = $FirstName + '.' + $LastName + $domain
     $title = Read-Host "Please enter user Title"
     $branch = Read-Host "Please enter the branch number here"
     $Description = $Branch + ' / ' + $title
     $OU = Get-ADOrganizationalUnit -Filter 'Name -like $branch' | select distinguishedName -ExpandProperty distinguishedname

    #Creates User AD account and O365 Mailbox On prem first, which later syncs to O365
    New-RemoteMailbox -Name $DisplayName -FirstName $FirstName -LastName $LastName -DisplayName $DisplayName -SamAccountName $SamAccountName -Confirm:$false -PrimarySmtpAddress $Email -Password $Password -UserPrincipalName $Email -ResetPasswordOnNextLogon $true
    
    #There are sometimes delays from our sync from Exchange on prem to AD, so I put a 30 second timer just in case.
    Start-Sleep -Seconds 30

    #manipulate AD object after allowing 30 seconds to sync, how long it typically takes our environment to sync. Your replication may be faster.
    #We have core groups, but you could also make this "smarter" by using If statements to apply groups based on title/description.
    Add-ADGroupMember -identity Group1 -Members $SamAccountName
    Add-ADGroupMember -identity "Group 3" -Members $SamAccountName
    Set-ADUser -Identity $SamAccountName -Description $description -Title $title 

    #Move to OU, first get the Root OU from user, then specify the target OU. Our OU's have a "Users" sub folder so I added that to the $TargetOU
     $CN = get-aduser -identity $samAccountName -properties DistinguishedName | Select distinguishedname -expandproperty distinguishedName
     $TargetOU = Get-ADOrganizationalUnit -Filter 'Name -like $branch' | select distinguishedname -ExpandProperty distinguishedname 
     $TargetOU = "OU=Users," + $TargetOU
     Move-ADObject -Identity $CN -TargetPath $TargetOU
}

function Enable-OOO {
$identity = Read-Host "Enter the username of the person you would like to turn OOO on for" 
$message = Read-Host "Please enter the message you would like to add to the mailbox"
Set-MailboxAutoReplyConfiguration -Identity $identity -AutoReplyState Enabled -InternalMessage "$message" -ExternalMessage "$message"
}

function Disable-OOO {
$identity = Read-Host "Enter the username of the person you would like to turn OOO off for" 
Set-MailboxAutoReplyConfiguration -Identity $identity -AutoReplyState Disabled
}

#You will have to get your own license SKU's by running the Get-MsolAccountSku command. From there you can specify which licenses you would like to remove
#This applies to the EnableUserAccess and Disable-UserAccessfunctions

#At our organization, admins disable user access by disabling the AD account, convert the mailbox to shared, and remove any groups/licenses
#We laid off people due to Covid, so we had to eventually enable their accounts again. This function makes that easy (instead of going to 3 places)
function ReEnable-UserAccess{
$username = Read-Host "Enter username of the person you would like to re-enable"
Enable-ADAccount -Identity $username

#older versions of exchange have an issue where the RemoteRecipientType changes to "Migrated" when converted to shared. This manually fixes that in the AD object.
set-aduser $username -replace @{msExchRemoteRecipientTYpe="1"} 
#recipient type of 1 = ProvisionMailbox

#Change this to match your groups that were removed if applicable
Add-ADGroupMember -Identity Group1 -Members $username
Add-ADGroupMember -Identity "Group 2" -Members $username
Start-sleep -seconds 15
 
#Convert back to normal mailbox
Set-RemoteMailbox -Identity $username -Type regular
Set-MailboxAutoReplyConfiguration -Identity $username -AutoReplyState Disabled
 
#Disabled the set-MSOL since our AD licenses sync to O365. Kept in case I ever need it.
#Set-MsolUserLicense -UserPrincipalName "$identity@nsm-seating.com" -AddLicenses "nsmseating:ENTERPRISEPACK"
}

#I just made this quick function just in case I needed it. usually user access is disabled far before it gets to me.
function Disable-UserAccess {
$identity = Read-Host "Enter the username of the persons mailbox you would like to convert to shared"
Set-RemoteMailbox -Identity $identity -Type shared
Set-MsolUserLicense -UserPrincipalName "$identity@yourdomain.com" -RemoveLicenses "<Your License SKU here>"
}
 
function Approve-MobileDevice {
$identity = Read-Host "Enter username, such as bob.joe"
Get-MobileDevice -Mailbox $identity | fl FriendlyName, Identity, DeviceAccessState, DeviceID 
"Copy the the DeviceId of the quarantined phone for the next part"
$deviceID = Read-host "Copy the the DeviceId of the quarantined phone and paste DeviceID here"
Set-CASMailbox -identity $identity -ActiveSyncAllowedDeviceIDs @{add= $deviceID}
}

function Delegate-Mailbox {
$identity = Read-Host "Enter the username of the user whos mailbox you want to delegate, in the firstName.lastName format"
$trustee = Read-Host "Enter the username of the trustee, firstName.lastName (person who is getting access)"
Add-MailboxPermission -Identity $identity -User $trustee -AccessRights FullAccess
}

function Remove-DelegatedMailbox {
$identity = Read-Host "Enter the username of the user whos mailbox you want to remove permissions from, in the firstName.lastName format"
$trustee = Read-Host "Enter the username of the person you want removed, firstName.lastName "
Remove-MailboxPermission -Identity $identity -User $trustee -AccessRights FullAccess
}

Switch(Read-Host 'Select from the following options: 
_________________________________________________________________________________________________

       1. Create a New User (on Prem, syncs to O365)

       2. Enable OOO for a user
        
       3. Disable OOO for a user

       4. Disable user access (convert mailbox to shared and remove E3 license)

       5. Enable user access (Convert mailbox to regular and assign E3 license)

       6. Approve a mobile device

       7. Delegate mailbox

       8. Remove delegated mailbox permissions

       9. Exit
_________________________________________________________________________________________________ 
      
       Selection
       ===>') {

   1{Create-User}
   2{Enable-OOO}
   3{Disable-OOO}
   4{Disable-UserAccess}
   5{ReEnable-UserAccess}
   6{Approve-MobileDevice}
   7{Delegate-Mailbox}
   8{Remove-DelegatedMailbox}
   9{break -noexit}
}
}




