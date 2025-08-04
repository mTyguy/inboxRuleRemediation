# version 0.1 #
# Define the Application (Client) ID and Secret
$ApplicationClientId = 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX'
$ApplicationClientSecret = ''
$TenantId = 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX'

# Convert the Client Secret to a Secure String
$SecureClientSecret = ConvertTo-SecureString -String $ApplicationClientSecret -AsPlainText -Force

# Create a PSCredential Object Using the Client ID and Secure Client Secret
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationClientId, $SecureClientSecret

# Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

####
# The Delete call requires the Id of the inbox rule, to get that we need to make a Get request for the rule first.
# We ask the script executor for 1. the user's e-mail address 2. the display name of the rule.
# Can then send the Delete call.
# Requires the application permission MailboxSettings.ReadWrite

# Ask script executor for E-mail Address of the impacted user.
$impactedUserEmail = Read-Host -Prompt "Enter the impacted user's E-mail address (not case sensitive)"

# Ask script executor for Display Name of the inbox rule.
$inboxruledisplayname = Read-Host -Prompt "Enter the Inbox Rule Display Name (case sensitive)"

#Get Inbox Rule by User Principal Name & Display Name

$getInboxRule = (invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules?`$filter=displayName eq '$inboxruledisplayname'" -OutputType PSObject).value

# Set inbox rule Id to variable and then delete the inbox rule
foreach ($_ in $getInboxRule.id) {(Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules/$_")}

# End session, suppress output
Disconnect-MgGraph | Out-Null
