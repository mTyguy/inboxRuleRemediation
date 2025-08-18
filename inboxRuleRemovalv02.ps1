# version 0.2 #
# Define the Application (Client) ID, Secret, and Tenant ID
$ApplicationClientId = 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX'
$ApplicationClientSecret = ''
$TenantId = 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX'

############################
#Do not edit below this line
############################

###
#Authenticate to Mg-Graph

# Convert the Client Secret to a Secure String
$SecureClientSecret = ConvertTo-SecureString -String $ApplicationClientSecret -AsPlainText -Force
# Create a PSCredential Object Using the Client ID and Secure Client Secret
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationClientId, $SecureClientSecret
# Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

###
#Define Functions

#Create Interactive Menu
function Show-Menu {
   Write-Host “1) Press ‘1’ to VIEW the paramters of the inbox rule"
   Write-Host “2) Press ‘2’ to REMOVE the inbox rule”
   Write-Host “3) Press ‘3’ to RELOAD the E-MAIL ADDRESS”
   Write-Host “4) Press ‘4’ to RELOAD the inbox rule NAME”
   Write-Host “Q) Press ‘q’ to quit.”
}

#Create Inbox Rule Search function
function InboxRule-Search {
   try {
      $InboxRuleDisplayNameSearch = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules?`$filter=displayName eq '$inboxruledisplayname'" -OutputType PSObject -ErrorVariable InboxRuleSearchError -ErrorAction stop).value.displayName
         #create if statement to stop early if inbox rule is not present
         if
            ([string]::IsNullOrWhiteSpace($InboxRuleDisplayNameSearch)) {
               
               Write-Host "No Inbox Rule present for $impactedUserEmail with the name $inboxruledisplayname"
               break
            }
         else {
            Write-Host "Success! See the Rule paramters below:"
            Write-Host "========================================================="
            Write-Host "Display Name =" $InboxRuleDisplayName
            }

      Write-Host "Is Enabled? =" (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules?`$filter=displayName eq '$inboxruledisplayname'" -OutputType PSObject -ErrorVariable InboxRuleSearchError -ErrorAction stop).value.isEnabled

      Write-Host "Conditions =" (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules?`$filter=displayName eq '$inboxruledisplayname'" -ErrorVariable InboxRuleSearchError -ErrorAction stop).value.conditions

      Write-Host "Actions =" (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules?`$filter=displayName eq '$inboxruledisplayname'" -OutputType PSObject -ErrorVariable InboxRuleSearchError -ErrorAction stop).value.actions

      Write-Host "--------------------------------------------------"

      Write-Host "Here is the full Json Respone:" (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules?`$filter=displayName eq '$inboxruledisplayname'" -OutputType Json -ErrorVariable InboxRuleSearchError -ErrorAction stop)
   }
   catch {
      Write-Host "An Error Occured, See the Below" 
      "========================================================="
      $InboxRuleSearchError.Message
   }
}

#Create Inbox Rule Remove function
function InboxRule-Remove {
   try {
      $getInboxRuleId = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules?`$filter=displayName eq '$inboxruledisplayname'" -OutputType PSObject -ErrorVariable InboxRuleRemovalSearchError -ErrorAction stop -StatusCodeVariable InboxRuleRemovalSearchStatusCode).value.id
         #create if statement to stop early if inbox rule is not present
         if
            ([string]::IsNullOrWhiteSpace($getInboxRuleId)) {
               Write-Host "No Inbox Rule present for $impactedUserEmail with the name $inboxruledisplayname"
               break
            }
         else {}

      Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$impactedUserEmail/mailFolders/inbox/messageRules/$getInboxRuleId" -ErrorVariable InboxRuleRemovalError -ErrorAction stop -StatusCodeVariable InboxRuleRemovalErrorStatusCode

         #create message if rule removed or not
         if
            ($InboxRuleRemovalErrorStatusCode -eq "204") {
               Write-Host "successfully removed the inbox rule" $inboxruledisplayname
            }
         else {}
      }

   catch {
      Write-Host "An Error Occured, See the Below" 
      "========================================================="
      $InboxRuleRemovalSearchError.Message
      $InboxRuleRemovalError.Message
   }
}

###
#Script

cls
Write-Host “================ Inbox Rule Remover v.02 ================”
Write-Host “Please End Script utilizing Q when you are done :)"

# Ask script executor for E-mail Address of the impacted user.
$impactedUserEmail = Read-Host -Prompt "Enter the impacted user's E-mail address (not case sensitive)"

# Ask script executor for Display Name of the inbox rule.
$inboxruledisplayname = Read-Host -Prompt "Enter the Inbox Rule Display Name (case sensitive)"

Write-Host “========================================================="
Write-Host “Loaded E-mail Address =" $impactedUserEmail
Write-Host “Loaded Inbox Rule Name =" $inboxruledisplayname 
Write-Host “========================================================="

do {
   Show-Menu
      $input = Read-Host “What would you like to do?”
      switch ($input)
         {
         ‘1’ {
            InboxRule-Search
            } 
         ‘2’ {
            InboxRule-Remove
            } 
         ‘3’ {
            $impactedUserEmail = Read-Host -Prompt "Enter the impacted user's E-mail address (not case sensitive)"
            Write-Host “========================================================="
            Write-Host “Loaded E-mail Address =" $impactedUserEmail
            Write-Host “Loaded Inbox Rule Name =" $inboxruledisplayname 
            Write-Host “========================================================="
            }
         ‘4’ {
            $inboxruledisplayname = Read-Host -Prompt "Enter the Inbox Rule Display Name (case sensitive)"
            Write-Host “========================================================="
            Write-Host “Loaded E-mail Address =" $impactedUserEmail
            Write-Host “Loaded Inbox Rule Name =" $inboxruledisplayname 
            Write-Host “========================================================="
            }
        'q’ {
            Disconnect-MgGraph | Out-Null
            Write-Host "Thanks for closing the script correctly :)"
            return
            }
        }
   pause
}
until ($input -eq ‘q’)
