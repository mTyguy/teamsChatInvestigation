# Version 0.1 #
# Utilizes Microsoft Graph PowerShell Module
# Required API permissions: Chat.Read.All

# Set Parameters
# We require the UserPrincipalName & the specific ChatID of the Teams chat we are interested in
param (                       
  [Parameter(Mandatory,
  HelpMessage="Enter the e-mail address of the user")]
  [string]$upn,

  [Parameter(Mandatory,
  HelpMessage="Enter the chatID without quotes")]
  [string]$chatId
)

### Authenticate ###
# Define the Application (Client) ID and Secret
$ApplicationClientId = 'XXXXXXXX-XXXX-XXXX-XXXXXXXXXXXX'
$ApplicationClientSecret = '<secret-string>'
$TenantId = 'XXXXXXXX-XXXX-XXXX-XXXXXXXXXXXX'

# Convert the Client Secret to a Secure String
$SecureClientSecret = ConvertTo-SecureString -String $ApplicationClientSecret -AsPlainText -Force

# Create a PSCredential Object Using the Client ID and Secure Client Secret
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ApplicationClientId, $SecureClientSecret
# Connect to Microsoft Graph Using the Tenant ID and Client Secret Credential
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome

### Start script ###
# Grab chat meta data and create array
$chatMetaData = (Invoke-MgGraphRequest -Method GET https://graph.microsoft.com/v1.0/chats/$($chatId)?`$expand=members)

$chatMetaDataArray = [PSCustomObject] @{
  'chatId'                  = $chatMetaData.id
  'chat_Created_Timestamp'  = $chatMetaData.createdDateTime
  'chat_Type'               = $chatMetaData.chatType
  Member1 = [ordered] @{
    'member_Name'   = $chatMetaData.members[0].displayName  
    'tenant_id'     = $chatMetaData.members[0].tenantId
    'member_email' = $chatMetaData.members[0].email
    'member_id'     = $chatMetaData.members[0].id
  }
  Member2 = [ordered] @{
    'member_Name'   = $chatMetaData.members[1].displayName
    'tenant_id'     = $chatMetaData.members[1].tenantId
    'member_email' = $chatMetaData.members[1].email
    'member_id'     = $chatMetaData.members[1].id
  }
}
Write-Host "Chat Meta Data:" -Foreground DarkYellow
$chatMetaDataArray
[Environment]::NewLine

# Summarize the activity
$eventSummary = "At $($chatMetaDataArray.chat_Created_Timestamp), a $($chatMetaDataArray.chat_type) type Teams chat was created between $($chatMetaDataArray.Member1.member_name) ($($chatMetaDataArray.Member1.member_email)) & $($chatMetaDataArray.Member2.member_name) ($($chatMetaDataArray.Member2.member_email))."
Write-Host "Event Summary:" -Foregroundcolor DarkYellow
Write-Host $eventSummary -Foregroundcolor DarkGreen
[Environment]::NewLine

# Display if each chat member is part of home tenant or not
Write-Host "User internal/external status:" -Foregroundcolor DarkYellow
if ($chatMetaDataArray.Member1.tenant_id -eq $TenantId) {
  $Member1_status = "$($chatMetaDataArray.Member1.member_Name) ($($chatMetaDataArray.Member1.member_email)) is an internal user"
  Write-Host "$($chatMetaDataArray.Member1.member_Name) ($($chatMetaDataArray.Member1.member_email)) is an internal user" -Foregroundcolor DarkGreen
  } else {
    $Member1_status = "$($chatMetaDataArray.Member1.member_Name) ($($chatMetaDataArray.Member1.member_email)) is NOT an internal user"
    Write-Host "$($chatMetaDataArray.Member1.member_Name) ($($chatMetaDataArray.Member1.member_email)) is NOT an internal user" -Foregroundcolor DarkRed
}

if ($chatMetaDataArray.Member2.tenant_id -eq $TenantId) {
  $Member2_status = "$($chatMetaDataArray.Member2.member_Name) ($($chatMetaDataArray.Member2.member_email)) is an internal user"
  Write-Host "$($chatMetaDataArray.Member2.member_Name) ($($chatMetaDataArray.Member2.member_email)) is an internal user" -Foregroundcolor DarkGreen
  } else {
    $Member2_status = "$($chatMetaDataArray.Member2.member_Name) ($($chatMetaDataArray.Member2.member_email)) is NOT an internal user"
    Write-Host "$($chatMetaDataArray.Member2.member_Name) ($($chatMetaDataArray.Member2.member_email)) is NOT an internal user" -Foregroundcolor DarkRed
}
[Environment]::NewLine

# Get chat messages
$chatData = (Invoke-MgGraphRequest -Method GET https://graph.microsoft.com/v1.0/users/$upn/chats/$chatId/messages).value

#Display all messages in the chat
Write-Host "See the below messages from this chat:" -Foregroundcolor DarkYellow
foreach ($_ in $chatData) {
  Write-Host "Timestamp: $($_.createdDateTime), Message Author: $($_.from.user.displayname), Message: $($_.body.content)" -ForegroundColor DarkCyan
}
[Environment]::NewLine

# Prep data to save it locally
$saveMessages = foreach ($_ in $chatData) {
  Write-Output "Timestamp: $($_.createdDateTime), Message Author: $($_.from.user.displayname), Message: $($_.body.content)"
  [Environment]::NewLine
}
$outputData = @($chatMetaDataArray;$Member1_status;$Member2_status;$saveMessages)

# Prompt if the user wants to save this data to a file
Write-Host "Would you like to save this output locally?" -Foregroundcolor DarkYellow
$prompt = Read-Host "yes/y or no/n"

if ($prompt -eq "yes" -or $prompt -eq "y") {
  $outputData | Out-File .\teamsmessage.txt
  Write-Host "Output has been saved to local directory, closing script" -Foregroundcolor DarkGreen
} elseif ($prompt -eq "no" -or $prompt -eq "n") {
  Write-Host "Output will not be saved, closing script" -Foregroundcolor DarkGreen
}

# End Mg-Graph session
Disconnect-MgGraph | Out-null
