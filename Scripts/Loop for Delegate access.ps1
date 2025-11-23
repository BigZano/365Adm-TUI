# Import Exchange Online module and connect to Exchange Online
Import-Module ExchangeOnlineManagement


# Prompt for EXO admin email for connection
Write-Host "Please enter your EXO admin email address for connection:" -ForegroundColor Cyan
$exoAdminEmail = Read-Host
Connect-ExchangeOnline -UserPrincipalName $exoAdminEmail

# Prompt for the target user email to search for delegate access
Write-Host "Please enter the target user email address to search for delegate access:" -ForegroundColor Yellow
$targetUserEmail = Read-Host

# Initialize array to hold permissions
$script:permissionsReport = [System.Collections.ArrayList]::new()

# Get all mailboxes, shared mailboxes, and distribution lists
$mailboxes = Get-Mailbox -ResultSize Unlimited
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
$distributionLists = Get-DistributionGroup -ResultSize Unlimited

# Function to process permissions
function Process-Permissions($identity, $type, $targetUserEmail) {
    # Retrieve "Send As" permissions for the target user
    $sendAsPermissions = Get-RecipientPermission -Identity $identity.Identity | Where-Object { $_.Trustee -eq $targetUserEmail -and $_.AccessRights -eq 'SendAs' }
    foreach ($permission in $sendAsPermissions) {
        $script:permissionsReport.Add([PSCustomObject]@{
            ObjectType  = $type
            Object      = $identity.PrimarySmtpAddress
            Trustee     = $targetUserEmail
            AccessRight = "Send As"
        }) | Out-Null
    }

    # Retrieve "Send on Behalf" permissions for the target user
    if ($identity.GrantSendOnBehalfTo) {
        foreach ($delegate in $identity.GrantSendOnBehalfTo) {
            # Get the recipient details to compare with target email
            try {
                $delegateRecipient = Get-Recipient -Identity $delegate -ErrorAction Stop
                if ($delegateRecipient.PrimarySmtpAddress -eq $targetUserEmail) {
                    $script:permissionsReport.Add([PSCustomObject]@{
                        ObjectType  = $type
                        Object      = $identity.PrimarySmtpAddress
                        Trustee     = $targetUserEmail
                        AccessRight = "Send on Behalf"
                    }) | Out-Null
                }
            } catch {
                # If we can't resolve the delegate, skip it
                Write-Warning "Could not resolve delegate: $delegate"
            }
        }
    }

    # Retrieve "Full Access" permissions for the target user (mailboxes and shared mailboxes only)
    if ($type -ne "Distribution List") {
        $fullAccessPermissions = Get-MailboxPermission -Identity $identity.Identity | Where-Object { $_.User -eq $targetUserEmail -and $_.IsInherited -eq $false -and $_.AccessRights -eq 'FullAccess' }
        foreach ($permission in $fullAccessPermissions) {
            $script:permissionsReport.Add([PSCustomObject]@{
                ObjectType  = $type
                Object      = $identity.PrimarySmtpAddress
                Trustee     = $targetUserEmail
                AccessRight = "Full Access"
            }) | Out-Null
        }
    }
}

 # Process regular mailboxes
foreach ($mailbox in $mailboxes) {
    Process-Permissions $mailbox "Mailbox" $targetUserEmail
}

# Process shared mailboxes
foreach ($sharedMailbox in $sharedMailboxes) {
    Process-Permissions $sharedMailbox "Shared Mailbox" $targetUserEmail
}

# Process distribution lists
foreach ($distributionList in $distributionLists) {
    Process-Permissions $distributionList "Distribution List" $targetUserEmail
}

# Display the report
$permissionsReport | Format-Table -AutoSize

# Export to CSV
$permissionsReport | Export-Csv -Path ".\MailboxAndGroupPermissions.csv" -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false