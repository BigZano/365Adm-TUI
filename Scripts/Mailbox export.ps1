# Connect to Exchange Online
Write-Host "Enter your Exchange Online admin email:" -ForegroundColor Cyan
$adminUPN = Read-Host
Connect-ExchangeOnline -UserPrincipalName $adminUPN

# Connect to Microsoft Graph for license information
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "User.Read.All" -NoWelcome

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Initialize an array to store mailbox information
$mailboxInfo = @()

# Loop through each mailbox
foreach ($mailbox in $mailboxes) {
    # Get mailbox statistics
    $stats = Get-MailboxStatistics $mailbox.UserPrincipalName

    # Get license information using Microsoft Graph
    try {
        $mgUser = Get-MgUser -UserId $mailbox.UserPrincipalName -Property "AssignedLicenses" -ErrorAction Stop
        $licenses = ($mgUser.AssignedLicenses.SkuId | ForEach-Object {
            try {
                (Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $_ }).SkuPartNumber
            } catch {
                $_
            }
        }) -join ', '
    } catch {
        $licenses = "Unable to retrieve"
        Write-Warning "Could not get licenses for $($mailbox.UserPrincipalName): $_"
    }

    # Create a custom object with mailbox information
    $mailboxObject = [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        UserPrincipalName = $mailbox.UserPrincipalName
        MailboxType = $mailbox.RecipientTypeDetails
        TotalItemSize = $stats.TotalItemSize.Value.ToMB()
        Licenses = ($licenses -join ', ')
    }

    # Add the mailbox object to the array
    $mailboxInfo += $mailboxObject
}

# Export the mailbox information to CSV
$outputPath = ".\MailboxReport.csv"
$mailboxInfo | Export-Csv -Path $outputPath -NoTypeInformation

Write-Host "`nMailbox report exported to: $outputPath" -ForegroundColor Green

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from services successfully." -ForegroundColor Green