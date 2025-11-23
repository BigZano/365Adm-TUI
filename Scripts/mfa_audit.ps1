#  Import-Module Microsoft.Graph

Write-Host "Starting MFA Audit Script..." -ForegroundColor Cyan
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
# Connect to MgGraph
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "UserAuthenticationMethod.Read.All" -ContextScope Process

Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
Write-Host "Retrieving all licensed users..." -ForegroundColor Yellow
# Get licensed users and include Id property
$licensedUsers = Get-MgUser -All -Filter "accountEnabled eq true" -Select "Id,DisplayName,UserPrincipalName,AssignedLicenses"
$licensedUsers = $licensedUsers | Where-Object { $_.AssignedLicenses.Count -gt 0 }
Write-Host "Found $($licensedUsers.Count) licensed users" -ForegroundColor Green

$usersWithMfaStatus = @()
$total = $licensedUsers.Count
$count = 0

Write-Host "`nStarting MFA status check for each user..." -ForegroundColor Yellow

foreach ($user in $licensedUsers) {
    $count++
    Write-Progress -Activity "Checking MFA status" -Status "$count of $total" -PercentComplete (($count / $total) * 100)
    $mfaStatus = "Disabled"
    try {
        $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop
        if ($authMethods) {
            if ($authMethods | Where-Object { $_.AdditionalProperties['@odata.type'] -match 'MicrosoftAuthenticator|PhoneAuthenticationMethod|Fido2' }) {
                $mfaStatus = "Enabled"
            }
        }
    } catch {
        Write-Warning "Error for $($user.UserPrincipalName): $_"
        $mfaStatus = "Unknown"
    }

    $usersWithMfaStatus += [PSCustomObject]@{
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        MfaStatus = $mfaStatus
    }

    Start-Sleep -Milliseconds 300 # Slight delay to avoid throttling
}

$sortedUsersWithMfaStatus = $usersWithMfaStatus | Sort-Object -Property MfaStatus, UserPrincipalName

Write-Host "`nExporting results to CSV..." -ForegroundColor Yellow
$sortedUsersWithMfaStatus | Export-Csv -Path "LicensedUsersWithMfaStatus.csv" -NoTypeInformation

# Display summary
$mfaEnabled = ($sortedUsersWithMfaStatus | Where-Object { $_.MfaStatus -eq "Enabled" }).Count
$mfaDisabled = ($sortedUsersWithMfaStatus | Where-Object { $_.MfaStatus -eq "Disabled" }).Count
$mfaUnknown = ($sortedUsersWithMfaStatus | Where-Object { $_.MfaStatus -eq "Unknown" }).Count

Write-Host "`nAudit Complete!" -ForegroundColor Cyan
Write-Host "Summary:" -ForegroundColor White
Write-Host "- Users with MFA Enabled: $mfaEnabled" -ForegroundColor Green
Write-Host "- Users with MFA Disabled: $mfaDisabled" -ForegroundColor Red
Write-Host "- Users with Unknown Status: $mfaUnknown" -ForegroundColor Yellow
Write-Host "`nResults have been exported to 'LicensedUsersWithMfaStatus.csv'" -ForegroundColor Cyan
