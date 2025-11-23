Add-Type -AssemblyName System.Windows.Forms

# Create and configure SaveFileDialog for CSV
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Title = "Select where to save the output CSV file"
$saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$saveFileDialog.DefaultExt = "csv"
$saveFileDialog.FileName = "UserAuthPolicies.csv"
$saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')

if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $outputFile = $saveFileDialog.FileName
    Write-Host "Saving output to: $outputFile"

    # Connect to ExchangeOnline - prompt or hardcode admin UPN
    $adminUPN = Read-Host "Enter Exchange Online admin user principal name"
    Connect-ExchangeOnline -UserPrincipalName $adminUPN

    Start-Sleep -Seconds 5

    # Get tenant-wide organization config (optional: write separately)
    $orgConfig = Get-OrganizationConfig
    $tenantModernAuthEnabled = $orgConfig.OAuth2ClientProfileEnabled

    # Get authentication policies
    $authPolicies = Get-AuthenticationPolicy | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name
            AllowBasicAuthPop = $_.AllowBasicAuthPop
            AllowBasicAuthImap = $_.AllowBasicAuthImap
            AllowBasicAuthSmtp = $_.AllowBasicAuthSmtp
            AllowBasicAuthActiveSync = $_.AllowBasicAuthActiveSync
            AllowBasicAuthAutodiscover = $_.AllowBasicAuthAutodiscover
            AllowBasicAuthWebServices = $_.AllowBasicAuthWebServices
            AllowBasicAuthPowershell = $_.AllowBasicAuthPowershell
            AllowBasicAuthMAPI = $_.AllowBasicAuthMAPI
        }
    } | Group-Object -Property Name -AsHashTable -AsString

    $tenantDefaultPolicyName = "Default"
    $users = Get-User -ResultSize Unlimited

    $userAuthPolicies = foreach ($user in $users) {
        $policyName = $user.AuthenticationPolicy
        if ([string]::IsNullOrEmpty($policyName)) {
            $policyName = $tenantDefaultPolicyName
        }

        $policySettings = $authPolicies[$policyName]

        $legacyAuthAllowed = $false
        if ($policySettings) {
            $legacyAuthAllowed = (
                $policySettings.AllowBasicAuthPop -or
                $policySettings.AllowBasicAuthImap -or
                $policySettings.AllowBasicAuthSmtp -or
                $policySettings.AllowBasicAuthActiveSync -or
                $policySettings.AllowBasicAuthAutodiscover -or
                $policySettings.AllowBasicAuthWebServices -or
                $policySettings.AllowBasicAuthPowershell -or
                $policySettings.AllowBasicAuthMAPI
            )
        }

        [PSCustomObject]@{
            UserPrincipalName       = $user.UserPrincipalName
            EffectiveAuthPolicy     = $policyName
            LegacyAuthAllowed       = $legacyAuthAllowed
            TenantModernAuthEnabled = $tenantModernAuthEnabled
        }
    }

    # Export user data as CSV
    $userAuthPolicies | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

    Write-Host "User authentication policy report saved to $outputFile"
}
else {
    Write-Host "Save operation cancelled by user. No output generated."
}


# disconnect from ExchOnline
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Green