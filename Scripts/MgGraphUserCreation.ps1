Import-Module Microsoft.Graph

# Verify connection and required scopes
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
    $context = Get-MgContext
    Write-Host "Connected to Microsoft Graph as: $($context.Account)" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Get available licenses once at the start
Write-Host "Retrieving available licenses..." -ForegroundColor Yellow
$availableLicenses = Get-MgSubscribedSku

# Create a mapping of SKU part numbers to friendly names
$skuMapping = @{
    "ENTERPRISEPACK" = "Microsoft 365 E3"
    "ENTERPRISEPREMIUM" = "Microsoft 365 E5"
    "STANDARDPACK" = "Microsoft 365 Business Basic"
    "O365_BUSINESS" = "Microsoft 365 Apps for Business"
    "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM" = "Microsoft 365 Business Standard"
    "SPB" = "Microsoft 365 Business Premium"
    "EXCHANGESTANDARD" = "Exchange Online Plan 1"
    "EXCHANGEENTERPRISE" = "Exchange Online Plan 2"
    "SHAREPOINTSTANDARD" = "SharePoint Online Plan 1"
    "SHAREPOINTENTERPRISE" = "SharePoint Online Plan 2"
    "POWER_BI_PRO" = "Power BI Pro"
    "POWER_BI_STANDARD" = "Power BI (free)"
    "TEAMS_EXPLORATORY" = "Microsoft Teams Exploratory"
    "PHONESYSTEM_VIRTUALUSER" = "Phone System - Virtual User"
    "MCOSTANDARD" = "Skype for Business Online Plan 2"
    "FLOW_FREE" = "Power Automate (free)"
    "POWERAPPS_VIRAL" = "Power Apps (free)"
    "DEVELOPERPACK" = "Microsoft 365 E3 Developer"
    "EMS" = "Enterprise Mobility + Security E3"
    "EMSPREMIUM" = "Enterprise Mobility + Security E5"
    "RIGHTSMANAGEMENT" = "Azure Information Protection Plan 1"
    "VISIOCLIENT" = "Visio Online Plan 2"
    "PROJECTONLINE_PLAN_1" = "Project Online Essentials"
    "PROJECTONLINE_PLAN_2" = "Project Online Professional"
}

if ($availableLicenses.Count -eq 0) {
    Write-Host "No licenses found in tenant." -ForegroundColor Red
    $hasLicenses = $false
} else {
    $hasLicenses = $true
    Write-Host "`nAvailable Licenses:" -ForegroundColor Green
    for ($i = 0; $i -lt $availableLicenses.Count; $i++) {
        $sku = $availableLicenses[$i].SkuPartNumber
        $friendlyName = if ($skuMapping.ContainsKey($sku)) { $skuMapping[$sku] } else { $sku }
        $available = $availableLicenses[$i].PrepaidUnits.Enabled - $availableLicenses[$i].ConsumedUnits
        $total = $availableLicenses[$i].PrepaidUnits.Enabled
        
        # Color code based on availability
        $color = if ($available -gt 0) { "Cyan" } else { "DarkGray" }
        Write-Host "[$($i + 1)] $friendlyName - Available: $available / $total" -ForegroundColor $color
    }
    Write-Host "[0] Skip license assignment" -ForegroundColor Yellow
}

do {
    # Prompt for user details
    $displayName = Read-Host "Enter Display Name"
    
    # UPN with validation
    do {
        $userPrincipalName = Read-Host "Enter User Principal Name (e.g., user@domain.com)"
        
        # Basic UPN validation
        if ($userPrincipalName -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
            Write-Host "Invalid email format. Please try again." -ForegroundColor Red
            $isValidUPN = $false
        } else {
            # Check if user already exists
            try {
                $existingUser = Get-MgUser -UserId $userPrincipalName -ErrorAction SilentlyContinue
                if ($existingUser) {
                    Write-Host "User with UPN '$userPrincipalName' already exists!" -ForegroundColor Red
                    $isValidUPN = $false
                } else {
                    $isValidUPN = $true
                }
            } catch {
                $isValidUPN = $true  # User doesn't exist, which is good
            }
        }
    } while (-not $isValidUPN)
    
    $usageLocation = Read-Host "Enter Usage Location (2-letter country code, e.g., US, CA, GB)"
    
    # Password with validation
    do {
        $password = Read-Host "Enter Password (min 8 chars, must include uppercase, lowercase, number)" -AsSecureString
        $passwordPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
        )
        
        # Basic password validation
        $isValid = $passwordPlain.Length -ge 8 -and 
                   $passwordPlain -cmatch '[A-Z]' -and 
                   $passwordPlain -cmatch '[a-z]' -and 
                   $passwordPlain -cmatch '[0-9]'
        
        if (-not $isValid) {
            Write-Host "Password does not meet requirements. Please try again." -ForegroundColor Red
        }
    } while (-not $isValid)

    # Display user details for confirmation
    Write-Host "`n--- User Details Confirmation ---" -ForegroundColor Yellow
    Write-Host "Display Name: $displayName" -ForegroundColor Cyan
    Write-Host "User Principal Name: $userPrincipalName" -ForegroundColor Cyan
    Write-Host "Usage Location: $usageLocation" -ForegroundColor Cyan
    Write-Host "Mail Nickname: $($userPrincipalName.Split("@")[0])" -ForegroundColor Cyan
    Write-Host "Account Enabled: Yes" -ForegroundColor Cyan
    Write-Host "Force Password Change: No" -ForegroundColor Cyan
    Write-Host "--------------------------------" -ForegroundColor Yellow
    
    $confirmCreate = Read-Host "Create this user? (Y/N)"
    
    if ($confirmCreate -ne "Y" -and $confirmCreate -ne "y") {
        Write-Host "User creation cancelled. Starting over..." -ForegroundColor Yellow
        continue
    }

    # Create user object
    $userParams = @{
        AccountEnabled    = $true
        DisplayName       = $displayName
        UserPrincipalName = $userPrincipalName
        UsageLocation     = $usageLocation.ToUpper()
        MailNickname      = $userPrincipalName.Split("@")[0]
        PasswordProfile   = @{
            Password              = $passwordPlain
            ForceChangePasswordNextSignIn = $false
        }
    }

    # Create the user
    try {
        $newUser = New-MgUser @userParams
        Write-Host "User '$displayName' created successfully." -ForegroundColor Green
        
        # License assignment
        if ($hasLicenses) {
            Write-Host "`nLicense Selection for ${displayName}:" -ForegroundColor Yellow
            $licenseChoice = Read-Host "Enter license number (0 to skip)"
            
            if ($licenseChoice -ne "0" -and $licenseChoice -match '^\d+$') {
                $licenseIndex = [int]$licenseChoice - 1
                if ($licenseIndex -ge 0 -and $licenseIndex -lt $availableLicenses.Count) {
                    $selectedLicense = $availableLicenses[$licenseIndex]
                    $sku = $selectedLicense.SkuPartNumber
                    $friendlyName = if ($skuMapping.ContainsKey($sku)) { $skuMapping[$sku] } else { $sku }
                    $available = $selectedLicense.PrepaidUnits.Enabled - $selectedLicense.ConsumedUnits
                    
                    # Check if license has available capacity
                    if ($available -le 0) {
                        Write-Host "`nWarning: '$friendlyName' has no available licenses!" -ForegroundColor Red
                        Write-Host "Available: $available / $($selectedLicense.PrepaidUnits.Enabled)" -ForegroundColor Yellow
                        Write-Host "To add more licenses, visit Microsoft 365 Admin Center > Billing > Licenses" -ForegroundColor Yellow
                        
                        $proceed = Read-Host "Do you want to try assigning anyway? This will likely fail. (Y/N)"
                        if ($proceed -ne "Y" -and $proceed -ne "y") {
                            Write-Host "License assignment skipped." -ForegroundColor Yellow
                            continue
                        }
                    }
                    
                    try {
                        $licenseParams = @{
                            AddLicenses = @(
                                @{
                                    SkuId = $selectedLicense.SkuId
                                    DisabledPlans = @()
                                }
                            )
                            RemoveLicenses = @()
                        }
                        
                        Set-MgUserLicense -UserId $newUser.Id @licenseParams
                        Write-Host "License '$friendlyName' assigned successfully." -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Failed to assign license: $($_.Exception.Message)" -ForegroundColor Red
                        if ($_.Exception.Message -like "*insufficient*" -or $_.Exception.Message -like "*quota*") {
                            Write-Host "This is likely due to insufficient license quantity. Please add licenses in Admin Center." -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "Invalid license selection." -ForegroundColor Red
                }
            }
        }
    }
    catch {
        Write-Host "Failed to create user: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Ask if the user wants to add another
    $addAnother = Read-Host "Add another user? (Y/N)"
} while ($addAnother -eq "Y" -or $addAnother -eq "y")

# Clean up
Write-Host "`nScript completed. Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
Disconnect-MgGraph
Write-Host "Disconnected successfully." -ForegroundColor Green