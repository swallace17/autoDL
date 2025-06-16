# =========================
# Verify required modules
# =========================

Write-Output "Checking for required modules:"

Get-Module -ListAvailable | Where-Object { $_.Name -like "Microsoft.Graph*" } | Select Name, Version, Path
Write-Output "`n"

$requiredModules = @("Microsoft.Graph", "ExchangeOnlineManagement")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        throw "❌Required module '$module' is not available in the current environment."
    } else {
        Write-Output "✅Module '$module' is available."
    }
}

# =========================
# Connect to services
# =========================
Write-Output "`n--- Starting Security Group to Distribution List Sync ---`n"

if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
    throw "Connect-MgGraph is still not recognized. Module may be incomplete or not properly imported."
}
Connect-MgGraph -Identity
Write-Output "✅ Connected to Microsoft Graph using Managed Identity."

Connect-ExchangeOnline -ManagedIdentity -Organization "<your-tenant>.onmicrosoft.com"
Write-Output "✅ Connected to Exchange Online using Managed Identity."

# =========================
# Load Security Group list from Automation Variable
# =========================
$csvRaw = Get-AutomationVariable -Name "SecurityGroupCsv"
if (-not $csvRaw) {
    throw "❌ Automation variable 'SecurityGroupCsv' is empty or not found."
}

try {
    $groupMappings = $csvRaw | ConvertFrom-Csv
    Write-Output "📄 Loaded $($groupMappings.Count) security group(s) from automation variable."
} catch {
    throw "❌ Failed to parse CSV from 'SecurityGroupCsv': $($_.Exception.Message)"
}

# =========================
# Constants
# =========================
$dlSuffix = "-dl"
$emailDomain = "yourdomain.tld"
$indent = "    "

# =========================
# Loop through each group
# =========================
foreach ($mapping in $groupMappings) {
    $securityGroupName = $mapping.SecurityGroup
    # Set core naming values FIRST
    $dlDisplayName = "$securityGroupName - autoDL"
    $dlAlias = ($securityGroupName + "-autodl").ToLower()
    $distributionGroupEmail = "$securityGroupName@$emailDomain"

    Write-Output "`n🔁 Processing: $securityGroupName → $distributionGroupEmail"

    # --- Lookup Security Group ---
    $securityGroup = Get-MgGroup -Filter "displayName eq '$securityGroupName'" -ConsistencyLevel eventual
    if (-not $securityGroup.Id) {
        Write-Warning "❌ Security group '$securityGroupName' not found. Skipping."
        continue
    }
    Write-Output "🔍 Security group retrieved: $($securityGroup.DisplayName), ID: $($securityGroup.Id)"

    # --- Create DL if it doesn't exist ---
    $dl = Get-DistributionGroup -Identity $distributionGroupEmail -ErrorAction SilentlyContinue
    if (-not $dl) {
        Write-Output "${indent}🆕 Distribution list not found. Creating: $distributionGroupEmail"
        try {
            $dlDisplayName = "$securityGroupName - autoDL"
            $dlAlias = ($securityGroupName + "-autodl").ToLower()  # ensure valid characters
            $distributionGroupEmail = "$securityGroupName@$emailDomain"

            New-DistributionGroup `
                -Name $dlDisplayName `
                -DisplayName $dlDisplayName `
                -Alias $dlAlias `
                -PrimarySmtpAddress $distributionGroupEmail | Out-Null

            Write-Output "${indent}✅ DL created: $distributionGroupEmail"
        } catch {
            Write-Warning "❌ Failed to create DL: $($_.Exception.Message). Skipping group."
            continue
        }
    } else {
        Write-Output "${indent}🔎 Found existing DL: $distributionGroupEmail"
    }

    # --- Get Security Group Members ---
    $rawGroupMembers = Get-MgGroupMember -GroupId $securityGroup.Id -All
    Write-Output "${indent}👥 Security Group members fetched: $($rawGroupMembers.Count)"

    $securityUPNs = $rawGroupMembers | ForEach-Object {
        if ($_.AdditionalProperties -and $_.AdditionalProperties.userPrincipalName) {
            $_.AdditionalProperties.userPrincipalName.ToLower()
        } else {
            Write-Warning "⚠️ Missing UPN for member ID: $($_.Id)"
        }
    } | Where-Object { $_ }

    Write-Output "${indent}📋 Parsed UPNs from security group: $($securityUPNs.Count) users"
    Write-Output "${indent}🧾 Members from Security Group:"
    $rawGroupMembers | Where-Object {
        $_.AdditionalProperties -and $_.AdditionalProperties.userPrincipalName
    } | ForEach-Object {
        $name = $_.AdditionalProperties.displayName
        $upn  = $_.AdditionalProperties.userPrincipalName
        Write-Output "${indent}    👤 $name <$upn>"
    }

    # --- Get Distribution Group Members ---
    $distributionGroupMembers = @(Get-DistributionGroupMember -Identity $distributionGroupEmail -ResultSize Unlimited |
        Where-Object { $_.RecipientType -eq "UserMailbox" -or $_.RecipientType -eq "MailUser" })

    Write-Output "${indent}👥 Distribution Group members fetched: $($distributionGroupMembers.Count)"

    $distributionUPNs = $distributionGroupMembers | ForEach-Object {
        $_.PrimarySmtpAddress.ToLower()
    }

    Write-Output "${indent}📋 Parsed UPNs from distribution group: $($distributionUPNs.Count) users"
    Write-Output "${indent}🧾 Members from Distribution Group:"
    $distributionGroupMembers | ForEach-Object {
        $name = $_.Name
        $upn  = $_.PrimarySmtpAddress
        Write-Output "${indent}    👤 $name <$upn>"
    }

    # --- Sync Members ---
    $usersToAdd = $securityUPNs | Where-Object { $_ -notin $distributionUPNs }
    $usersToRemove = $distributionUPNs | Where-Object { $_ -notin $securityUPNs }

    if ($usersToAdd.Count -eq 0 -and $usersToRemove.Count -eq 0) {
        Write-Output "${indent}🟢 No changes required — memberships are already in sync."
    } else {
        foreach ($upn in $usersToAdd) {
            try {
                Write-Output "${indent}➕ Adding $upn to $distributionGroupEmail"
                Add-DistributionGroupMember -Identity $distributionGroupEmail -Member $upn
            } catch {
                Write-Warning "❌ Failed to add ${upn}: $($_.Exception.Message)"
            }
        }

        foreach ($upn in $usersToRemove) {
            try {
                Write-Output "${indent}➖ Removing $upn from $distributionGroupEmail"
                Remove-DistributionGroupMember -Identity $distributionGroupEmail -Member $upn -Confirm:$false
            } catch {
                Write-Warning "❌ Failed to remove ${upn}: $($_.Exception.Message)"
            }
        }
    }

    Write-Output "${indent}✅ Sync complete: '$securityGroupName' → '$distributionGroupEmail'"
}
