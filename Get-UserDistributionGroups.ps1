# ============================================================
#  Get-UserDistributionGroups.ps1
#  Looks up all distribution lists a user (or multiple users)
#  belongs to. Supports single entry or bulk lookup via
#  CSV / Excel. Uses DisableWAM auth — fully MFA compatible.
# ============================================================

# ── Install required modules if not already present ────────
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Cyan
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module (needed for .xlsx support)..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# ════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ════════════════════════════════════════════════════════════

function Validate-Email {
    param ([string]$Email)
    return $Email -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

function Get-GroupsForUser {
    param (
        [string]$Email,
        [array]$AllGroups
    )

    Write-Host "`nLooking up groups for: $Email" -ForegroundColor Cyan

    $matchedGroups = $AllGroups | Where-Object {
        (Get-DistributionGroupMember $_.Identity -ResultSize Unlimited |
            Select-Object -ExpandProperty PrimarySmtpAddress) -contains $Email
    }

    if ($matchedGroups) {
        $matchedGroups | Select-Object DisplayName, PrimarySmtpAddress | Format-Table -AutoSize
    }
    else {
        Write-Host "  '$Email' is not a member of any distribution lists." -ForegroundColor Yellow
    }

    return $matchedGroups
}

# ════════════════════════════════════════════════════════════
#  CONNECT TO EXCHANGE ONLINE
# ════════════════════════════════════════════════════════════

Write-Host "`n--- Admin Credentials ---" -ForegroundColor Yellow
$UserPrincipalName = Read-Host "Enter your admin email (e.g. admin@contoso.com)"

Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Write-Host "A browser window will open — sign in and complete MFA when prompted." -ForegroundColor Yellow
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -DisableWAM -ShowProgress $true

try {

    # ════════════════════════════════════════════════════════
    #  PROMPT: SINGLE OR BULK
    # ════════════════════════════════════════════════════════

    Write-Host "`n--- Lookup Mode ---" -ForegroundColor Yellow
    Write-Host "  [1] Look up a single user manually"
    Write-Host "  [2] Bulk look up users from a CSV or Excel file"
    do {
        $mode = Read-Host "Enter 1 or 2"
    } while ($mode -notin @("1", "2"))

    # Pre-fetch all distribution groups once for efficiency
    Write-Host "`nFetching all distribution groups (this may take a moment)..." -ForegroundColor Cyan
    $allGroups = Get-DistributionGroup -ResultSize Unlimited

    Write-Host "Found $($allGroups.Count) distribution group(s)." -ForegroundColor Green

    # ════════════════════════════════════════════════════════
    #  MODE 1 — SINGLE USER
    # ════════════════════════════════════════════════════════

    if ($mode -eq "1") {
        Write-Host "`n--- User Details ---" -ForegroundColor Yellow

        do {
            $UserEmail = Read-Host "Enter the user's email address"
            if (-not (Validate-Email $UserEmail)) {
                Write-Host "That doesn't look like a valid email address — please try again." -ForegroundColor Red
            }
        } while (-not (Validate-Email $UserEmail))

        Get-GroupsForUser -Email $UserEmail -AllGroups $allGroups
    }

    # ════════════════════════════════════════════════════════
    #  MODE 2 — BULK FROM CSV OR EXCEL
    # ════════════════════════════════════════════════════════

    elseif ($mode -eq "2") {
        Write-Host "`n--- Bulk Lookup ---" -ForegroundColor Yellow
        Write-Host "Your file must have a column named 'Email' (a 'Name' column is optional but recommended)." -ForegroundColor Cyan

        do {
            $filePath = Read-Host "Enter the full path to your CSV or Excel file (e.g. C:\Users\You\users.csv)"
            $filePath = $filePath.Trim('"')
            if (-not (Test-Path $filePath)) {
                Write-Host "File not found — please check the path and try again." -ForegroundColor Red
            }
        } while (-not (Test-Path $filePath))

        $ext = [System.IO.Path]::GetExtension($filePath).ToLower()

        try {
            if ($ext -eq ".csv") {
                $users = Import-Csv -Path $filePath
            }
            elseif ($ext -in @(".xlsx", ".xls")) {
                $users = Import-Excel -Path $filePath
            }
            else {
                Write-Host "Unsupported file type '$ext'. Please use .csv, .xlsx, or .xls." -ForegroundColor Red
                exit 1
            }
        }
        catch {
            Write-Host "ERROR reading file: $_" -ForegroundColor Red
            exit 1
        }

        # Validate required Email column exists
        $firstRow = $users | Select-Object -First 1
        if (-not ($firstRow.PSObject.Properties.Name -contains "Email")) {
            Write-Host "ERROR: File must contain an 'Email' column. Please check your file and try again." -ForegroundColor Red
            exit 1
        }

        $total   = ($users | Measure-Object).Count
        $counter = 0
        $results = @()

        Write-Host "`nProcessing $total user(s)..." -ForegroundColor Cyan

        foreach ($user in $users) {
            $counter++
            $email = $user.Email.Trim()
            $name  = if ($user.PSObject.Properties.Name -contains "Name") { $user.Name.Trim() } else { $email }

            Write-Host "`n[$counter/$total] $name <$email>" -ForegroundColor White

            if ([string]::IsNullOrWhiteSpace($email)) {
                Write-Host "  Skipping — missing email." -ForegroundColor Yellow
                continue
            }

            if (-not (Validate-Email $email)) {
                Write-Host "  Skipping — '$email' is not a valid email address." -ForegroundColor Yellow
                continue
            }

            $groups = Get-GroupsForUser -Email $email -AllGroups $allGroups

            # Collect results for summary
            foreach ($group in $groups) {
                $results += [PSCustomObject]@{
                    UserName         = $name
                    UserEmail        = $email
                    GroupDisplayName = $group.DisplayName
                    GroupEmail       = $group.PrimarySmtpAddress
                }
            }
        }

        # ── Summary ───────────────────────────────────────
        Write-Host "`n════════════════════════════════" -ForegroundColor Cyan
        Write-Host "  Bulk lookup complete" -ForegroundColor Cyan
        Write-Host "  Users looked up : $total"
        Write-Host "  Group memberships found : $($results.Count)" -ForegroundColor Green
        Write-Host "════════════════════════════════" -ForegroundColor Cyan

        # Offer to export results
        if ($results.Count -gt 0) {
            $export = Read-Host "`nExport results to CSV? (Y/N)"
            if ($export -eq "Y" -or $export -eq "y") {
                $exportPath = Read-Host "Enter export path (e.g. C:\Users\You\results.csv)"
                $exportPath = $exportPath.Trim('"')
                $results | Export-Csv -Path $exportPath -NoTypeInformation
                Write-Host "Results exported to '$exportPath'." -ForegroundColor Green
            }
        }
    }

    Write-Host "`nDone." -ForegroundColor Green

}
finally {
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
}
