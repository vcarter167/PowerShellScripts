# ============================================================
#  Add-UsersToDistributionList.ps1
#  Adds existing internal users to a distribution list.
#  Supports single entry or bulk import via CSV / Excel.
#  Uses DisableWAM auth — fully MFA compatible.
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

function Add-UserToDistributionList {
    param (
        [string]$UserName,
        [string]$UserEmailAddress,
        [string]$DistributionListName,
        [string]$DLDisplayName,
        [hashtable]$MemberTable
    )

    if (-not $MemberTable.ContainsKey($UserEmailAddress.ToLower())) {
        try {
            Add-DistributionGroupMember -Identity $DistributionListName -Member $UserEmailAddress
            Write-Host "  '$UserName' added to '$DLDisplayName'." -ForegroundColor Green
        }
        catch {
            Write-Host "  ERROR adding '$UserName': $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "  '$UserName' is already a member of '$DLDisplayName' — skipping." -ForegroundColor Yellow
    }
}

function Validate-Email {
    param ([string]$Email)
    return $Email -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
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

    Write-Host "`n--- Add Mode ---" -ForegroundColor Yellow
    Write-Host "  [1] Add a single user manually"
    Write-Host "  [2] Bulk add users from a CSV or Excel file"
    do {
        $mode = Read-Host "Enter 1 or 2"
    } while ($mode -notin @("1", "2"))

    # ════════════════════════════════════════════════════════
    #  PROMPT: DISTRIBUTION LIST (shared for both modes)
    # ════════════════════════════════════════════════════════

    Write-Host "`n--- Distribution List ---" -ForegroundColor Yellow
    $DistributionList = Read-Host "Enter the distribution list name or email address"

    Write-Host "`nValidating distribution list..." -ForegroundColor Cyan
    $dl = Get-DistributionGroup -Identity $DistributionList -ErrorAction SilentlyContinue
    if (-not $dl) {
        Write-Host "ERROR: Distribution list '$DistributionList' not found. Check the name or email and try again." -ForegroundColor Red
        exit 1
    }
    Write-Host "Found: $($dl.DisplayName) ($($dl.PrimarySmtpAddress))" -ForegroundColor Green

    # Pre-fetch DL members once for efficient lookup
    Write-Host "Fetching current DL membership..." -ForegroundColor Cyan
    $memberTable = @{}
    Get-DistributionGroupMember -Identity $DistributionList -ResultSize Unlimited |
        ForEach-Object { $memberTable[$_.PrimarySmtpAddress.ToLower()] = $true }

    # ════════════════════════════════════════════════════════
    #  MODE 1 — SINGLE USER
    # ════════════════════════════════════════════════════════

    if ($mode -eq "1") {
        Write-Host "`n--- User Details ---" -ForegroundColor Yellow

        do {
            $UserName = Read-Host "Enter the user's full name"
        } while ([string]::IsNullOrWhiteSpace($UserName))

        do {
            $UserEmail = Read-Host "Enter the user's email address"
            if (-not (Validate-Email $UserEmail)) {
                Write-Host "That doesn't look like a valid email address — please try again." -ForegroundColor Red
            }
        } while (-not (Validate-Email $UserEmail))

        Add-UserToDistributionList `
            -UserName             $UserName `
            -UserEmailAddress     $UserEmail `
            -DistributionListName $DistributionList `
            -DLDisplayName        $dl.DisplayName `
            -MemberTable          $memberTable
    }

    # ════════════════════════════════════════════════════════
    #  MODE 2 — BULK FROM CSV OR EXCEL
    # ════════════════════════════════════════════════════════

    elseif ($mode -eq "2") {
        Write-Host "`n--- Bulk Import ---" -ForegroundColor Yellow
        Write-Host "Your file must have columns named 'Name' and 'Email'." -ForegroundColor Cyan

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

        # Validate required columns exist
        $firstRow = $users | Select-Object -First 1
        if (-not ($firstRow.PSObject.Properties.Name -contains "Name") -or
            -not ($firstRow.PSObject.Properties.Name -contains "Email")) {
            Write-Host "ERROR: File must contain 'Name' and 'Email' columns. Please check your file and try again." -ForegroundColor Red
            exit 1
        }

        $total   = ($users | Measure-Object).Count
        $success = 0
        $skipped = 0
        $failed  = 0
        $counter = 0

        Write-Host "`nProcessing $total user(s)..." -ForegroundColor Cyan

        foreach ($user in $users) {
            $counter++
            $name  = $user.Name.Trim()
            $email = $user.Email.Trim()

            Write-Host "`n[$counter/$total] $name <$email>" -ForegroundColor White

            if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($email)) {
                Write-Host "  Skipping — missing name or email." -ForegroundColor Yellow
                $skipped++
                continue
            }

            if (-not (Validate-Email $email)) {
                Write-Host "  Skipping — '$email' is not a valid email address." -ForegroundColor Yellow
                $skipped++
                continue
            }

            try {
                Add-UserToDistributionList `
                    -UserName             $name `
                    -UserEmailAddress     $email `
                    -DistributionListName $DistributionList `
                    -DLDisplayName        $dl.DisplayName `
                    -MemberTable          $memberTable
                $success++
            }
            catch {
                Write-Host "  ERROR processing '$name': $_" -ForegroundColor Red
                $failed++
            }
        }

        # ── Summary ───────────────────────────────────────
        Write-Host "`n════════════════════════════════" -ForegroundColor Cyan
        Write-Host "  Bulk import complete" -ForegroundColor Cyan
        Write-Host "  Total rows : $total"
        Write-Host "  Processed  : $success" -ForegroundColor Green
        Write-Host "  Skipped    : $skipped" -ForegroundColor Yellow
        Write-Host "  Failed     : $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
        Write-Host "════════════════════════════════" -ForegroundColor Cyan
    }

    Write-Host "`nDone." -ForegroundColor Green

}
finally {
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
}
