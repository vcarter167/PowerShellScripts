# ============================================================
#  Add-UsersToSharedMailbox.ps1
#  Grants Full Access to a shared mailbox for one or more
#  users. Supports single entry or bulk import via CSV/Excel.
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

function Validate-Email {
    param ([string]$Email)
    return $Email -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

function Add-UserToSharedMailbox {
    param (
        [string]$UserName,
        [string]$UserEmail,
        [string]$SharedMailbox
    )

    try {
        Add-MailboxPermission `
            -Identity        $SharedMailbox `
            -User            $UserEmail `
            -AccessRights    FullAccess `
            -InheritanceType All `
            -ErrorAction     Stop
        Write-Host "  '$UserName' granted Full Access to '$SharedMailbox'." -ForegroundColor Green
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*parameter cannot be found that matches parameter name 'Member'*" -or
            $errorMessage -like "*Cannot process argument transformation on parameter 'Member'*") {
            # Suppress known benign Exchange errors
            return
        }
        Write-Host "  ERROR adding '$UserName': $errorMessage" -ForegroundColor Red
    }
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
    #  PROMPT: SHARED MAILBOX
    # ════════════════════════════════════════════════════════

    Write-Host "`n--- Shared Mailbox ---" -ForegroundColor Yellow

    do {
        $SharedMailbox = Read-Host "Enter the shared mailbox email address (e.g. sharedbox@contoso.com)"
        if (-not (Validate-Email $SharedMailbox)) {
            Write-Host "That doesn't look like a valid email address — please try again." -ForegroundColor Red
        }
    } while (-not (Validate-Email $SharedMailbox))

    # Validate shared mailbox exists
    Write-Host "`nValidating shared mailbox..." -ForegroundColor Cyan
    $mbx = Get-Mailbox -Identity $SharedMailbox -ErrorAction SilentlyContinue
    if (-not $mbx) {
        Write-Host "ERROR: Shared mailbox '$SharedMailbox' was not found. Check the email and try again." -ForegroundColor Red
        exit 1
    }
    Write-Host "Found: $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))" -ForegroundColor Green

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

        Add-UserToSharedMailbox `
            -UserName      $UserName `
            -UserEmail     $UserEmail `
            -SharedMailbox $SharedMailbox
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
                Add-UserToSharedMailbox `
                    -UserName      $name `
                    -UserEmail     $email `
                    -SharedMailbox $SharedMailbox
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
