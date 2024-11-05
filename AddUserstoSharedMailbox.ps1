# Import the required .NET assemblies for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create an OpenFileDialog to allow the admin to upload the CSV file
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop') # Initial directory set to Desktop
$OpenFileDialog.Filter = "CSV files (*.csv)|*.csv" # Filter for CSV files
$OpenFileDialog.Title = "Select the CSV file to upload"

# Show the OpenFileDialog and get the selected file
if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVFilePath = $OpenFileDialog.FileName
    Write-Host "Selected CSV file: $CSVFilePath" -ForegroundColor Cyan
} else {
    Write-Host "No CSV file selected. Exiting..." -ForegroundColor Yellow
    exit
}

# Prompt the admin to input the shared mailbox email
$inputBox = New-Object System.Windows.Forms.Form
$inputBox.Text = "Input Shared Mailbox Email"
$inputBox.Width = 1300
$inputBox.Height = 1150

$inputTextBox = New-Object System.Windows.Forms.TextBox
$inputTextBox.Width = 200
$inputTextBox.Top = 20
$inputTextBox.Left = 40
$inputBox.Controls.Add($inputTextBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Top = 60
$okButton.Left = 100
$okButton.Add_Click({
    $inputBox.Close()
})
$inputBox.Controls.Add($okButton)

$inputBox.ShowDialog()
$sharedMailbox = $inputTextBox.Text

# Validate the input for shared mailbox email
if (-not $sharedMailbox -or $sharedMailbox -eq "") {
    Write-Host "No shared mailbox email provided. Exiting..." -ForegroundColor Yellow
    exit
}

Write-Host "Shared mailbox email set to: $sharedMailbox" -ForegroundColor Cyan

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline

# Read the CSV file
Write-Host "Reading the CSV file..."
$users = Import-Csv -Path $CSVFilePath

# Loop through each user in the CSV file and add them to the shared mailbox
foreach ($user in $users) {
    $userEmail = $user."Email"

    Write-Host "Adding user $userEmail to shared mailbox $sharedMailbox..."

    try {
        # Grant Full Access permission to the shared mailbox
        Add-MailboxPermission -Identity $sharedMailbox -User $userEmail -AccessRights FullAccess -InheritanceType All -ErrorAction Stop

        Write-Host "User $userEmail added successfully with Full Access to $sharedMailbox." -ForegroundColor Green

    } catch {
        $errorMessage = $_.Exception.Message

        # Silently suppress specific errors related to the 'Member' parameter
        if ($errorMessage -like "*parameter cannot be found that matches parameter name 'Member'*" -or 
            $errorMessage -like "*Cannot process argument transformation on parameter 'Member'*") {
            continue
        }

        # Display other errors
        Write-Host "Failed to add user $userEmail to $sharedMailbox $errorMessage" -ForegroundColor Red
    }
}

# Disconnect from Exchange Online when done
Write-Host "Disconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false
