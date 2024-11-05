Install-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name ImportExcel -Force

# Import necessary modules
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline

# Define the path to your CSV file
$CSVFilePath = "C:\Path\To\File.csv"

# Read the CSV file
Write-Host "Reading the CSV file..."
$users = Import-Csv -Path $CSVFilePath



# Function to add users to a distribution list
function Add-UserToDistributionList {
    param (
        [string]$UserEmailAddress,
        [string]$DistributionListName
    )

    # Check if the user is already a member of the distribution list
    $isMember = Get-DistributionGroupMember -Identity $DistributionListName -Member $UserEmailAddress -ErrorAction SilentlyContinue
    if (!$isMember) {
        # Add user to distribution list
        Add-DistributionGroupMember -Identity $DistributionListName -Member $UserEmailAddress
        Write-Host "User $($UserEmailAddress) added to $($DistributionListName)." -ForegroundColor Green
    } else {
        Write-Host "User $($UserEmailAddress) is already a member of $($DistributionListName), skipping..." -ForegroundColor Yellow
    }
}

# Distribution list name
$distributionList = "Your Distro List Name"

# Loop through users from the CSV and add them to the organization and distribution list
foreach ($user in $users) {
    $userName = $user.Name
    $userEmail = $user."Email"

    Write-Host "Processing user: $($userName)"

   

    # Add the user to the distribution list
    Add-UserToDistributionList -UserEmailAddress $userEmail -DistributionListName $distributionList

    Write-Host "$($userName) processed successfully." -ForegroundColor Cyan
}

# Disconnect from Exchange Online when done
Disconnect-ExchangeOnline -Confirm:$false