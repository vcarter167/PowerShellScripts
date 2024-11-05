#Install module and Connect to Entra ID Tenant
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline

# Import the CSV file
$contacts = Import-Csv -Path "C:\Path\To\File.csv"

# Function to add contacts to the organization
function Add-OrganizationContact {
    param (
        [string]$ContactName,
        [string]$ExternalEmailAddress
    )

    # Check if the contact already exists
    $existingContact = Get-MailContact -Identity $ContactName -ErrorAction SilentlyContinue
    if (!$existingContact) {
        # Create a new mail contact
        New-MailContact -Name $ContactName -ExternalEmailAddress $ExternalEmailAddress -FirstName $ContactName.Split(" ")[0] -LastName $ContactName.Split(" ")[1]
    } else {
        Write-Host "Contact $($ContactName) already exists, skipping..."
    }
}

# Function to add contacts to a distribution list
function Add-ContactToDistributionList {
    param (
        [string]$ContactEmailAddress,
        [string]$DistributionListName
    )

    # Check if the contact is already a member of the distribution list
    $isMember = Get-DistributionGroupMember -Identity $DistributionListName -Member $ContactEmailAddress -ErrorAction SilentlyContinue
    if (!$isMember) {
        # Add contact to distribution list
        Add-DistributionGroupMember -Identity $DistributionListName -Member $ContactEmailAddress
    } else {
        Write-Host "Contact $($ContactEmailAddress) is already a member of $($DistributionListName), skipping..."
    }
}

# Distribution list name
$distributionList = "Your Distro List Name"

# Loop through contacts and add them to the organization and distribution list
foreach ($contact in $contacts) {
    Write-Host "Adding contact: $($contact.Name)"
    
    # Add the contact to the organization
    Add-OrganizationContact -ContactName $contact.Name -ExternalEmailAddress $contact.Email

    # Add the contact to the distribution list
    Add-ContactToDistributionList -ContactEmailAddress $contact.Email -DistributionListName $distributionList

    Write-Host "$($contact.Name) added successfully."
}

# Disconnect from Exchange Online when done
Disconnect-ExchangeOnline -Confirm:$false