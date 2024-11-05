Set-Executionpolicy bypass
# Import the Entra ID module
Import-Module -Name AzureAD

# Connect to Entra ID AD
Connect-AzureAD

# Load the user data from the CSV file
$userData = Import-Csv -Path "C:\Path\To\File.csv"

# Edit the user profiles
foreach ($user in $userData) {
    $upn = $user.'Object Id'
    $jobTitle = $user.Title
    $departments = $user.Department
    $officeLocation = $user."Office"
    $displayName = $user.'Display name'
    
    # Convert department code to plain text
    $department = $departments.substring(3)
  
    # Edit the user's job title, department, and office
    Set-AzureADUser -ObjectId $upn -JobTitle $jobTitle
    Set-AzureADUser -ObjectId $upn -Department $department
    Set-AzureADUser -ObjectId $upn -PhysicalDeliveryOfficeName $officeLocation
    
   
    Write-Host "User profile updated successfully for $displayName who is now a $jobTitle!"
}
Write-Host "All user profiles updated successfully!"
