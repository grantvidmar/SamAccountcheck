# Function to check and install Microsoft.Graph module
function Install-MicrosoftGraphModule {
    $moduleName = 'Microsoft.Graph'

    # Check if the module is already installed
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "The $moduleName module is not installed. Installing..."
        Install-Module -Name $moduleName -Scope CurrentUser -Force
    } else {
        Write-Host "$moduleName module is already installed."
    }
}

# Function to connect to Microsoft Graph
function Connect-MicrosoftGraph {
    # Connect to Microsoft Graph
    Connect-MgGraph
}

# Function to update CSV with user information
function Update-CSVWithUserInfo {
    param (
        [string]$csvPath
    )

    # Import the CSV file
    $csvData = Import-Csv -Path $csvPath

    # Loop through each row in the CSV
    foreach ($row in $csvData) {
        $userUPN = $row.UPN
        
        # Check if UPN is empty
        if ([string]::IsNullOrWhiteSpace($userUPN)) {
            Write-Host "UPN is empty for one of the rows. Skipping..."
            continue
        }
        
        # Retrieve user information from Microsoft Graph
        try {
            $user = Get-MgUser -UserId $userUPN -Property DisplayName, UserPrincipalName, OnPremisesSamAccountName, Id, Department, Mail
            if ($user) {
                # Update the row with retrieved user information
                $row.DisplayName = $user.DisplayName
                $row.'SAM Account' = $user.OnPremisesSamAccountName
                $row.ObjectID = $user.Id
                $row.Department = $user.Department
                $row.Email = $user.Mail
            } else {
                Write-Host "User not found or unable to retrieve information for UPN: $userUPN"
            }
        } catch {
            Write-Host "Error retrieving user information for UPN: $userUPN. Error: $_"
        }
    }

    # Export updated data back to CSV
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "CSV file updated successfully."
}

# Main script execution

# Check and install Microsoft.Graph module if needed
Install-MicrosoftGraphModule

# Connect to Microsoft Graph
Connect-MicrosoftGraph

# Define the path to the CSV file
$csvFilePath = "C:/temp/NewEveryfileUsers.csv"

# Update CSV with user information
Update-CSVWithUserInfo -csvPath $csvFilePath

# Wait for user to press any key before closing
Write-Host "Press any key to exit..."
[void][System.Console]::ReadKey($true)
