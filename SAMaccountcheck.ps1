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

# Function to get user information by UPN
function Get-UserInfoByUPN {
    param (
        [string]$UserUPN
    )

    # Retrieve user information from Microsoft Graph
    $user = Get-MgUser -UserId $UserUPN -Property DisplayName, UserPrincipalName, OnPremisesSamAccountName, Id, Department, Mail

    if ($user) {
        # Display user information
        [PSCustomObject]@{
            DisplayName         = $user.DisplayName
            UPN                 = $user.UserPrincipalName
            'SAM Account'       = $user.OnPremisesSamAccountName
            ObjectID            = $user.Id
            Department          = $user.Department
            Email               = $user.Mail
        }
    } else {
        Write-Host "User not found or unable to retrieve user information."
    }
}

# Main script execution

# Check and install Microsoft.Graph module if needed
Install-MicrosoftGraphModule

# Connect to Microsoft Graph
Connect-MicrosoftGraph

# Prompt for UPN input
$userUPN = Read-Host "Enter the User Principal Name (UPN) of the user"

# Get and display user information
$userInfo = Get-UserInfoByUPN -UserUPN $userUPN
if ($userInfo) {
    $userInfo | Format-List
}

# Wait for user to press any key before closing
Write-Host "Press any key to exit..."
[void][System.Console]::ReadKey($true)
