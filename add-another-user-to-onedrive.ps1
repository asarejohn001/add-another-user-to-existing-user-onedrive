<#
Author: John Asare
Date: 8/8/2024

Description: Manage OneDrive. Read more at https://github.com/asarejohn001/add-another-user-to-existing-user-onedrive?tab=readme-ov-file
#>

# Install SharePoint module - if you have not already installed
Install-Module -Name Microsoft.Online.SharePoint.PowerShell

# Import the SharePoint Online Management Shell module
Import-Module Microsoft.Online.SharePoint.PowerShell

function Get-Log {
    param (
        [string]$LogFilePath,
        [string]$LogMessage
    )

    # Create the log entry with the current date and time
    $logEntry = "{0} - {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $LogMessage

    # Append the log entry to the log file
    Add-Content -Path $LogFilePath -Value $logEntry
}

# Connect to EXO
try {
    # Connect to SharePoint Online
    Connect-SPOService -Url https://<Your-Tenant>-admin.sharepoint.com -ShowProgress $true -ErrorAction Stop
    Get-Log -LogFilePath $logFilePath -LogMessage "Successfully connected to Exchange Online."
    
} catch {
    # Handle the error if connection fails
    Get-Log -LogFilePath $logFilePath -LogMessage "Failed to connect to Exchange Online. Exiting script. Error details: $_"
    Write-Host "Couldn't connect to EXO. Check log file"
    exit
}

# Set log file path
$logFilePath = "./log.txt"

# Define the path to your CSV file
$csvFilePath = "C:\path\to\your\users.csv"

# Import the CSV file
$users = Import-Csv -Path $csvFilePath

# Loop through each pair of CurrentUserID and DelegateUserID in the CSV file
foreach ($user in $users) {
    $currentUserID = $user.CurrentUserID
    $delegateUserID = $user.DelegateUserID
    
    # Construct the URL to the user's OneDrive site
    $oneDriveUrl = "https://<Your-Tenant>-my.sharepoint.com/personal/$($currentUserID -replace '@','_')"

    # Display the current processing information
    Write-Host "Adding delegate: $delegateUserID to OneDrive of: $currentUserID"

    # Grant permission to the delegate user
    Set-SPOUser -Site $oneDriveUrl -LoginName $delegateUserID -IsSiteCollectionAdmin $true
}

# Disconnect from SharePoint Online
Disconnect-SPOService

Write-Host "Delegates have been added to OneDrive successfully."
