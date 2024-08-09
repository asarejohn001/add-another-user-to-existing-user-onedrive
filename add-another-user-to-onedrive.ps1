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

# Connect to SharePoint Online
try {
    Connect-SPOService -Url https://<Your-Tenant>-admin.sharepoint.com -ShowProgress $true -ErrorAction Stop
    Get-Log -LogFilePath $logFilePath -LogMessage "Successfully connected to SharePoint Online."
    
} catch {
    # Handle the error if connection fails
    Get-Log -LogFilePath $logFilePath -LogMessage "Failed to connect to SharePoint Online. Exiting script. Error details: $_"
    Write-Host "Couldn't connect to SharePoint Online. Check log file"
    exit
}

# Set log file path
$logFilePath = "./log.txt"

# Define the path to your CSV file
$csvFilePath = "C:\path\to\your\users.csv"

# Import the CSV file
$users = Import-Csv -Path $csvFilePath

# Get your username
$admin = "admin#example.com"

# Sign script progess
Write-Host "Script is in-progress..."

# Loop through each pair of CurrentUserID and DelegateUserID in the CSV file
foreach ($user in $users) {
    $currentUserID = $user.CurrentUserID
    $delegateUserID = $user.DelegateUserID
    
    # Construct the URL to the user's OneDrive site
    $oneDriveUrl = "https://<Your-Tenant>-my.sharepoint.com/personal/$($currentUserID -replace '@','_')"

    # Show which user processing
    Get-Log -LogFilePath $logFilePath -LogMessage "Processing OneDrive for: $currentUserID"

    try {
        # Add yourself to the SharePoint as admin before you can add others
        Set-SPOUser -Site $oneDriveUrl -LoginName $admin -IsSiteCollectionAdmin $true -ErrorAction Stop
        Get-Log -LogFilePath $logFilePath -LogMessage "Successfully added yourself as an admin to $currentUserID OneDrive"
    } catch {
        # Handle the error if adding yourself as admin fails
        Get-Log -LogFilePath $logFilePath -LogMessage "Failed to added yourself as an admin to $currentUserID OneDrive. Error details: $_"
        Write-Host "Failed to add, check log file"
        exit  # Exit the script because further operations depend on this action
    }

    try {
        # Grant permission to the delegate user
        Set-SPOUser -Site $oneDriveUrl -LoginName $delegateUserID -IsSiteCollectionAdmin $true -ErrorAction Stop
        Get-Log -LogFilePath $logFilePath -LogMessage "Successfully added yourself as an admin to $currentUserID OneDrive"
    } catch {
        # Handle the error if granting permission to the delegate user fails
        Get-Log -LogFilePath $logFilePath -LogMessage "Failed to add delegate: $delegateUserID to $currentUserID Onedrive. Error details: $_"
        Write-Host "Failed, check log file"
    }
}

# Disconnect from SharePoint Online
Disconnect-SPOService

Write-Host "Script execution completed."
