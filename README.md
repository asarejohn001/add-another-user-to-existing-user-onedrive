# add-another-user-to-existing-user-onedrive

## Scenario
You want to grant project managers access to their coordinator's OneDrive after the end of the project.

## Solution
Using the [Connect-SPOService](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/connect-sposervice?view=sharepoint-ps), we can perform this action on a bulk process. The [Set-SPOUser](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/set-spouser?view=sharepoint-ps) grants the Delegate user permission on the CurrentUser user's OneDrive site. This gives the delegate full administrative access to the OneDrive.

## Script explained
The [add-another-user-to-onedrive.ps1](./add-another-user-to-onedrive.ps1) will
1. Install the [Microsoft.Online.SharePoint.PowerShell](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/?view=sharepoint-ps) module if not already installed on your computer. 
2. Set a function to log the script runtime success and error messages.
3. Connect to SharePoint Online.
> [!IMPORTANT]
> Change the SharePoint URL variable to your tenant's URL
4. Set variable for log file, CSV file, and for the admin user.
> [!IMPORTANT]
> Modify the variables to match yours
5. 