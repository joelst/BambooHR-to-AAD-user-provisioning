# BambooHR to Entra ID User Provisioning

User provisioning from BambooHR to Microsoft Entra ID (formerly Azure Active Directory) without a requirement for SCIM or additional licenses.

## Introduction

This is a fork of the AMAZING work done by PaulTony BHR-to-AzAd-user-Provisioning (which has been removed from GitHub). Their work saved me many hours trying to put this together and I'm very grateful!

What is different about this from the original?

1. It is focused to work for one of my projects while trying to make it generic enought to work for other scenarios. Primarily the changes are to make this work for this project.
2. Use the [Microsoft Graph 2.x PowerShell modules](https://devblogs.microsoft.com/microsoft365dev/microsoft-graph-powershell-v2-is-now-in-public-preview-half-the-size-and-will-speed-up-your-automations/). There are some differences that forced changes.
3. Make it run as an Azure Automation runbook. Added parameters to allow for easy customization without modifying the code, making it easier to execute it in Azure Automation.
4. My project is for a cloud-first Entra ID organization that but may also work with some hybrid user objects.
5. I've got a _little OCD_ when it comes to silly things like variable names. Although I _really_ tried to fight my irrational desire to change them, I failed. No, there wasn't anything wrong with the original ones, just try explaining that to my OCD.
6. If you use this, please star the repo or reach out to me.

## Changes & Updates

This section will keep track of changes made over time.

### TODO

- Create an Azure Bicep or ARM template for easy deployment.

### December 2025 changes

- Added details to the Teams chat message about the changes made during the script execution.
- Added `DaysToKeepAccountsAfterTermination` parameter. When a user has been disabled past this set days, a notification will be sent that the account needs to be deleted. Also now update the EmployeeLeaveDateTime in Entra ID to make this work better.

> Note: In the future, I will probably add an option to automatically delete the account after this time. Let me know if you want this now!

- Remove user location information on termination to ensure the accounts gets removed from all groups.

### October 2025 changes

- Converted custom `-TestOnly` switch to PowerShell's standard `-WhatIf` parameter using `SupportsShouldProcess` pattern for better compliance with PowerShell standards.
- Set `ConfirmImpact = 'None'` to prevent script hanging in unattended Azure Automation environments while preserving `-WhatIf` functionality.
- Added connection state tracking to prevent redundant connections to Azure services.
- Created `Connect-ExchangeOnlineIfNeeded` helper function to optimize Exchange connections.
- Restructured Teams adaptive card logic to always send summary notifications regardless of execution mode (production, WhatIf, or no changes).
- Fixed blank duration display in completion logs by properly calculating script execution time.
- Added automatic loading of email addresses (`AdminEmailAddress`, `NotificationEmailAddress`, `HelpDeskEmailAddress`, `LicenseId`) from Azure Automation variables with fallback to generated defaults.
- Added many comments and notes both in the script and a separate developer guide to help understand and customize this solution.

### May 2025 changes

- When an employee is offboarded, the following tasks are now completed.

  - The user's owned Windows devices are Autopilot reset.
  - The user is removed as owner of devices and a note is added.
  - All of the meetings owned by the user are canceled.
  - The mailbox has an out of office set.
  - All owned groups are transferred to the user's former manager.

### February 2025 changes

- Reconfigured the script to run in Azure Automation using variables for most configuration. Create the following variables:

  - BambooHrApiKey
  - BhrCompanyName
  - CompanyName
  - TeamsCardUri
  - TenantId

- Removed certificate and application registrations authentication and added using a system managed identity.
- Review [Add-ManagedIdentityPermissions.ps1](./Add-ManagedIdentityPermissions.ps1) to get started assigning the Automation account appropriate permissions.

### January 2025 changes

- Added mailbox delegation based on group membership with an array with the group and mailboxes.
- Added license tracking. This will report to make sure there are enough licenses and send an email and teams message if it breaches the defined free licenses.
- Update to new Teams webhooks.
- Requires Microsoft.Graph.Identity.DirectoryManagement
- A welcome message to new users. The formatting is ugly, but you can customize it.
- Offboarding: Removed mailbox is set to Shared and the manager is given permissions to it. You will need to manually remove these later.
- Re-onboarding: If user is disabled and then re-enabled it will undo the shared mailbox permissions.
- Minor text changes and error reporting changes

### November 2023 changes

- Improvements to make running in Azure Automation easier.
- Minor documentation and verbiage changes to messages.
- Clean up of no longer needed code.
- Sprinkle in some Entra Id renaming into user text.
- Now removes the Entra Id Manager when user is terminated.

### July 2023 changes

- Fixed errors I introduced around name changes.
- If an employee is not active and their account is disabled don't keep synchronizing their department and other info. Also, set the department to "" when they are deactivated.
- Added days ahead parameter to set the number of days to look ahead for new hires.
- Changed email format to html to allow for improved formatting and clickable URLs.
- When a user is terminated, the former manager is now assigned permissions to user's shared mailbox.

### March 2023 changes

- Created parameters for common information needed.
- Updated password creation with a function that avoids characters that are difficult to differentiate (like 0,O and 1,l,I) and avoids any characters that might cause PowerShell issues (like $ and `)
- The API key from BambooHR is automatically formatted, you can simply provide the key as-is from BambooHR
- Microsoft Graph 2.0-preview8 PowerShell module tested.
- Removed the line numbers from the log messages, because I messed up the line numbers. The stack trace should provide enough info.
- Added `-TestOnly` parameter as a pseudo `-WhatIf` parameter. It will log what would have been executed but will not make any changes. *(Note: This has been replaced with standard PowerShell `-WhatIf` in October 2025)*
- Changed the Bamboo report to pull future employees, as this was required for my project. Added the `-CurrentOnly` parameter if you would rather only process current employees.
- Moved most screen output to Verbose channel. If you are troubleshooting, run with -Verbose to see details.
- Fixed a bug in logging where the wrong time was being logged. You can see the fix in the incorrectly named [Original_BambooHr_User_Provisioning.ps1](./Original_BambooHR_User_Provisioning.ps1) as it has a couple small changes from the [upstream version](https://github.com/PaulTony-coocoo/BHR-to-AzAD-user-provisioning/blob/main/GENERALIZED_AUTO_USER_PROVISIONING.ps1).
- Added NotificationEmailAddress to copy all messages to so that HR or IT can keep an eye what changed process. This will give them the information needed to track down when a user's information changed.
- Added Sync-GroupMailboxDelegation function to set shared mailbox permissions.
- Added Photo Sync, when a new user is created it will attempt to add the photo from BambooHR to Entra ID. This is not kept in sync afterward, just on initial creation.
- Added Teams Adaptive Card logging. This is very simple and ugly now, but moved from sending an email to the webhook. This should get better in the future.

## Known issues

- ExtensionAttribute1 with the LastUpdate from BHR no longer works.
- User off boarding process may have a couple issues that needs to be retested.
- There are a number of areas where the process can be streamlined and _may be_ addressed in the future.
- Entra Id accounts are disabled and not deleted. You will need to either add logic to automatically delete accounts or do this manually.

## Testing and Validation

**IMPORTANT**: Always test thoroughly before production use:

1. **Preview Mode**: Run with `-WhatIf` parameter to preview changes without applying them:
   ```powershell
   .\Start-BambooHRUserProvisioning.ps1 -WhatIf
   ```

2. **Validate Configuration**: Ensure all Azure Automation variables are properly configured.

3. **Test with Subset**: Start with a small group of users before full deployment.

4. **Monitor Logs**: Review both PowerShell output and Teams notifications for any issues.

5. **Azure Automation**: The script is optimized for unattended execution in Azure Automation runbooks with proper error handling and connection state management.

## Future ideas

If you have suggestions or questions, feel free to reach out.

- Add off boarding steps for things like:
  - One Drive for Business data access
  - Transfer group ownership
  - Deleting disabled accounts after suitable waiting period.
  - Other suggestions

## Getting started

This is part of BambooHR and Entra ID integration process. This will ensure employees have Entra ID accounts. This does not create or update on-premises Active Directory Domain Services accounts.

### Prerequisites

1. **Entra ID Application Registration**: Create an Entra ID Enterprise application with managed identity for unattended authentication
   - Required Graph API permissions: `User.ReadWrite.All`, `Directory.ReadWrite.All`, `Mail.Send`
   - Review [Add-ManagedIdentityPermissions.ps1](./Add-ManagedIdentityPermissions.ps1) for permission setup

2. **Azure Automation Variables**: Configure the following variables in your Azure Automation account:

- **BambooHrApiKey** - API key created in BambooHR (secure string)
- **BHRCompanyName** - Company name used for the URL to access BambooHR  
- **CompanyName** - Your company display name
- **TenantId** - Microsoft Tenant ID
- **TeamsCardUri** - (Optional) Teams webhook URL for notifications
- **AdminEmailAddress** - (Optional) Administrator email for notifications (falls back to generated default)
- **NotificationEmailAddress** - (Optional) HR notification email address (falls back to generated default)
- **HelpDeskEmailAddress** - (Optional) Help desk email for user support (falls back to generated default)

2. Required modules: Microsoft Graph 2.x, ExchangeOnlineManagement, PSTeams

3. **Test thoroughly**: Always run with `-WhatIf` parameter first to preview changes before applying them.

> **IMPORTANT:** This is a sample solution and should be used by those comfortable testing, retesting, and validating before **even considering** using it in production. This content is provided _AS IS_ with _no_ guarantees or assumptions of quality, functionality, or support.

- You are responsible to comply with all applicable laws and regulations.
- With great power comes great responsibility.
- Friends don't let friends run untested directory scripts in production.

The script will extract employee data from BambooHR and for each user will execute one of the following processes:

1. **Attribute corrections** - If the user has an existing account, is an active employee, and the last changed time in Entra ID differs from BambooHR, this process will compare each Entra ID user object attribute with the data from BambooHR and correct them if necessary.
2. **Name change** - If the user has an existing account but the name does not match BambooHR, this process will correct the user's Name, UPN, and email address.
3. **New employee** - If there is no Entra ID account for the employee, this process will create a new user with the data extracted from BambooHR.

Variables usage description:

- Bamboo HR related variables:

  - $bhrDisplayName - The Display Name of the user in BambooHR
  - $bhrLastName - The Last name of the user in BambooHR
  - $bhrFirstName - The First Name of the user in BambooHR
  - $bhrLastChanged - The Date when the user's details were last changed in BambooHR
  - $bhrHireDate - The Hire Date of the user set in BambooHR
  - $bhrEmployeeNumber - The EmployeeID of the user set in BambooHR
  - $bhrJobTitle - The Job Title of the user set in BambooHR
  - $bhrDepartment - The Department of the user set in BambooHR
  - $bhrSupervisorEmail - The Manager of the user set in BambooHR
  - $bhrWorkEmail - The company email address of the user set in BambooHR
  - $bhrEmploymentStatus - The current status of the employee: Active, Terminated and if contains "Suspended" is in "maternity leave"
  - $bhrStatus - The employee account status in BambooHR: Valid values are "Active" and "Inactive"

- Entra Id related variables:

  - $entraIdUpnObjDetails - All Entra Id user object attributes extracted via WorkEmail lookup
  - $entraIdEidObjDetails - All Entra Id user object attributes extracted via EmployeeID lookup
  - $entraIdWorkEmail - UserPrincipalName/EmailAddress of the Entra Id user account - string
  - $entraIdJobTitle - Job Title of the Entra Id user account - string
  - $entraIdDepartment - Department of the Entra Id user account - string
  - $entraIdStatus - Login ability status of the Entra Id user account - boolean -can be True(Account is Active) or False(Account is Disabled)
  - $entraIdEmployeeNumber - Employee Number set on Entra Id user account(assigned by HR upon hire) - string
  - $entraIdSupervisorEmail - Direct Manager email address set on the Entra Id user account
  - $entraIdDisplayName - The Display Name set on the Entra Id user account - string
  - $entraIdFirstName - The Given Name set on the Entra Id user account - string
  - $entraIdLastName - The Surname set on the Entra Id user account - string
  - $entraIdCompanyName - The Company Name set on the Entra Id user account - string
  - $entraIdHireDate - The Employee Hire Date set on the Entra Id user account - string
  - $entraIdOfficeLocation - The Office Location set on the Entra Id user account - string
  - $entraIdWorkPhone - The Business Phone number set on the Entra Id user account - string
  - $entraIdMobilePhone - The Mobile Phone number set on the Entra Id user account - string

## Script Architecture

### Key Functions and Operations

**Runtime Initialization:**
- `Write-PSLog` - Enhanced logging with correlation tracking
- `Get-NewPassword` - Secure password generation
- `Initialize-PerformanceCache` - Caching for improved performance

**Data Extraction:**
- **Success**: Save employee data to `$employees` variable, clear `$response` to save memory, continue processing
- **Failure**: Send email alert, log error information, and terminate script

**Microsoft Graph Connection:**
- **Success**: Continue with user processing
- **Failure**: Send alert and terminate script

**User Processing:**
- Sequential processing of each employee with comprehensive error handling
- Retry logic for failed operations
- Teams notifications for status updates
