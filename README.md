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

## Customization via Azure Automation variable

To avoid maintaining a personalized script, provide overrides in an Automation variable named `BHR_CustomizationsJson`. The script loads this JSON and applies overrides at runtime.

### Example JSON to match _Start-BambooHRUserProvisioning.ps1

```json
{
  "DaysToKeepAccountsAfterTermination": 14,
  "MailboxDelegationParams": [
    { "Group": "CG-SharedMailboxDelegatedAccessScheduling", "DelegateMailbox": "Scheduling" },
    { "Group": "CG-SharedMailboxDelegatedAccessCustomerCare", "DelegateMailbox": "CustomerCare" },
    { "Group": "CG-SharedMailboxDelegatedAccessSalesLeads", "DelegateMailbox": "Lead" },
    { "Group": "CG-SharedMailboxDelegatedAccessSocial", "DelegateMailbox": "Social" },
    { "Group": "CG-SharedMailboxDelegatedAccessJuneCommerce", "DelegateMailbox": "junec" }
  ]
}
```

You can also override values like `TeamsCardUri`, `AdminEmailAddress`, `NotificationEmailAddress`, `HelpDeskEmailAddress`, `UsageLocation`, `DaysAhead`, `EnableMobilePhoneSync`, `CurrentOnly`, `ForceSharedMailboxPermissions`, `DefaultProfilePicPath`, `EmailSignature`, and `WelcomeUserText`.

## Webhook-triggered sync

In addition to the scheduled reconciliation runbook, this repo now includes a standalone webhook-focused runbook:

- `Start-BambooHRUserProvisioning.ps1` - scheduled reconciliation / catch-up sync
- `Start-BambooHrWebhookSync.ps1` - targeted sync for BambooHR webhook events

The webhook runbook is meant to process only the employees identified in a BambooHR webhook payload while reusing the same core Entra ID lifecycle rules.

### Why keep both?

1. **Webhook runbook** gives faster reaction to employee changes.
2. **Scheduled runbook** remains the safety net for missed events, reconciliation, and broader periodic checks.

### Azure Automation setup

1. Publish `Start-BambooHrWebhookSync.ps1` as a separate PowerShell 7 runbook.
2. Create an **Azure Automation webhook** that points to this runbook.
3. Store the webhook URL securely when it is created. Azure only shows it once.
4. Optionally create a secure Automation variable named `BambooHrWebhookPrivateKey` if you want the runbook to validate BambooHR's HMAC signature headers.

> **Important:** Azure Automation webhooks are bearer-secret URLs. Treat them like passwords.

### BambooHR setup

Use a **Global Webhook** in BambooHR first for the simplest admin-managed integration.

Recommended settings:

1. **Destination URL:** the Azure Automation webhook URL
2. **Format:** `JSON`
3. **Monitor fields:** Work Email, Hire Date, Status / Employment Status, Reporting To, Job Title, Department, Division, Location, Mobile Phone / Work Phone, and other fields that materially affect provisioning decisions
4. **Posted fields:** employee identifier, work email, action, timestamp, and a few human-readable fields for logging

The runbook treats the webhook payload as a **change signal**, then re-reads BambooHR to retrieve the latest authoritative employee data before making changes in Entra ID.

## Changes & Updates

This section will keep track of changes made over time.

### TODO

- Create an Azure Bicep or ARM template for easy deployment.

### March 2026 changes

- Added `FullSync` and `ModifiedWithinDays` parameters. Delta sync (processing only recently modified employees) is now the default; `ModifiedWithinDays` controls the lookback window (default: 2 days). Pass `FullSync = $true` to process all employees regardless of last-changed date.
- Refactored feature-toggle parameters from `[switch]` to `[bool]` for Azure Automation runbook compatibility. Removed `MaxParallelUsers`.
- Significantly expanded Pester test suite: introduced AST-based function extraction (`Get-FunctionDefinitionsFromFile`), static variable-reference validation, and new tests covering hire date conversion, tenant email validation, phone number comparison, and offboarding completion markers.

### January 2026 changes

- Expanded offboarding pipeline with four new steps: removes the user's FIDO2 passkeys, reassigns all owned groups to the former manager, removes the user from all group memberships, and strips mailbox permissions.
- Improved error handling for mailbox permission removal and shared mailbox conversion failures; errors are now logged individually with clearer messages.
- Added `LicenseId` format validation with improved failure logging when the license check cannot complete.

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

1. **Azure Automation Account with System-Assigned Managed Identity**: Use a **system-assigned** managed identity (not user-assigned) so the identity lifecycle is tied to the Automation Account itself. If the account is deleted, the identity and its permissions are automatically revoked.
   - Review [Add-ManagedIdentityPermissions.ps1](./Add-ManagedIdentityPermissions.ps1) for permission setup
   - See [Security Hardening](#security-hardening) below for required access controls

2. **Azure Automation Variables**: Configure the following variables in your Azure Automation account:

- **BambooHrApiKey** - API key created in BambooHR (secure string)
- **BHRCompanyName** - Company name used for the URL to access BambooHR
- **CompanyName** - Your company display name
- **TenantId** - Microsoft Tenant ID
- **TeamsCardUri** - (Optional) Teams webhook URL for notifications
- **AdminEmailAddress** - (Optional) Administrator email for notifications (falls back to generated default)
- **NotificationEmailAddress** - (Optional) HR notification email address (falls back to generated default)
- **HelpDeskEmailAddress** - (Optional) Help desk email for user support (falls back to generated default)

2. **Required modules**:
   - `Microsoft.Graph.Users`, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Identity.SignIns`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Calendar`, `Microsoft.Graph.Files`
   - `ExchangeOnlineManagement`, `PSTeams`
   - **Optional** (script degrades gracefully): `Microsoft.Graph.DeviceManagement`, `Microsoft.Graph.DeviceManagement.Enrollment`, `Microsoft.Graph.Applications`

3. **Test thoroughly**: Always run with `-WhatIf` parameter first to preview changes before applying them.

### Security Hardening

This runbook's managed identity holds broad Graph API permissions (User.ReadWrite.All, Directory.ReadWrite.All, Mail.Send, Device.ReadWrite.All, and more). Anyone who can run a job in this Automation Account can exercise those permissions. The following controls are **strongly recommended**:

#### Use a System-Assigned Managed Identity

- **System-assigned** (not user-assigned) ties the identity lifecycle to the Automation Account. Deleting the account automatically revokes all permissions.
- A user-assigned identity can be attached to other resources, expanding the blast radius if any of them are compromised.
- Enable it under **Automation Account > Identity > System assigned > Status: On**.

#### Restrict Access to the Automation Account

- Assign Azure RBAC on the Automation Account resource to **only** privileged administrators who need to manage runbooks. Remove `Contributor` or `Owner` at the resource-group level for non-admin users.
- Recommended role assignments:
  | Role                     | Who                                  | Why                                                     |
  | ------------------------ | ------------------------------------ | ------------------------------------------------------- |
  | `Automation Contributor` | IT admins who manage runbooks        | Can edit/run jobs but not manage account-level settings |
  | `Automation Operator`    | Operators who trigger scheduled runs | Can start jobs but cannot edit runbook code             |
  | `Reader`                 | Auditors                             | Can view logs but not modify or execute                 |
- **Do not** grant `Automation Contributor` or `Automation Operator` to general staff, help desk, or service accounts that don't need it.
- Consider using [Azure PIM (Privileged Identity Management)](https://learn.microsoft.com/en-us/entra/id-governance/privileged-identity-management/pim-configure) for just-in-time elevation to these roles.

#### Limit Network and API Exposure

- If your tenant supports it, use [Conditional Access for workload identities](https://learn.microsoft.com/en-us/entra/identity/conditional-access/workload-identity) to restrict the managed identity to specific IP ranges (the Automation Account's outbound IPs).
- **Do not** expose the Automation Account's webhook URL (if configured) to untrusted networks.

#### Audit and Monitor

- Enable **Diagnostic Settings** on the Automation Account to send job logs to a Log Analytics workspace.
- Create an alert rule for unexpected job executions (e.g., jobs started outside scheduled windows or by unfamiliar principals).
- Periodically review the managed identity's app role assignments in **Entra ID > Enterprise Applications > (your Automation Account name) > Permissions**.

#### Microsoft Sentinel / Log Analytics Detections

If you use Microsoft Sentinel (or a Log Analytics workspace with Entra ID audit and sign-in logs), the following analytics rules are recommended to detect misuse of the managed identity or the Automation Account.

**1. Managed identity used outside expected schedule**

The runbook should only run during its scheduled windows. Any Graph API activity from the identity at other times is suspicious.

```kql
// Alert: Managed identity sign-in outside business hours or scheduled window
// Prerequisite: Entra ID sign-in logs sent to Log Analytics (AADServicePrincipalSignInLogs)
let automationAccountName = "<your-automation-account-name>";
let allowedHoursUtc = dynamic([6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]); // adjust to your schedule
AADServicePrincipalSignInLogs
| where TimeGenerated > ago(1d)
| where ServicePrincipalName =~ automationAccountName
| where hourofday(TimeGenerated) !in (allowedHoursUtc)
| project TimeGenerated, ServicePrincipalName, ServicePrincipalId, IPAddress, ResourceDisplayName, ResultType
```

**2. App role assignments added or removed from the managed identity**

Detects someone granting additional Graph permissions to the identity (privilege escalation).

```kql
// Alert: App role assignment changed for the Automation Account managed identity
// Prerequisite: Entra ID audit logs (AuditLogs)
let managedIdentityObjectId = "<your-managed-identity-object-id>";
AuditLogs
| where TimeGenerated > ago(7d)
| where OperationName has_any ("Add app role assignment to service principal",
                                "Remove app role assignment from service principal")
| mv-expand TargetResources
| where tostring(TargetResources.id) == managedIdentityObjectId
| project TimeGenerated, OperationName, InitiatedBy, TargetResources, CorrelationId
```

**3. Unusual Graph API operations from the managed identity**

If the identity starts performing bulk operations it normally wouldn't (e.g., reading mail content, mass user deletion), this detects the anomaly.

```kql
// Alert: Unusual resource access by the Automation Account identity
let automationAccountName = "<your-automation-account-name>";
AADServicePrincipalSignInLogs
| where TimeGenerated > ago(1d)
| where ServicePrincipalName =~ automationAccountName
| summarize CallCount = count() by ResourceDisplayName, bin(TimeGenerated, 1h)
| where CallCount > 500  // tune threshold based on your employee count
| project TimeGenerated, ResourceDisplayName, CallCount
```

**4. Automation Account runbook modified or created**

Detects unauthorized code changes to the runbook (someone injecting malicious code that runs with the identity's permissions).

```kql
// Alert: Runbook created or updated in the Automation Account
// Prerequisite: Azure Activity logs sent to Log Analytics
AzureActivity
| where TimeGenerated > ago(7d)
| where ResourceProviderValue =~ "Microsoft.Automation"
| where OperationNameValue has_any ("MICROSOFT.AUTOMATION/AUTOMATIONACCOUNTS/RUNBOOKS/WRITE",
                                     "MICROSOFT.AUTOMATION/AUTOMATIONACCOUNTS/RUNBOOKS/DRAFT/WRITE",
                                     "MICROSOFT.AUTOMATION/AUTOMATIONACCOUNTS/RUNBOOKS/PUBLISH/ACTION")
| project TimeGenerated, Caller, CallerIpAddress, OperationNameValue, ResourceGroup, _ResourceId
```

**5. Automation job started by unexpected principal**

Flags jobs started by users other than the schedule or the expected admin accounts.

```kql
// Alert: Automation job started by unfamiliar caller
// Prerequisite: Azure Activity logs
let allowedCallers = dynamic(["<schedule-principal-id>", "<admin-upn>"]);
AzureActivity
| where TimeGenerated > ago(1d)
| where ResourceProviderValue =~ "Microsoft.Automation"
| where OperationNameValue =~ "MICROSOFT.AUTOMATION/AUTOMATIONACCOUNTS/JOBS/WRITE"
| where Caller !in (allowedCallers)
| project TimeGenerated, Caller, CallerIpAddress, OperationNameValue, ResourceGroup
```

> **Tip:** Import these as **Scheduled Analytics Rules** in Sentinel with a 1-day lookback and appropriate entity mappings. For environments without Sentinel, create **Log Analytics alert rules** with the same queries.

#### Additional Recommendations

- Store the BambooHR API key as an **encrypted** Automation variable (not a plain-text variable).
- Rotate the BambooHR API key on a regular schedule (e.g., quarterly).
- If you use the shared mailbox delegation feature, the managed identity also needs the **Exchange Administrator** Entra role — review whether this is necessary for your deployment and remove it if not.

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
