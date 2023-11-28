# BambooHR to Azure AD user provisioning
User provisioning from BambooHR to Entra Id formerly known as Azure Active Directory (AAD)

## Introduction

This is a fork of the AMAZING work done by [PaulTony BHR-to-AzAd-user-Provisioning](https://github.com/PaulTony-coocoo/BHR-to-AzAD-user-provisioning). Their work saved me many hours having to put together this myself and I'm very grateful!

Anyone is free to take what is here and I will happily contribute my changes back to the original author, if they find my ideas worthy. 

What is different about this from the original?

1. It is focused on one project but trying to anticipate other scenarios. Primarily the changes are to make this work for this project.
2. Moving to use the [Microsoft Graph 2.x PowerShell modules](https://devblogs.microsoft.com/microsoft365dev/microsoft-graph-powershell-v2-is-now-in-public-preview-half-the-size-and-will-speed-up-your-automations/). There are some differences that forced changes.
3. Make it run using Azure Automation or as an Azure Function. So I'm adding parameters to allow for easy customization without modifying the code, making it easier to execute it in Azure Automation or an Azure Function.
4. My project is for a cloud-first Entra Id (AAD) organization that but has on some hybrid user objects. This has lead to issues with some attributes not being writable.
5. I've got a _little OCD_ when it comes to silly things like variable names. Although I _really_ tried to fight my irrational desire to change them, I failed. No, there wasn't anything wrong with the original ones, just try explaining that to my OCD. 

## Changes & Updates

This section will keep track of changes made over time.

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
- Added `-TestOnly` parameter as a pseudo `-Whatif` parameter. It will log what would have been executed but will not make any changes.
- Changed the Bamboo report to pull future employees, as this was required for my project. Added the `-CurrentOnly` parameter if you would rather only process current employees.
- Moved most screen output to Verbose channel. If you are troubleshooting, run with -Verbose to see details.
- Fixed a bug in logging where the wrong time was being logged. You can see the fix in the incorrectly named [Original_BambooHr_User_Provisioning.ps1](./Original_BambooHR_User_Provisioning.ps1) as it has a couple small changes from the [upstream version](https://github.com/PaulTony-coocoo/BHR-to-AzAD-user-provisioning/blob/main/GENERALIZED_AUTO_USER_PROVISIONING.ps1).
- Added NotificationEmailAddress to copy all messages to so that HR or IT can keep an eye what changed process. This will give them the information needed to track down when a user's information changed.
- Added Sync-GroupMailboxDelegation function to set shared mailbox permissions.
- Added Photo Sync, when a new user is created it will attempt to add the photo from Bamboo to AAD. This is not kept in sync afterward, just on initial creation.
- Added Teams Adaptive Card logging. This is very simple and ugly now, but moved from sending an email to the webhook. This should get better in the future.

## Known issues

- ExtensionAttribute1 with the LastUpdate from BHR no longer works.
- User off boarding process may have a couple issues that needs to be retested.
- There are a number of areas where the process can be streamlined and _may be_ addressed in the future.
- Entra Id accounts are disabled and not deleted. You will need to either add logic to automatically delete accounts or do this manually.

## Future ideas

If you have suggestions or questions, feel free to reach out. 

- Add off boarding steps for things like:
  - One Drive for Business data access
  - Transfer group ownership
  - Deleting disabled accounts after suitable waiting period.
  - Other suggestions


## Getting started

This is part of Bamboo HR and Entra Id (Azure Active Directory) integration process. This will make sure employees have Entra Id accounts. This does not create or update on premises Active Directory Domain Services accounts. 

1. Create Azure AD Enterprise application for unattended auth using a certificate for Azure AD object management and Exchange Online management
 - Mail.Send

2. Set the following variables: 
 - BambooHRApiKey - API key created in BambooHR
 - AdminEmailAddress - Email address to receive email alerts
 - CompanyName - Company name used for the URL to access BambooHR
 - TenantID - Microsoft Tenant Id
 - AADCertificateThumbPrint - Certificate thumbprint created for the application in AAD
 - AzAppClientID - Application created for Application.

3. Required modules: MGGraph 2.0, ExchangeOnlineManagement, PSTeams

> **IMPORTANT:** This is a sample solution and should be used by those comfortable testing, retesting, and validating before **thinking** about using it in production. This content is provided *AS IS* with *no* guarantees or assumptions of quality, functionality, or support. 

- You are responsible to comply with all applicable laws and regulations. 
- With great power comes great responsibility.
- Friends don't let friends run untested directory scripts in production.
- Don't take any wooden nickels.

The script will extract the employee data from BambooHR and for each user and will run one of the following processes:
1. Attribute corrections - if the user has an existing account , and is an active employee, and, the last changed time in Azure AD differs from BambooHR, then this first block will compare each of the AzAD User object attributes with the data extracted from BHR and correct them if necessary.
2. Name changed - If the user has an existing account, but the name does not match with the one from BHR, then, this block will run and correct the user Name, UPN,	emailAddress.
3. New employee, and there is no Entra Id account, this script block will create a new user with the data extracted from BHR.

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

- Entra Id (AAD) related variables:
  - $aadUPN_OBJdetails - All AzAD user object attributes extracted via WorkEmail
  - $aadEID_OBJdetails - All AzAD user object attributes extracted via EmployeeID
  - $aadWorkEmail - UserPrincipalName/EmailAddress of the AzureAD user account - string
  - $aadJobTitle - Job Title of the AzureAd user account - string
  - $aadDepartment - Department of the AzureAD user account - string
  - $aadStatus - Login ability status of the AzureAD user account - boolean -can be True(Account is Active) or False(Account is Disabled)
  - $aadEmployeeNumber - Employee Number set on AzureAD user account(assigned by HR upon hire) - string
  - $aadSupervisorEmail - Direct Manager Name set on the AzureAD user account
  - $aadDisplayName - The Display Name set on the AzureAD user account - string
  - $aadFirstName - The First Name set on the AzureAD user account - string
  - $aadLastName - The Last Name set on the AzureAD user account - string
  - $aadCompanyName - The company Name set on the AzureAD user account - string - Always will be "Tec Software Solutions"
  - $aadHireDate - The Hire Date set on the AzureAD user account - string

## Major functions and logical operations that take place in the script:

Initiate Script run time capture

 - Write-Log
 - New-Password

Extract Employee Data from BHR

  - If BHR employee data extraction = successful -> Save data to $employees and clear the $Response variable to save memory -> Continue
  - If BHR employee data extraction = Failed -> Send email alert, save error info to log file and terminate script

Connect to AzAd via graph module

  - If connection successful -> Continue
  - If connection failure -> Send alert + terminate script
