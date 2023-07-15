# BambooHR to Azure AD user provisioning
User provisioning from BambooHR to AAD

## Updated 7/15/2023
This is a fork of the AMAZING work done by [PaulTony BHR-to-AzAd-user-Provisioning](https://github.com/PaulTony-coocoo/BHR-to-AzAD-user-provisioning). Their work saved me many hours having to put together this myself and I'm very grateful!

Anyone is free to take what is here and I will happily contribute my changes back to the original author, if they find my ideas worthy. 

What is different about this from the original?

1. It is focused on one project but trying to anticipate other scenarios. Primarily the changes are to make this work for this project.
2. Trying to use the [Microsoft Graph 2.x beta PowerShell modules](https://devblogs.microsoft.com/microsoft365dev/microsoft-graph-powershell-v2-is-now-in-public-preview-half-the-size-and-will-speed-up-your-automations/). There are some differences that forced changes.
3. Make this easily run using Azure Automation or as an Azure Function. So I'm adding parameters to allow for easy customization without modifying the code, making it easier to execute it in Azure Automation or an Azure Function.
4. My project is for a cloud-first AAD organization that but has on some hybrid user objects. This has lead to issues with some attributes not being writable. I'm trying to build a process that works for both withouth being too complicated.
5. I've got a _little OCD_ when it comes to silly things like variable names. Although I _really_ tried to fight my irrational desire to change them, I failed. No, there wasn't anything wrong with the original ones, just try explaining that to my OCD. 
6. I'd like to add sync for a couple other things like photos.

## Latest changes

- Fixed errors I introduced around name changes.
- If an employee is not active and their account is disabled don't keep synchronizing their department and other info. Also, set the department to "" when they are deactivated.
- Added days ahead parameter to set the number of days to look ahead for new hires.
- Changed email format to html to allow for improved formatting and clickable URLs.


### March 2023 changes

- Created parameters for common information needed.
- Updated password creation with a function that avoids characaters that are difficult to differentiate (like 0,O and 1,l,I) and avoids any characters that might cause PowerShell issues (like $ and `)
- The API key from BambooHR is automatically formatted, you can simply provide the key as-is from BambooHR
- Microsoft Graph 2.0-preview8 PowerShell moduled tested.
- Removed the line numbers from the log messages, because I messed up the line numbers. However the stack trace should provide enough info.
- Added `-TestOnly` parameter as a pseudo `-Whatif` parameter. It will log what would have been executed but will not make any changes.
- Changed the Bamboo report to pull future employees, as this fit my project needs. However I added the `-CurrentOnly` parameter if you would rather only process current employees.
- Moved most screen output to Verbose channel. If you are troubleshooting, run with -Verbose to see details.
- Fixed a bug in logging where the wrong time was being logged. You can see the fix in the incorrectly named [Original_BambooHr_User_Provisioning.ps1](./Original_BambooHR_User_Provisioning.ps1) as it has a couple small changes from the [upstream version](https://github.com/PaulTony-coocoo/BHR-to-AzAD-user-provisioning/blob/main/GENERALIZED_AUTO_USER_PROVISIONING.ps1).
- Added NotificationEmailAddress to copy all messages to so that HR or IT can keep an eye what changed process. This will give them the information needed to track down when a user's information changed.
- Added Sync-GroupMailboxDelegation function to set shared mailbox permissions. This need some work to upcate to latest APIs, however included because I am using it elsewhere.
- Added Photo Sync, when a new user is created it will attempt to add the photo from Bamboo to AAD. This is not kept in sync afterward, just on initial creation.
- Added Teams Adaptive Card logging. This is very simple and ugly now, but moved from sending an email to the webhook. This should get better in the future.

## Known issues

- Started work on importing the user's picture from BambooHR into AAD when the account is created. This does not work yet.
- Started work to build into an Azure Function, however this has not been checked in yet.
- Broke updating of the ExtensionAttribute1 with the LastUpdate from BHR. It does not even try to do this.
- User offboarding process has a couple issues and needs to be retested.
- There are a number of areas where the process can be streamlined and _may be_ addressed in the future.

If you have suggestions or questions, feel free to reach out. 

This is part of Bamboo HR and Azure Active Directory integration process. This will make sure employees have Azure AD accounts. This does not on its own create or update on premises Active Directory Domain Services accounts. 

## Setting up this application to run unattended (like as an Azure Function)

1. Create Azure AD Enterprise application for unattended auth using a certificate for Azure AD object management and Exchange Online management
 - Mail.Send
2. Set the following variables: 
 - BambooHRApiKey - API key created in BambooHR
 - AdminEmailAddress - Email address to recieve email alerts
 - CompanyName - Company name used for the URL to access BambooHR
 - TenantID - Microsoft Tenant Id
 - AADCertificateThumbPrint - Certificate thumbprint created for the application in AAD
 - AzAppClientID - Application created for Application.

3. Required modules: MGGraph 2.0 , ExchangeOnlineManagement module

## IMPORTANT: This is a sample solution and should be used by those comfortable testing and retesting and validating before thinking about using it in production. 
This content is provided *AS IS* with *no* guarantees or assumptions of quality, functionality, or support. 

- You are responsible to comply with all applicable laws and regulations. 
- With great power comes great responsibility.
- Friends don't let friends run untested directory scripts in production.
- Don't take any wooden nickels.

The script will extract the employee data from BambooHR and for each user and will run one of the following processes:
1. Attribute corrections - if the user has an existing account , and is an active employee, and, the last changed time in Azure AD differs from BambooHR, then this first block will compare each of the AzAD User object attributes with the data extracted from BHR and correct them if necessary
2. Name changed - If the user has an existing account, but the name does not match with the one from BHR, then, this block will run and correct the user Name, UPN,	emailaddress
3. New employee, and there is no account in AzureAD for him, this script block will create a new user with the data extracted from BHR

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

- AAD related variables:
  - $aadUPN_OBJdetails - All AzAD user object attributes extracted via WorkEmail
  - $aadEID_OBJdetails - All AzAD user object attributes extracted via EmployeeID
  - $aadWorkemail - UserPrincipalName/EmailAddress of the AzureAD user account - string
  - $aadJobTitle - Job Title of the AzureAd user account - string
  - $aadDepartment - Department of the AzureAD user account - string
  - $aadStatus - Login ability status of the AzureAD user account - boolean -can be True(Account is Active) or False(Account is Disabled)
  - $aadEmployeeNumber - Employee Number set on AzureAD user account(assigned by HR upon hire) - string
  - $aadSupervisorEmail - Direct Manager Name set on the AzureAD user account
  - $aadDisplayname - The Display Name set on the AzureAD user account - string
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


## Out of date 3/31/23

$employees - system.array custom object containing an array of "user attributes" in blocks(arrays). Each block(array) represents 1 user and their details: Firstname,Lastname,email address, etc. The script cycles throug each block, and performs operations for each user attributes within the particular block of data. FOR EACH block of data within the variable $employees, representing 1 user with all its attributes, perform the below operations. 

There are 3 major script blocks: 

1. User attribute correction, 
2. User attribute correction for name changing situations, 
3. User account creation

For each of the returned employees:

1. Save the data of each attribute to a variable, to be used in comparison operations with the user attibutes in AzAD.
	- GET user attributes from AzAD via UPN (UserPrincipalName, aka workmail in BHR).
	- GET user attributes from AzAD via EmployeeID.
	- The 2 GET operations are performed, in order to: verify if the user account exists in AzAD, to verify and match each of the user attributes between AzAD and BHR and to verify IF the name of the user has been changed (hence, the GET based on the EmployeeID).
		
	- If (EmployeeID of the account interrogated based on UPN(workmail) matches the EmployeeID extracted via BHR EmployeeID
		- And if the UPN via workmail matches the UPN via EmployeeID and matches workmail
		- And if objID extracted via UPN matches objID extracted via EmployeeID
		- And if LastChanged in BHR match LastChanged on ExtensionAttribute1
		- And if ObjectDetails extracted via EmployeeID IS not empty (the command to get the user object found an existing object with that EmployeeID)
		- And if ObjectDetails extracted via UPN IS not empty (the command to get the user object found an existing object with that UPN)
		- And if The Employee is not in suspended state (Suspended state is reserved for maternity leave) )
			- Then (Run the attibutes correction block)
				- Save each AzAD attributes of the user object to be compared one by one with the details pulled from BambooHR
				- Compare Job Title -> adjust in AzAD if different from BHR
				- Compare Department -> adjust in AzAD if different from BHR
				- Compare Manager -> adjust in AzAD if different from BHR
				- Compare HireDate -> adjust in AzAD if different from BHR
				- Compare Active status-> adjust in AzAD if different from BHR 
					- If employee left the company (Inactive in BHR) -> Remove auth methods, change pass, deactivate AzAD account
					- If active in BHR but disabled in AzAD -> Activate AzAD account
				- Compare EmployeeID -> adjust in AzAD if different from BHR
				- Compare Company Name -> adjust in AzAD if different from BHR
				- Compare Last Changed -> adjust in AzAD (ExtensionAttribute1) if different from BHR}													
															
	- If (ObjID extracted via UPN not match ObjID extracted via EmployeeID and IF status in BHR not like "Suspended")
		- Then (Run the name correction block){
			- Save Object Attributes extracted via EmployeeID to variables, to be compared with BHR attributes
			- Compare LastModified -> Set LastModified from BambooHR to ExtensionAttribute1 in AzAD
			- Compare LastName -> Change last name in Azure
			- Compare Firstname -> Change FirstName in AzureAD
			- Compare Display Name -> Set Display Name in AzureAD
			- Compare EmailAddress -> Set EmailAddress in AzureAD
			- Compare UPN -> Set UPN In AzureAD -> Send email to user and admin informing on the change}
														
	- If (the get-mguser command did not return any valid object (Both via UPN and EmployeeID) )
		- Then (Trigger the account creation block){
			- Create Password
			- Create account using New-MgUser command with the employee details saved in the variables at the begining of the "Foreach" block"
			- If account Creation successfull -> Sent the account details (login and pass) via email to admin
			- If ccount Creation failed -> Send email to admin with the error details informing on the failure}

