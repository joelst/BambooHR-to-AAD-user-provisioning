# BambooHR to Azure AD user provisioning
User provisioning from BambooHR to AzAD

This is part of Bamboo HR and Azure Active Directory integration process. This will make sure employees have Azure AD accounts. This does not function to create or update on premises Active Directory Domain Services accounts. 

Setting up this application as an Azure Function

1. Create Azure AD Enterprise application for unattended auth using a certificate for Azure AD object management and Exchange Online management
2. Set the following variables: 
 - BambooHRApiKey - API key created in BambooHR
 - AdminEmailAddress - Email address to recieve email alerts
 - CompanyName - Company name used for the URL to access BambooHR
 - TenantID - Microsoft Tenant Id
 - AADCertificateThumbPrint - Certificate thumbprint created for the application in AAD
 - AzAppClientID - Application created for Application.
3. Required modules: MGGraph module, ExchangeOnlineManagement module

The script will extract the employee data from BambooHR and for each user and will run one of the following processes:
1. Attribute corrections - if the user has an existing account , and is an active employee, and, the last changed time in Azure AD differs from BambooHR, then this first block will compare each of the AzAD User object attributes with the data extracted from BHR and correct them if necessary
2. Name changed - If the user has an existing account, but the name does not match with the one from BHR, then, this block will run and correct the user Name, UPN,	emailaddress
3. New employee, and there is no account in AzureAD for him, this script block will create a new user with the data extracted from BHR

Variables usage description:

- $BHR_displayName - The Display Name of the user in BambooHR
- $BHR_lastName - The Last name of the user in BambooHR
- $BHR_firstName - The First Name of the user in BambooHR
- $BHR_lastChanged - The Date when the user's details were last changed in BambooHR
- $BHR_hireDate - The Hire Date of the user set in BambooHR
- $BHR_employeeNumber - The EmployeeID of the user set in BambooHR
- $BHR_jobTitle - The Job Title of the user set in BambooHR
- $BHR_department - The Department of the user set in BambooHR
- $BHR_supervisorEmail - The Manager of the user set in BambooHR
- $BHR_workEmail - The company email address of the user set in BambooHR
- $BHR_EmploymentStatus - The current status of the employee: Active, Terminated and if contains "Suspended" is in "maternity leave"
- $HR_status - The employee account status in BambooHR: Valid values are "Active" and "Inactive"

- $azAD_UPN_OBJdetails - All AzAD user object attributes extracted via WorkEmail
- $azAD_EID_OBJdetails - All AzAD user object attributes extracted via EmployeeID
- $azAD_workemail - UserPrincipalName/EmailAddress of the AzureAD user account - string
- $azAD_jobTitle - Job Title of the AzureAd user account - string
- $azAD_department - Department of the AzureAD user account - string
- $azAD_status - Login ability status of the AzureAD user account - boolean -can be True(Account is Active) or False(Account is Disabled)
- $azAD_employeeNumber - Employee Number set on AzureAD user account(assigned by HR upon hire) - string
- $azAD_supervisorEmail - Direct Manager Name set on the AzureAD user account
- $azAD_displayname - The Display Name set on the AzureAD user account - string
- $azAD_firstName - The First Name set on the AzureAD user account - string
- $azAD_lastName - The Last Name set on the AzureAD user account - string
- $azAD_CompanyName - The company Name set on the AzureAD user account - string - Always will be "Tec Software Solutions"
- $azAD_hireDate - The Hire Date set on the AzureAD user account - string

## Major function and logical operations that take place in the script:

Initiate Script run time capture

	{Password generation Function}
	
	{Logging Function}
	
Extract Employee Data from BHR
	- If BHR employee data extraction = successful -> Save data to $employees and clear the $Response variable to save memory -> Continue
	- If BHR employee data extraction = Failed -> Send email alert, save error info to log file and terminate script

Connect to AzAd via graph module
	- If connection successful -> Continue
 	- If connection failure -> Send alert + terminate script
	
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

