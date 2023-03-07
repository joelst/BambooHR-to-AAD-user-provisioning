# BambooHR to Azure AD user provisioning
User provisioning from BambooHR to AzAD

1. Make sure to replace the word "domain" with your domain name
2. Replace the emailaddresses for the alerting parts of the script with: the email address of the mailbox used to send the alert and the email address of the mailbox that will receive the alerts.
3. Insert BHR API key
4. Create AzAD Enterprise app for unattended auth via a certificate, for AzAD objects management + ExchangeOnline management
5. Modules needed: MGGraph module, ExchangeOnlineManagement module

The script will extract the employee data from BambooHR, then, for the data of each user, there are 3 operating blocks, that will run if the conditions are fulfilled. The 3 blocks are:
1. Attribute corrections - if the user has an existing account , and is an active employee, and, the last changed time in AzAD differs from BambooHR, then this first block will compare each of the AzAD User object attributes with the data extracted from BHR and correct them if necessary
2. Name Changing situations - If the user has an existing account, but the name does not match with the one from BHR, then, this block will run and correct the user Name, UPN,	emailaddress
3. If the user is a new employee, and there is no account in AzureAD for him, this script block will create a new user with the data extracted from BHR

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

###This portion describes the major function and logical operations that take place in the script:
<#
Initiate Script run time capture
{
	{Password generation Function}
	
	{Logging Function}
	
Extract Employee Data from BHR
	If BHR employee data extraction = successful -> Save data to $employees and clear the $Response variable to save memory -> Continue
	If BHR employee data extraction = Failed -> Send email alert, save error info to log file and terminate script

Connect to AzAd via graph module
	IF connection successful -> Continue
	IF connection failure -> Send alert + terminate script
	
$employees - system.array custom object containing an array of "user attributes" in blocks(arrays). Each block(array) represents 1 user and 
their details: Firstname,Lastname,email address, etc. The script cycles throug each block, and performs operations for each user attributes 
within the particular block of data. FOR EACH block of data within the variable $employees, representing 1 user with all its attributes, perform the below operations. There are 3 major script blocks: 1. User attribute correction, 2. User attribute correction for name changing situations, 3. User account creation
$employees|FOR EACH
	{
		Save the data of each attribute to a variable, to be used in comparison operations with the user attibutes in AzAD.
		
		GET user attributes from AzAD via UPN (UserPrincipalName, aka workmail in BHR).
		GET user attributes from AzAD via EmployeeID.
		The 2 GET operations are performed, in order to: verify if the user account exists in AzAD, to verify and match each of the user attributes between
		AzAD and BHR and to verify IF the name of the user has been changed (hence, the GET based on the EmployeeID).
		
		IF (EmployeeID of the account interrogated based on UPN(workmail) matches the EmployeeID extracted via BHR EmployeeID
		AND IF the UPN via workmail matches the UPN via EmployeeID and matches workmail
		AND IF objID extracted via UPN matches objID extracted via EmployeeID
		AND IF LastChanged in BHR match LastChanged on ExtensionAttribute1
		AND IF ObjectDetails extracted via EmployeeID IS not empty (the command to get the user object found an existing object with that EmployeeID)
		AND IF ObjectDetails extracted via UPN IS not empty (the command to get the user object found an existing object with that UPN)
		AND IF The Employee is not in suspended state (Suspended state is reserved for maternity leave) )
					THEN(Run the attibutes correction block){
						Save each AzAD attributes of the user object to be compared one by one with the details pulled from BambooHR
						Compare Job Title -> adjust in AzAD if different from BHR
						Compare Department -> adjust in AzAD if different from BHR
						Compare Manager -> adjust in AzAD if different from BHR
						Compare HireDate -> adjust in AzAD if different from BHR
						Compare Active status-> adjust in AzAD if different from BHR -> IF employee left the company (Inactive in BHR) -> Remove auth methods, change pass, 
                                                                                        deactivate AzAD account
																					 -> IF active in BHR but disabled in AzAD -> Activate AzAD account
						Compare EmployeeID -> adjust in AzAD if different from BHR
						Compare Company Name -> adjust in AzAD if different from BHR
						Compare Last Changed -> adjust in AzAD (ExtensionAttribute1) if different from BHR}
															
															
		IF (ObjID extracted via UPN not match ObjID extracted via EmployeeID and IF status in BHR not like "Suspended")
					THEN(Run the name correction block){
						Save Object Attributes extracted via EmployeeID to variables, to be compared with BHR attributes
						Compare LastModified -> Set LastModified from BambooHR to ExtensionAttribute1 in AzAD
						Compare LastName -> Change last name in Azure
						Compare Firstname -> Change FirstName in AzureAD
						Compare Display Name -> Set Display Name in AzureAD
						Compare EmailAddress -> Set EmailAddress in AzureAD
						Compare UPN -> Set UPN In AzureAD -> Send email to user and admin informing on the change}
														
		IF (the get-mguser command did not return any valid object (Both via UPN and EmployeeID) )
					THEN(Trigger the account creation block){
						Create Password
						Create account using New-MgUser command with the employee details saved in the variables at the begining of the "Foreach" block"
						IF -> account Creation successfull -> Sent the account details (login and pass) via email to admin
						   -> account Creation failed -> Send email to admin with the error details informing on the failure}
	}
	End Script
#>

