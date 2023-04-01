<#

This file is originally from https://github.com/PaulTony-coocoo/BHR-to-AzAD-user-provisioning

The script will extract the employee data from BambooHR, then, for the data of each user, there are 3 operating blocks, that will run if
the conditions are fulfilled. The 3 blocks are:
						1. Attribute corrections - if the user has an existing account , and is an active employee, and, the last changed time in AzAD differs from BambooHR, then 
                        this first block will compare each of the AzAD User object attributes with the data extracted from BHR and correct them if necessary
						2. Name Changing situations - If the user has an existing account, but the name does not match with the one from BHR, then, this block will run and 
                        correct the user Name, UPN,	emailaddress
						3. If the user is a new employee, and there is no account in AzureAD for him, this script block will create a new user with the data extracted from BHR
Variables usage description:
$BHR_displayName - The Display Name of the user in BambooHR
$BHR_lastName - The Last name of the user in BambooHR
$BHR_firstName - The First Name of the user in BambooHR
$BHR_lastChanged - The Date when the user's details were last changed in BambooHR
$BHR_hireDate - The Hire Date of the user set in BambooHR
$BHR_employeeNumber - The EmployeeID of the user set in BambooHR
$BHR_jobTitle - The Job Title of the user set in BambooHR
$BHR_department - The Department of the user set in BambooHR
$BHR_supervisorEmail - The Manager of the user set in BambooHR
$BHR_workEmail - The company email address of the user set in BambooHR
$BHR_EmploymentStatus - The current status of the employee: Active, Terminated and if contains "Suspended" is in "maternity leave"
$BHR_status - The employee account status in BambooHR: Valid values are "Active" and "Inactive"
$azAD_UPN_OBJdetails - All AzAD user object attributes extracted via WorkEmail
$azAD_EID_OBJdetails - All AzAD user object attributes extracted via EmployeeID
$azAD_workemail - UserPrincipalName/EmailAddress of the AzureAD user account - string
$azAD_jobTitle - Job Title of the AzureAd user account - string
$azAD_department - Department of the AzureAD user account - string
$azAD_status - Login ability status of the AzureAD user account - boolean -can be True(Account is Active) or False(Account is Disabled)
$azAD_employeeNumber - Employee Number set on AzureAD user account(assigned by HR upon hire) - string
$azAD_supervisorEmail - Direct Manager Name set on the AzureAD user account
$azAD_displayname - The Display Name set on the AzureAD user account - string
$azAD_firstName - The First Name set on the AzureAD user account - string
$azAD_lastName - The Last Name set on the AzureAD user account - string
$azAD_CompanyName - The company Name set on the AzureAD user account - string - Always will be "Company Name"
$azAD_hireDate - The Hire Date set on the AzureAD user account - string
###This portion describes the major function and logical operations that take place in the script:
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
within the particular block of data. FOR EACH block of data within the variable $employees, representing 1 user with all its attributes, perform the below operations. There are 3 major
script blocks: 1. User attribute correction, 2. User attribute correction for name changing situations, 3. User account creation
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
						Compare Active status-> adjust in AzAD if different from BHR -> IF employee left the company (Inactive in BHR) -> Remove auth methods, change pass, deactivate AzAD account
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

#Script start
$runtime = Measure-Command -Expression {
    #####################################Provision users to AzureAD using the employee details from BambooHR#######################################



    #####ERROR LOGGING FUNCTION#####
    $log_filename = "Log_" + (get-date -Format yyyy-MM-dd_HH-mm-s) + ".csv"
    $log_file = "C:\Reports\automation\$log_filename"
    function Write-Log {
        [CmdletBinding()]
        param(
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [string]$Message,
 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('Information', 'Warning', 'Error')]
            [string]$Severity = 'Information'
        )

        [pscustomobject]@{
            Time     = (Get-Date -Format "yyyy/MM/dd HH:mm:s")
            Message  = $Message
            Severity = $Severity
        } | Export-Csv -Path $log_file -Append -NoTypeInformation
    }#####ERROR LOGGING FUNCTION BLOCK CLOSURE##### 


    #####PASSWORD GENERATOR FUNCTION#####
    function Get-RandomPassword {
        param (
            [Parameter(Mandatory)]
            [int] $length
        )
        $charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!?-_'.ToCharArray()
        $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $bytes = New-Object byte[]($length)
 
        $rng.GetBytes($bytes)
 
        $result = New-Object char[]($length)
 
        for ($i = 0 ; $i -lt $length ; $i++) {
            $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
        }
 
        return (-join $result)
    }#####PASSWORD GENERATOR FUNCTION BLOCK CLOSURE#####


    #Getting all users details from BambooHR and passing the extracted info to the variable $employees
    $headers = @{}
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Authorization", "asd APIKEYHERE")
    $error.clear()
    Try {
        Invoke-RestMethod `
            -Uri 'https://api.bamboohr.com/api/gateway.php/domain/v1/reports/custom?format=json&onlyCurrent=true' `
            -Method POST `
            -Headers $headers `
            -ContentType 'application/json' `
            -Body '{"fields":["status","hireDate","department","employeeNumber","firstName","lastName","displayName","jobTitle","supervisorEmail","workEmail","lastChanged","employmentHistoryStatus"]}' `
            -OutVariable response 

    } 
    Catch {
        #If error returned, the API call to BambooHR failed and no usable employee data has been returned, write to log file and exit script
        
        $BHR_api_call_err = $_
        $BHR_api_call_err_message = $BHR_api_call_err.Exception.Message
        $BHR_api_call_err_category = $BHR_api_call_err.CategoryInfo.Category
        $BHR_api_call_err_stack = $BHR_api_call_err.ScriptStackTrace
        
        Write-log -Message "API call to return user information from BambooHR has failed. Script terminated at line 225. Following are the error details. `
EXCEPTION MESSAGE: `
$BHR_api_call_err_message `
CATEGORY: `
$BHR_api_call_err_category `
SCRIPT STACK TRACE: `
$BHR_api_call_err_stack" -Severity Error
        #Send email alert with the generated error
        $params = @{
            Message         = @{
                Subject      = "User creation automation: Connection to BambooHR has failed"
                Body         = @{
                    ContentType = "text"
                    Content     = "Hello, `
`
Connection to BambooHR endpoint has failed. Below are the error details generated. `
EXCEPTION MESSAGE: `
$BHR_api_call_err_message `
CATEGORY: `
$BHR_api_call_err_category `
SCRIPT STACK TRACE: `
$BHR_api_call_err_stack `
`
Regards, `
Automated User Account Management Service"
                }
                ToRecipients = @(
                    @{
				                    EmailAddress = @{
                            Address = "admin_monitor@domain.com"
                        }
                    }
                )
            }
            SaveToSentItems = "True"
        }

        Send-MgUserMail -BodyParameter $params -UserId Impersonated_Mailbox@domain.com

        Exit
    }#INVOKE-RESTMETHOD CATCH ERROR BLOCK CLOSURE 

    #If no error returned, it means that the script was not interrupted by the "Exit" command within the "Catch" block. Write info below to log file and continue
    Write-Log -Message "Successfully extracted the employees information from BambooHR. Line 168 'Try' did not generate errors" -Severity Information
  

    #Saving only the employee data to $employees variable and eliminate $response variable to save memory
    $employees = $response.employees
    Remove-Variable -Name response

    #Connect to AzureAD using PS Graph Module, authenticating as the configured service principal for this operation, with certificate auth
    $error.Clear()

    Connect-MgGraph -TenantId domain.onmicrosoft.com -CertificateThumbprint CERTTHUMBPRINTHERE -ClientId AzAPPClientID
    if ($?) {
        #If no error returned, write to log file and continue
        Write-Log -Message "Successfully connected to tenant. Command at line 239 did not generate errors." -Severity Information
    }
    else {

        #If error returned, write to log file and exit script
        $mgconnerror = $error
        $mgerr_exception = $mgconnerror.Exception
        $mgerr_category = $mgconnerror.CategoryInfo
        $mgerr_errID = $mgconnerror.FullyQualifiedErrorId
        $mgerr_stack = $mgconnerror.ScriptStackTrace
        Write-Log -Message "Connection to tenant has failed. Script terminated at line 299. Below are the error details.`
EXCEPTION: `
$mgerr_exception `
CATEGORY: `
$mgerr_category `
FULLY QUALIFIED ERROR ID: `
$mgerr_errID `
SCRIPT STACK TRACE: `
$mgerr_stack" -Severity Error

        #Send email alert with the generated error 
        $params = @{
            Message         = @{
                Subject      = "User creation automation: MGGraph Connection to AzAD has failed"
                Body         = @{
                    ContentType = "text"
                    Content     = "Hello, `
`
Connection to AzAD tenant has failed. Below are the error details generated. `
EXCEPTION: `
$mgerr_exception `
CATEGORY: `
$mgerr_category `
FULLY QUALIFIED ERROR ID: `
$mgerr_errID `
SCRIPT STACK TRACE: `
$mgerr_stack `
`
Regards, `
Automated User Account Management Service"
                }
                ToRecipients = @(
                    @{
				                    EmailAddress = @{
                            Address = "admin_monitor@domain.com"
                        }
                    }
                )
            }
            SaveToSentItems = "True"
        }

        Send-MgUserMail -BodyParameter $params -UserId Impersonated_Mailbox@domain.com

        Disconnect-MgGraph
        Exit
    }
        
    Write-Log -Message "Entering Foreach loop to check the attributes of each user at a time. Foreach loop initiated at line 304" -Severity Information
    
    $employees | ForEach-Object {
        $error.Clear()

        #####On each loop, pass all employee data from BambooHR to variables, to be compared one by one with the user data from AzAD and set them, if necessary
        $BHR_lastChanged = $_.lastChanged
        $BHR_hireDate = $_.hireDate
        $BHR_employeeNumber = $_.employeeNumber
        $BHR_jobTitle = $_.jobTitle
        $BHR_department = $_.department
        $BHR_supervisorEmail = $_.supervisorEmail
        $BHR_workEmail = $_.workEmail
        $BHR_EmploymentStatus = $_.employmentHistoryStatus
        #Translating user "status" from BambooHR to boolean, to match and compare with the AzureAD user account status
        $BHR_status = $_.status
        if ($BHR_status -eq "Inactive")
        { $BHR_AccountEnabled = $False }
        if ($BHR_status -eq "Active")
        { $BHR_AccountEnabled = $True }

        #Normalizing user names, eliminating language specific characters
        $BHR_firstName = $_.firstName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t"
        $BHR_lastName = $_.lastName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t"
        $BHR_displayName = $_.displayName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t"
        $BHR_firstName = (get-culture).textinfo.ToTitleCase($BHR_firstName)
        $BHR_lastName = (get-culture).textinfo.ToTitleCase($BHR_lastName)
        $BHR_displayName = (get-culture).textinfo.ToTitleCase($BHR_displayName)
        Write-Log -Message "Eliminated language specific characters from names and capitalize first letters. Lines 324-329." -Severity Information

        Write-Log -Message "Verifying employee with the following details in BambooHR:
Firstname: $BHR_firstName
Lastname: $BHR_lastName
Display Name: $BHR_displayName
Work Email: $BHR_workEmail
Department: $BHR_department
Job Title: $BHR_jobTitle
Manager: $BHR_supervisorEmail
HireDate: $BHR_hireDate
Employee Number: $BHR_employeeNumber
Employee Status: $BHR_status" -Severity Information
        $azAD_UPN_OBJdetails = $null
        $azAD_EID_OBJdetails = $null
        <#
If the user start date is in the past, or in less than 7 days from current time, we can begin processing the user: create AzAD account
or correct the attributes in AzureAD for the employee, else, the employee found on BambooHR will not be processed
#>


        #$current_date = (get-date)
        if ((([datetime]$BHR_hireDate).AddDays(-14)) -lt (Get-Date)) {

            $error.clear()
    
            $azAD_UPN_OBJdetails = $null
            $azAD_EID_OBJdetails = $null

            #Checking if the user exists in AzAD and If there is an account with the EmployeeID of the user checked in the current loop
            Write-Log -Message "Checking if $BHR_workEmail has a valid and matching AzureAD account. At lines 363-364." -Severity Information
            Get-MgUser -UserId $BHR_workEmail -OutVariable azAD_UPN_OBJdetails -Property id, userprincipalname, Department, EmployeeID, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, EmployeeHireDate, OnPremisesExtensionAttributes -ExpandProperty manager
            Get-MgUser -Filter "employeeID eq '$BHR_employeenumber'" -OutVariable azAD_EID_OBJdetails -Property employeeid, userprincipalname, Department, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, EmployeeHireDate, OnPremisesExtensionAttributes  -ExpandProperty manager
            $error.clear()
            $UPN_ExtensionAttribute1 = ($azAD_UPN_OBJdetails | Select-Object @{Name = 'ExtensionAttribute1'; Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } }).ExtensionAttribute1
            $EID_ExtensionAttribute1 = ($azAD_EID_OBJdetails | Select-Object @{Name = 'ExtensionAttribute1'; Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } }).ExtensionAttribute1
   
            <#### If: empID of object returned by UPN or by empID is equal, and if object ID is the same as object ID of object returned by UPN and EmpID and
    if UPN = workemail from bamboo  AND if the last changed date from BambooHR is NOT equal to the last changed date saved in ExtensionAttribute1 in AzAD, check each attribute and set
    them correctly, according to BambooHR#>
   
            if ($azAD_EID_OBJdetails.EmployeeId -eq $azAD_UPN_OBJdetails.EmployeeId -and $azAD_EID_OBJdetails.UserPrincipalName -eq $azAD_UPN_OBJdetails.UserPrincipalName -eq $BHR_workEmail -and $azAD_UPN_OBJdetails.id -eq $azAD_EID_OBJdetails.id -and $BHR_lastChanged -ne $upn_ExtensionAttribute1 -and $azAD_EID_OBJdetails.Capacity -ne 0 -and $azAD_UPN_OBJdetails.Capacity -ne 0 -and $BHR_EmploymentStatus -notlike "*suspended*") { 
                Write-Log  -Message "$BHR_workEmail is a valid AzureAD Account, with matching EmployeeID and UPN from AzureAD to BambooHR, but different last modified date. IF statement passed as true, at line 373." -Severity Warning
                #saving AzAD attributes to be compared one by one with the details pulled from BambooHR
                $azAD_workemail = $azAD_UPN_OBJdetails.Mail
                $azAD_jobTitle = $azAD_UPN_OBJdetails.JobTitle
                $azAD_department = $azAD_UPN_OBJdetails.Department
                $azAD_status = $azAD_UPN_OBJdetails.AccountEnabled
                $azAD_employeeNumber = $azAD_UPN_OBJdetails.EmployeeID
                $azAD_supervisorEmail = ($azAD_UPN_OBJdetails | Select-Object @{Name = 'Manager'; Expression = { $_.Manager.AdditionalProperties.mail } }).manager
                $azAD_displayname = $azAD_UPN_OBJdetails.displayname
                $azAD_firstName = $azAD_UPN_OBJdetails.GivenName
                $azAD_lastName = $azAD_UPN_OBJdetails.Surname
                $azAD_CompanyName = $azAD_UPN_OBJdetails.CompanyName
                if ($azAD_UPN_OBJdetails.EmployeeHireDate) {
                    $azAD_hireDate = $azAD_UPN_OBJdetails.EmployeeHireDate.AddHours(12).ToString("yyyy-MM-dd") 
                }
                Write-Log -Message "AzAD user object found with the following attributes:
    Firstname: $azAD_firstName
    Lastname: $azAD_lastName
    Display Name: $azAD_displayname
    Work Email: $azAD_workemail
    Department: $azAD_department
    Job Title: $azAD_jobTitle
    Manager: $azAD_supervisorEmail
    HireDate: $azAD_hireDate
    Employee Number: $azAD_employeeNumber
    Employee Enabled: $azAD_status" -Severity Information
                $error.clear() 
                #Checking JobTitle if correctly set, if not, configure the JobTitle as set in BambooHR
                if ($azAD_jobTitle -ne $BHR_jobTitle) {

                    Update-MgUser -UserId $BHR_workEmail -JobTitle $BHR_jobTitle
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the JobTitle of $BHR_workEmail. Command on line 411 returned error. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        $error.Clear()
                        Write-Log -Message "JobTitle for $BHR_workEmail in AzureAD:$azAD_jobTitle and in BambooHR:$BHR_jobTitle.
Set the JobTitle found on BambooHR to the Azure User Object. IF condition result is True at line 408" -Severity Warning
                    }
                }
        
                #Checking department if correctly set, if not, configure the Department as set in BambooHR
                if ($azAD_department -ne $BHR_department) {
 
                    Update-MgUser -UserId $BHR_workEmail -Department $BHR_department
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the Department of $BHR_workEmail. Command on line 442 returned error. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        $error.Clear()
                        Write-Log -Message "Department for $BHR_workEmail in AzureAD:$azAD_department and in BambooHR:$BHR_department. `
Setting the Department found on BambooHR to the Azure User Object. IF condition result is True at line 439" -Severity Information
                    }
                }
        
                #Checking the manager if correctly set, if not, configure the manager as set in BambooHR
                if ($azAD_supervisorEmail -ne $BHR_supervisorEmail) {
                    $azAD_managerID = (Get-MgUser -UserId $BHR_supervisorEmail | Select-Object id).id
                    $NewManager = @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$azAD_managerID"
                    }

                    Set-MgUserManagerByRef -UserId $BHR_workEmail -BodyParameter $NewManager
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the Manager of $BHR_workEmail. Command on line 477 returnet error. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        $error.Clear()
                        Write-Log -Message "Manager of $BHR_workEmail in AzureAD:$azAD_supervisorEmail and in BambooHR:$BHR_supervisorEmail. `
Setting the Manager found on BambooHR to the Azure User Object. IF condition result is True at line 470" -Severity Warning
                    }
                }
            
                #Check and set the Employee Hire Date
                if ($azAD_hireDate -ne $BHR_hireDate) {
                
                    Update-MgUser -UserId $BHR_workEmail -EmployeeHireDate $BHR_hireDate
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the Employee Hire Date. Command on line 508 returnet error. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        $error.Clear()
                        Write-Log -Message "Setting the Hire date of $BHR_workEmail to $BHR_hireDate. IF condition result is True at line 505" -Severity Warning
                    }
                }

                #######Check if user is active in BambooHR, and set the status of the account as it is in BambooHR (active or inactive)
           
                if ($BHR_AccountEnabled -eq $False -and $BHR_EmploymentStatus -eq "Terminated") {

                    #As the account is marked "Inactive" in BHR and "Active" in AzAD, block sign-in, revoke sessions, change pass, remove auth methods
                    $error.clear()
                    Update-MgUser -UserId $BHR_workEmail -AccountEnabled:$BHR_AccountEnabled
                    Revoke-MgUserSign -UserId $BHR_workEmail

                    #Change to a random pass
                    $params = @{
                        PasswordProfile = @{
                            ForceChangePasswordNextSignIn = $true
                            Password                      = (Get-RandomPassword 12)
                        }
                    }
                    Update-MgUser -UserId $BHR_workEmail -BodyParameter $params
                    Get-MgUserMemberOf -UserId $BHR_workEmail
                    #Convert mailbox to shared
                    Connect-ExchangeOnline -CertificateThumbprint CERTTHUMBPRINT -AppId AzAPPID -Organization domain.onmicrosoft.com
                    Set-Mailbox -Identity $BHR_workEmail -Type Shared
                    Disconnect-ExchangeOnline -Confirm:$False
                    #Remove Licenses
                    Get-MgUserLicenseDetail -UserId $BHR_workEmail | ForEach-Object { Set-MgUserLicense -UserId $BHR_workEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }
                    Get-MgUserMemberOf -UserId $BHR_workEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $azAD_UPN_OBJdetails.id } #Remove Groups
                    $methodID = Get-MgUserAuthenticationMethod -UserId $BHR_workEmail | Select-Object id 
                    $methodsdata = Get-MgUserAuthenticationMethod -UserId $BHR_workEmail | Select-Object -ExpandProperty AdditionalProperties
                    $methods_count = ($methodID | Measure-Object | Select-Object count).count

                    #Pass through each authentication method and remove them
                    $error.Clear() 
                    for ($i = 0 ; $i -lt $methods_count ; $i++) {
       
                        if ((($methodsdata[$i]).Values) -like "*phoneAuthenticationMethod*") { Remove-MgUserAuthenticationPhoneMethod -UserId $BHR_workEmail -PhoneAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed phone auth method for $BHR_workEmail. Line 568." -Severity Warning }
                        if ((($methodsdata[$i]).Values) -like "*microsoftAuthenticatorAuthenticationMethod*") { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $BHR_workEmail -MicrosoftAuthenticatorAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed auth app method for $BHR_workEmail. Line 569." -Severity Warning }
                        if ((($methodsdata[$i]).Values) -like "*windowsHelloForBusinessAuthenticationMethod*") { Remove-MgUserAuthenticationFido2Method -UserId $BHR_workEmail -Fido2AuthenticationMethodId ($methodID[$i]).id ; Write-Log -Message "Removed PIN auth method for $BHR_workEmail. Line 570." -Severity Warning }
                    }
                    Update-MgUser -EmployeeId "LVR" -UserId $BHR_workEmail                                            
                    if ($error.Count -ne 0) {
                        $error | ForEach-Object {
                            $err_Exception = $_.Exception
                            $err_Target = $_.TargetObject
                            $err_Category = $_.CategoryInfo
                            Write-Log "Lines 568-570:AUTH METHOD REMOVAL ERROR. Could not remove authentication details. Error details below: `
                            Exception: `
                            $err_Exception `
                            Target Object: `
                            $err_Target `
                            Error Category: `
                            $err_Category " -Severity Error
                        }
                    }
                    else {
                        Write-Log -Message "Account $BHR_workEmail marked as inactive in BambooHR but found active in AzAD. Disabled AzAD account, revoked sessions and removing auth methods. IF condition result is True at line 536" -Severity Warning              
                        $error.Clear()
                    }
                }
          
                if ($BHR_AccountEnabled -eq $True -and $azAD_status -eq $False) {
                    # The account is marked "Active" in BHR and "Inactive" in AzAD, enable the AzAD account
                
                    Update-MgUser -UserId $BHR_workEmail -AccountEnabled:$BHR_AccountEnabled
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not activate the User Account at line 601. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        Write-Log -Message "Account $BHR_workEmail marked as Active in BambooHR but found Inactive in AzAD. Enabled AzAD account for sign-in. IF condition result is True at line 597." -Severity Warning
                        $error.Clear()
                    }                   
                }

                #Compare user employeeId with BambooHR and set it if not correct
                if ($BHR_employeeNumber -ne $azAD_employeeNumber) {
                    #Setting the Employee ID found in BHR to the user in AzAD
                    Update-MgUser -UserId $BHR_workEmail -EmployeeId $BHR_employeeNumber             
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the EmployeeID. Error on command at line 631. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        Write-Log -Message "The ID $BHR_employeeNumber has been set to $BHR_workEmail AzAD account. IF condition result is True at line 628." -Severity Warning
                        $error.Clear()
                    }
                }
            
                #Set Company name to "Company Name"
                if ($azAD_CompanyName -ne "Company Name") {
                    #Setting Comany Name as "Company Name" to the employee, if not already set
               
                    Update-MgUser -UserId $BHR_workEmail -CompanyName "Company Name"
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the Company Name of $BHR_workEmail. Error on line 662. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        Write-Log -Message "The $BHR_workEmail employee Company attribute has been set to: Company Name. IF condition result is True at line 658." -Severity Warning
                    }
                }

                #Set LastModified from BambooHR to ExtensionAttribute1 in AzAD
                if ($UPN_ExtensionAttribute1 -ne $BHR_lastChanged) {
                    #Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AzAD
                
                    #Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $BHR_lastChanged } -UserId $BHR_workEmail
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the ExtensionAttribute1. Error on line 692. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        Write-Log -Message "The $BHR_workEmail employee LastChanged attribute set to extensionAttribute1 as $BHR_lastChanged. On line 688." -Severity Warning
                    }
                }

                $error.clear()             
            }
   
            ########## Mechanism for name changing situations 
            if ($azAD_UPN_OBJdetails.id -ne $azAD_EID_OBJdetails.id -and $azAD_EID_OBJdetails.EmployeeID -eq $BHR_employeeNumber -and $historystatus -notlike "*suspended*") {
                $azAD_UPN = $azAD_EID_OBJdetails.UserPrincipalName
                $azAD_ObjectID = $azAD_EID_OBJdetails.id
                $azAD_workemail = $azAD_EID_OBJdetails.Mail
                $azAD_employeeNumber = $azAD_EID_OBJdetails.EmployeeID
                $azAD_displayname = $azAD_EID_OBJdetails.displayname
                $azAD_firstName = $azAD_EID_OBJdetails.GivenName
                $azAD_lastName = $azAD_EID_OBJdetails.Surname

                Write-Log -Message "Initiated name changing procedure at line 721 for AzAD user object found with the following attributes:
    Firstname: $azAD_firstName
    Lastname: $azAD_lastName
    Display Name: $azAD_displayname
    Work Email: $azAD_workemail
    UserPrincipalName: $azAD_UPN
    Employee Number: $azAD_employeeNumber" -Severity Information


                #######
            
                $error.Clear()
                #Set LastModified from BambooHR to ExtensionAttribute1 in AzAD
                if ($EID_ExtensionAttribute1 -ne $BHR_lastChanged) {
                    #Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AzAD
                    Write-Log -Message "The $BHR_workEmail employee LastChanged attribute set to extensionAttribute1 as $BHR_lastChanged. IF at line 744 returned TRUE." -Severity Information
                    Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $BHR_lastChanged } -UserId $azAD_ObjectID
                        
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the ExtensionAttribute1. Error on command at line 748. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        Write-Log -Message "ExtensionAttribute1 changed to: $BHR_lastChanged for employee $BHR_workEmail. On line 748." -Severity Warning
                    }
                }
                #Change last name in Azure
            
                if ($azAD_lastName -ne $BHR_lastName) {
                    Write-Log -Message "Changing the last name of $BHR_workEmail from $azAD_lastName to $BHR_lastName. IF condition returned true on line 775." -Severity Information
                    Update-MgUser -UserId $azAD_ObjectID -Surname $BHR_lastName
                        
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                               
                        Write-Log -Message "Could not change the Last Name. Error on line 778. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }
                    else {
                        Write-Log -Message "Successfully changed the last name of $BHR_workEmail from $azAD_lastName to $BHR_lastName on line 775." -Severity Warning
                    }
                }
            
                #Change First Name
                if ($azAD_firstName -ne $BHR_firstName) {
                    Update-MgUser -UserId $azAD_ObjectID -GivenName $BHR_firstName
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the First Name of $azAD_ObjectID. Error on line 807. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }#Change First Name error logging     
                    else {
                        Write-Log -Message "Successfully changed the first name of user with objID: $azAD_ObjectID. On line 807." -Severity Warning
                    }       
                }#Change first name script block closure
           
                #Change display name
                if ($azAD_displayname -ne $BHR_displayName) {
                    Update-MgUser -UserId $azAD_ObjectID -DisplayName $BHR_displayName

                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the Display Name. Error on line 835. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }#Change display name - Error logging
                    else {
                        Write-Log "The current Display Name: $azAD_displayname of $azAD_ObjectID has been changed to $BHR_displayName. If condition true on line 833." -Severity Warning
                    }        
                }#Change Display Name script block closure

                #Change Email Address
                if ($azAD_workemail -ne $BHR_workEmail) {
                    Update-MgUser -UserId $azAD_ObjectID -Mail $BHR_workEmail
                        
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the Email Address. Error on line 864. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }#Change Email Address error logging
                    else {
                        Write-Log "The current Email Address: $azAD_workemail of $azAD_ObjectID has been changed to $BHR_workEmail. If condition true on line 862." -Severity Warning
                    }             
                }

                #Change UserPrincipalName and send the details via email to the User
                if ($azAD_UPN -ne $BHR_workEmail) {

                    Update-MgUser -UserId $azAD_ObjectID -UserPrincipalName $BHR_workEmail
                       
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the UPN for $azAD_ObjectID. Error on line 893. Details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    } 
                    else {
                        Write-Log -Message "Changed the current UPN:$azAD_UPN of $azAD_ObjectID to $BHR_workEmail. IF Condition true on line 893." -Severity Warning
                        $params = @{
                            Message         = @{
                                Subject      = "Login details change for $BHR_displayName"
                                Body         = @{
                                    ContentType = "text"
                                    Content     = "Hello, `
`
As your name has been changed, we have also changed the login name for you. `
Your new login name is: $BHR_workEmail `
Please use the new login name when authenticating on DOMAIN devices and resources. `
`
Regards, `
Automated User Account Management Service"
                                }
                                ToRecipients = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $BHR_workEmail, "admin_monitor@domain.com"
				                                    }
                                    }
                                )
                            }
                            SaveToSentItems = "True"
                        }

                        Send-MgUserMail -BodyParameter $params -UserId Impersonated_Mailbox@domain.com
                    }
                }#Change UserPrincipalName and send the details via email to the User CLOSURE
            }#IF NAME CHANGE MECHANISM CLOSURE
            #######CREATE ACCOUNT FOR NEW EMPLOYEE#######   
            if ($azAD_UPN_OBJdetails.Capacity -eq 0 -and $azAD_EID_OBJdetails.Capacity -eq 0 -and $BHR_AccountEnabled -eq $True) {
                #Create AzAD account, as it doesn't have one, if user hire date is less than 14 days in the future, or is in the past
                Write-Log -Message "The Employee $BHR_workEmail does not have an AzAD account and Hire date is less than 14 days from present time or in the past." -Severity Information
        
                $PasswordProfile = @{
                    Password = (Get-RandomPassword 12)
                }
                $error.clear() 

                New-MgUser `
                    -Department $BHR_department `
                    -EmployeeId $BHR_employeeNumber `
                    -JobTitle $BHR_jobTitle `
                    -CompanyName "Company Name" `
                    -Surname $BHR_lastName `
                    -GivenName $BHR_firstName `
                    -DisplayName $BHR_displayName `
                    -AccountEnabled `
                    -Mail $BHR_workEmail `
                    -EmployeeHireDate $BHR_hireDate `
                    -UserPrincipalName $BHR_workEmail `
                    -PasswordProfile $PasswordProfile `
                    -MailNickname ($BHR_workEmail -replace "@domain.com", "") `
                    -UsageLocation RO `
                    -OnPremisesExtensionAttributes @{extensionAttribute1 = $BHR_lastChanged }
     


                if ($? -eq $True) {
                    Write-Log -Message "Account $BHR_workEmail successfully created. Please manually assign the required licenses and groups" -Severity Warning
                    $azAD_managerID = (Get-MgUser -UserId $BHR_supervisorEmail | Select-Object id).id
                    $NewManager = @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$azAD_managerID"
                    }
                    Start-Sleep -Seconds 8
                    Write-Log -Message "Setting the manager of the newly created user:$BHR_workEmail. Adding to BambooHR SAML enterprise app." -Severity Information
            
                    Set-MgUserManagerByRef -UserId $BHR_workEmail -BodyParameter $NewManager
            
                    #Assigning the user to BambooHR enterprise app
                    $uid = (get-mguser -UserId $BHR_workEmail | Select-Object ID).id
                    New-MgUserAppRoleAssignment -UserId $uid `
                        -PrincipalId $uid `
                        -ResourceId 6b419818-6e25-4f9c-9268-1f0d2ef78700 `
                        -AppRoleId a8972ac9-2341-4879-9028-2a9d979844a0

                    #Send mail with credentials of the newly created user
                    $pass = ($PasswordProfile.Values)
                    $params = @{
                        Message         = @{
                            Subject      = "User creation automation: $BHR_displayName"
                            Body         = @{
                                ContentType = "text"
                                Content     = "Hello, `
`
New employee user account created. Please manually assign a license to the user and add to all.employees group. `
User Principal Name: $BHR_workEmail `
Password:  $pass `
`
Regards, `
Automated User Account Management Service"
                            }
                            ToRecipients = @(
                                @{
                                    EmailAddress = @{
                                        Address = "admin_monitor@domain.com"
                                    }
                                }
                            )
                        }
                        SaveToSentItems = "True"
                    }

                    Send-MgUserMail -BodyParameter $params -UserId Impersonated_Mailbox@domain.com



                }#IF "User account creation succeded" closure
                else {
            
                    $creationerror = $error
                    $full_Error_Details = $creationerror | Select-Object *
                    $creationerror_exception = $creationerror.Exception.Message
                    $creationerror_category = $creationerror.CategoryInfo
                    $creationerror_errID = $creationerror.FullyQualifiedErrorId
                    $creationerror_stack = $creationerror.ScriptStackTrace    
          
          
                    Write-Log -Message "Account $BHR_workEmail creation failed. New-Mguser cmdlet at line 955 returned error. Full error details: `
$full_Error_Details" `
                        -Severity Error

                    $params = @{
                        Message         = @{
                            Subject      = "FAILURE: User creation automation $BHR_displayName"
                            Body         = @{
                                ContentType = "text"
                                Content     = "Hello, `
`
Account creation for user: $BHR_workEmail has failed. Please check the log: $log_filename for further details. Error message summary below. `
Error Message: `
$creationerror_exception `
Error Category: `
$creationerror_category `
Error ID: `
$creationerror_errID `
Stack: `
$creationerror_stack `
`
Regards, `
Automated User Account Management Service"
                            }
                            ToRecipients = @(
                                @{
                                    EmailAddress = @{
                                        Address = "admin_monitor@domain.com"
                                    }
                                }
                            )
                        }
                        SaveToSentItems = "True"
                    }#Send Mail Message parameters definition closure

                    Send-MgUserMail -BodyParameter $params -UserId Impersonated_Mailbox@domain.com

                }#else "User account creation failed" closure

            }#"CREATE ACCOUNT" Closure

                                                              
        }#If Hire Date is less than 14 days in the future or in the past closure
        else {
            #The user account does not need to be created as it does not satisfy the condition of the HireDate being 14 days or less in the future
            Write-Log -Message "The Employee $BHR_workEmail hire date is more than 14 days in the future. Will be created when HireDate is 14 days or less in the future." -Severity Information
        }#ELSE "Hire start date is more than 14 days in the future" closure

    }#Foreach User data in bamboo closure

}#RUNTIME BLOCK CLOSURE 
$runtime_seconds = $runtime.TotalSeconds
Write-Log -Message "Total time the script ran: $runtime_seconds" -Severity Information
Disconnect-MgGraph
Exit
#Script End
