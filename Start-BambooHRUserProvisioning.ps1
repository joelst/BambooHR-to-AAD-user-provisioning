#Requires -Module ExchangePowerShell,Microsoft.Graph.Users.Actions
<#
.SYNOPSIS
Script to synchronize employee information from BambooHR to Azure Active Directory. It does not support on premises Active Directory.

.DESCRIPTION
Extracts employee data from BambooHR and performs one of the following for each user extracted:

	1. Attribute corrections - if the user has an existing account , and is an active employee, and the last changed time in Azure AD differs from BambooHR, then this first block will compare each of the AzAD User object attributes with the data extracted from BHR and correct them if necessary
	2. Name change - If the user has an existing account, but the name does not match with the one from BHR, then, this block will run and correct the user Name, UPN,	emailaddress
	3. New employee, and there is no account in AzureAD for him, this script block will create a new user with the data extracted from BHR

.PARAMETER BambooHrApiKey
 Specifies the BambooHR API key

.PARAMETER AdminEmailAddress 
Specifies the email address to receive notifications

.PARAMETER CompanyName 
Specifies the BambooHR company name used in the URL

.PARAMETER TenantID 
Specifies the Microsoft tenant name (company.onmicrosoft.com)

.PARAMETER AADCertificateThumbprint 
Specifies the certificate thumbprint for the Azure AD client application.

.PARAMETER AzureClientAppId 
Specifies the Azure AD client application id

.NOTES
More documentation available in project README
#>

[CmdletBinding()]
param (
    [Parameter()]
    [String]
    $BambooHrApiKey,
    [Parameter()]
    [String]
    $AdminEmailAddress,
    [Parameter()]
    [string]
    $CompanyName,
    [Parameter()]
    [string]
    $TenantID,
    [Parameter()]
    [string]
    $AADCertificateThumbprint,
    [parameter()]
    [string]
    $AzureClientAppId
)

# Check if variables are blank and if there is a environment variable that should be applied for being used as an Azure Function
if ([string]::IsNullOrWhiteSpace($BambooHrApiKey) -and [string]::IsNullOrWhiteSpace($env:BambooHrApiKey)) {
   Write-Log "BambooHR API Key not defined" -Severity Error
   exit
}
elseif ([string]::IsNullOrWhiteSpace($AdminEmailAddress) -and (-not [string]::IsNullOrWhiteSpace($env:BambooHrApiKey))) {
    $BambooHrApiKey = $env:BambooHrApiKey
}

if ([string]::IsNullOrWhiteSpace($AdminEmailAddress) -and [string]::IsNullOrWhiteSpace($env:AdminEmailAddress)) {
    Write-Log "Admin Email Address not defined" -Severity Error
    exit
 }
 elseif ([string]::IsNullOrWhiteSpace($AdminEmailAddress) -and (-not [string]::IsNullOrWhiteSpace($env:AdminEmailAddress))) {
     $AdminEmailAddress = $env:AdminEmailAddress
 }

 if ([string]::IsNullOrWhiteSpace($CompanyName) -and [string]::IsNullOrWhiteSpace($env:CompanyName)) {
    Write-Log "Company name not defined" -Severity Error
    exit
 }
 elseif ([string]::IsNullOrWhiteSpace($CompanyName) -and (-not [string]::IsNullOrWhiteSpace($env:CompanyName))) {
    $CompanyName= $env:CompanyName
 }

 if ([string]::IsNullOrWhiteSpace($TenantID) -and [string]::IsNullOrWhiteSpace($env:TenantID)) {
    Write-Log "TenantID not defined" -Severity Error
    exit
 }
 elseif ([string]::IsNullOrWhiteSpace($TenantID) -and (-not [string]::IsNullOrWhiteSpace($env:TenantID))) {
     $TenantID = $env:TenantID
 }

 if ([string]::IsNullOrWhiteSpace($AADCertificateThumbprint) -and [string]::IsNullOrWhiteSpace($env:AADCertificateThumbprint)) {
    Write-Log "AAD Certificate Thumbprint not defined" -Severity Error
    exit
 }
 elseif ([string]::IsNullOrWhiteSpace($AADCertificateThumbprint) -and (-not [string]::IsNullOrWhiteSpace($env:AADCertificateThumbprint))) {
     $AADCertificateThumbprint = $env:AADCertificateThumbprint
 }

 if ([string]::IsNullOrWhiteSpace($AzureClientAppId) -and [string]::IsNullOrWhiteSpace($env:AzureClientAppId)) {
    Write-Log "Azure Client App Id not defined" -Severity Error
    exit
 }
 elseif ([string]::IsNullOrWhiteSpace($AzureClientAppId) -and (-not [string]::IsNullOrWhiteSpace($env:AzureClientAppId))) {
    $AzureClientAppId = $env:AzureClientAppId
 }

#Script start
$companyEmailDomain = $AdminEmailAddress.Split("@")[1]

$runtime = Measure-Command -Expression {
    # Provision users to AzureAD using the employee details from BambooHR

    # ERROR LOGGING FUNCTION
    $log_filename = "Log_" + (get-date -Format yyyy-MM-dd_HH-MM-s) + ".csv"
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
            Time     = (Get-Date -Format "yyyy/MM/dd HH:MM:s")
            Message  = $Message
            Severity = $Severity
        } | Export-Csv -Path $log_file -Append -NoTypeInformation
    } 

    function Get-NewPassword {
        <#
        .DESCRIPTION
            Generate a random password with the configured number of characters and special characters.
            Does not return characters that are commonly confused like 0 and O and 1 and l. Also removes characters that cause issues in PowerShell scripts.
        .EXAMPLE
            Get-NewPassword -PasswordLength 13 -SpecialChars 4
            Returns a password that is 13 characters long and includes 4 special characters.
        .NOTES
            Inspired by: http://blog.oddbit.com/2012/11/04/powershell-random-passwords/
        #>
        [CmdletBinding()]
        [OutputType([string])]
        param (
            [int]$PasswordLength = 12,
            # (REQUIRED)
            #
            # Specifies the total length of password to generate
    
            [int]$SpecialChars = 3
            # (REQUIRED)
            #
            # Specifies the number of special characters to include in the generated password.
        )
        $password = ""
    
        # punctuation options but doesn't include &,',",`,$,{,},[,],(),),|,;,, and a few others can break PowerShell or are difficult to read.
        $special = 43..46 + 94..95 + 126 + 33 + 35 + 61 + 63
        # Remove 0 and 1 because they can be confused with o,O,I,i,l
        $digits = 50..57
        # Remove O,o,i,I,l as these can be confused with other characters
        $letters = 65..72 + 74..78 + 80..90 + 97..104 + 106..107 + 109..110 + 112..122
        # Pick total minus the number of special chars of random letters and digits
        $chars = Get-Random -Count ($PasswordLength - $SpecialChars) -InputObject ($digits + $letters)
        # Pick the specified number of special characters
        $chars += Get-Random -Count $SpecialChars -InputObject ($special)
        # Mix up the chars so that the special char aren't just at the end and then convert each char number to the char and put in a string
        $password = Get-Random -Count $PasswordLength -InputObject ($chars) | ForEach-Object -Begin { $aa = $null } -Process { $aa += [char]$_ } -End { $aa }
    
        return $password
    }

    # Getting all users details from BambooHR and passing the extracted info to the variable $employees
    $headers = @{}
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Authorization", "asd $($BambooHRApiKey)")
    $error.clear()
    try {
        Invoke-RestMethod `
            -Uri "https://api.bamboohr.com/api/gateway.php/$($CompanyName)/v1/reports/custom?format=json&onlyCurrent=true" `
            -Method POST `
            -Headers $headers `
            -ContentType 'application/json' `
            -Body '{"fields":["status","hireDate","department","employeeNumber","firstName","lastName","displayName","jobTitle","supervisorEmail","workEmail","lastChanged","employmentHistoryStatus"]}' `
            -OutVariable response 

    } 
    catch {
        # If error returned, the API call to BambooHR failed and no usable employee data has been returned, write to log file and exit script
        
        $BHR_api_call_err = $_
        $BHR_api_call_err_message = $BHR_api_call_err.Exception.Message
        $BHR_api_call_err_category = $BHR_api_call_err.CategoryInfo.Category
        $BHR_api_call_err_stack = $BHR_api_call_err.ScriptStackTrace
        
        Write-log -Message "API call to return user information from BambooHR has failed. Script terminated. Following are the error details. `
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
Automated User Account Management"
                }
                ToRecipients = @(
                    @{
				                    EmailAddress = @{
                            Address = $AdminEmailAddress
                        }
                    }
                )
            }
            SaveToSentItems = "True"
        }

        Send-MgUserMail -BodyParameter $params -UserId "Impersonated_Mailbox@$($companyEmailDomain)"

        Exit
    }

    # If no error returned, it means that the script was not interrupted by the "Exit" command within the "Catch" block. Write info below to log file and continue
    Write-Log -Message "Successfully extracted employee information from BambooHR. 'Try' did not generate errors" -Severity Information
  
    # Saving only the employee data to $employees variable and eliminate $response variable to save memory
    $employees = $response.employees
    Remove-Variable -Name response

    # Connect to AzureAD using PS Graph Module, authenticating as the configured service principal for this operation, with certificate auth
    $error.Clear()

    Connect-MgGraph -TenantId $TenantID -CertificateThumbprint $AADCertificateThumbprint -ClientId $AzureClientAppId

    if ($?) {
        # If no error returned, write to log file and continue
        Write-Log -Message "Successfully connected to $TenantId. Command did not generate errors." -Severity Information
    }
    else {

        # If error returned, write to log file and exit script
        $mgconnerror = $error
        $mgerr_exception = $mgconnerror.Exception
        $mgerr_category = $mgconnerror.CategoryInfo
        $mgerr_errID = $mgconnerror.FullyQualifiedErrorId
        $mgerr_stack = $mgconnerror.ScriptStackTrace
        Write-Log -Message "Connection to tenant has failed. Script terminated. Below are the error details.`
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
                Subject      = "User creation automation: MGGraph Azure AD connection failed"
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
Automated User Account Management"
                }
                ToRecipients = @(
                    @{
				                    EmailAddress = @{
                            Address = $AdminEmailAddress
                        }
                    }
                )
            }
            SaveToSentItems = "True"
        }

        Send-MgUserMail -BodyParameter $params -UserId "Impersonated_Mailbox@$($companyEmailDomain)"

        Disconnect-MgGraph
        Exit
    }
        
    Write-Log -Message "Entering Foreach loop to check the attributes of each user at a time. Foreach loop initiated." -Severity Information
    
    $employees | ForEach-Object {
        $error.Clear()

        # On each loop, pass all employee data from BambooHR to variables, to be compared one by one with the user data from Azure AD and set them, if necessary
        $BHR_lastChanged = $_.lastChanged
        # Hire date as listed in Bamboo HR
        $BHR_hireDate = $_.hireDate
        # Employee number as listed in Bamboo HR
        $BHR_employeeNumber = $_.employeeNumber
        # Job title as listed in Bamboo HR
        $BHR_jobTitle = $_.jobTitle
        # Department as listed in Bamboo HR
        $BHR_department = $_.department
        # Supervisor email address as listed in Bamboo HR
        $BHR_supervisorEmail = $_.supervisorEmail
        # Work email address as listedin Bamboo HR
        $BHR_workEmail = $_.workEmail
        # Current status of the employee: Active, Terminated and if contains "Suspended" is in "maternity leave"
        $BHR_EmploymentStatus = $_.employmentHistoryStatus
        # Translating user "status" from BambooHR to boolean, to match and compare with the AzureAD user account status
        $BHR_status = $_.status
        if ($BHR_status -eq "Inactive")
        { $BHR_AccountEnabled = $False }
        if ($BHR_status -eq "Active")
        { $BHR_AccountEnabled = $True }

        # Normalizing user names, eliminating language specific characters
        # Last name of employee in Bamboo HR
        $BHR_firstName = $_.firstName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t"
        # First name of employee in Bamboo HR
        $BHR_lastName = $_.lastName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t"
        # The Display Name of the user in BambooHR
        $BHR_displayName = $_.displayName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t"
        $BHR_firstName = (get-culture).textinfo.ToTitleCase($BHR_firstName)
        $BHR_lastName = (get-culture).textinfo.ToTitleCase($BHR_lastName)
        $BHR_displayName = (get-culture).textinfo.ToTitleCase($BHR_displayName)
        Write-Log -Message "Eliminated language specific characters from names and capitalize first letters." -Severity Information

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

        $current_date = (get-date)
        if ((([datetime]$BHR_hireDate).AddDays(-14)) -lt $current_date) {

            $error.clear()
    
            $azAD_UPN_OBJdetails = $null
            $azAD_EID_OBJdetails = $null

            # Check if the user exists in AzAD and If there is an account with the EmployeeID of the user checked in the current loop
            Write-Log -Message "Checking if $BHR_workEmail has a valid and matching AzureAD account." -Severity Information
            Get-MgUser -UserId $BHR_workEmail -OutVariable azAD_UPN_OBJdetails -Property id, userprincipalname, Department, EmployeeID, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, EmployeeHireDate, OnPremisesExtensionAttributes -ExpandProperty manager
            Get-MgUser -Filter "employeeID eq '$BHR_employeenumber'" -OutVariable azAD_EID_OBJdetails -Property employeeid, userprincipalname, Department, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, EmployeeHireDate, OnPremisesExtensionAttributes  -ExpandProperty manager
            $error.clear()
            $UPN_ExtensionAttribute1 = ($azAD_UPN_OBJdetails | Select-Object @{Name = 'ExtensionAttribute1'; Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } }).ExtensionAttribute1
            $EID_ExtensionAttribute1 = ($azAD_EID_OBJdetails | Select-Object @{Name = 'ExtensionAttribute1'; Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } }).ExtensionAttribute1
   
            <# IF: empID of object returned by UPN or by empID is equal, and if object ID is the same as object ID of object returned by UPN and EmpID and
    if UPN = workemail from bamboo  AND if the last changed date from BambooHR is NOT equal to the last changed date saved in ExtensionAttribute1 in AzAD, check each attribute and set
    them correctly, according to BambooHR
    #>
   
            if ($azAD_EID_OBJdetails.EmployeeId -eq $azAD_UPN_OBJdetails.EmployeeId -and `
                    $azAD_EID_OBJdetails.UserPrincipalName -eq $azAD_UPN_OBJdetails.UserPrincipalName -eq $BHR_workEmail -and `
                    $azAD_UPN_OBJdetails.id -eq $azAD_EID_OBJdetails.id -and `
                    $BHR_lastChanged -ne $upn_ExtensionAttribute1 -and `
                    $azAD_EID_OBJdetails.Capacity -ne 0 -and `
                    $azAD_UPN_OBJdetails.Capacity -ne 0 -and `
                    $BHR_EmploymentStatus -notlike "*suspended*") { 
                Write-Log  -Message "$BHR_workEmail is a valid AzureAD Account, with matching EmployeeID and UPN from AzureAD to BambooHR, but different last modified date. If statement passed as true." -Severity Warning
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

                        Write-Log -Message "Could not change the JobTitle of $BHR_workEmail. Command returned error. Details below. `
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
Set the JobTitle found on BambooHR to the Azure User Object. IF condition result is True" -Severity Warning
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

                        Write-Log -Message "Could not change the Department of $BHR_workEmail. Command returned error. Details below. `
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
Setting the Department found on BambooHR to the Azure User Object. IF condition result is True" -Severity Information
                    }
                }
        
                # Checking the manager if correctly set, if not, configure the manager as set in BambooHR
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

                        Write-Log -Message "Could not change the Manager of $BHR_workEmail. Command returned error. Details below. `
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
Setting the Manager found on BambooHR to the Azure User Object. IF condition result is True" -Severity Warning
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

                        Write-Log -Message "Could not change the Employee Hire Date. Command returnet error. Details below. `
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
                        Write-Log -Message "Setting the Hire date of $BHR_workEmail to $BHR_hireDate. IF condition result is True" -Severity Warning
                    }
                }

                # Check if user is active in BambooHR, and set the status of the account as it is in BambooHR (active or inactive)
           
                if ($BHR_AccountEnabled -eq $False -and $BHR_EmploymentStatus -eq "Terminated") {

                    #As the account is marked "Inactive" in BHR and "Active" in AzAD, block sign-in, revoke sessions, change pass, remove auth methods
                    $error.clear()
                    Update-MgUser -UserId $BHR_workEmail -AccountEnabled:$BHR_AccountEnabled
                    Revoke-MgUserSign -UserId $BHR_workEmail

                    #Change to a random pass
                    $params = @{
                        PasswordProfile = @{
                            ForceChangePasswordNextSignIn = $true
                            Password                      = (Get-NewPassword)
                        }
                    }
                    Update-MgUser -UserId $BHR_workEmail -BodyParameter $params
                    Get-MgUserMemberOf -UserId $BHR_workEmail
                    # Convert mailbox to shared
                    Connect-ExchangeOnline -CertificateThumbprint $AADCertificateThumbprint -AppId $AzureClientAppId -Organization $TenantId 
                    Set-Mailbox -Identity $BHR_workEmail -Type Shared
                    Disconnect-ExchangeOnline -Confirm:$False
                    # Remove Licenses
                    Get-MgUserLicenseDetail -UserId $BHR_workEmail | ForEach-Object { Set-MgUserLicense -UserId $BHR_workEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }
                    Get-MgUserMemberOf -UserId $BHR_workEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $azAD_UPN_OBJdetails.id } #Remove Groups
                    $methodID = Get-MgUserAuthenticationMethod -UserId $BHR_workEmail | Select-Object id 
                    $methodsdata = Get-MgUserAuthenticationMethod -UserId $BHR_workEmail | Select-Object -ExpandProperty AdditionalProperties
                    $methods_count = ($methodID | Measure-Object | Select-Object count).count

                    # Pass through each authentication method and remove them
                    $error.Clear() 
                    for ($i = 0 ; $i -lt $methods_count ; $i++) {
       
                        if ((($methodsdata[$i]).Values) -like "*phoneAuthenticationMethod*") { Remove-MgUserAuthenticationPhoneMethod -UserId $BHR_workEmail -PhoneAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed phone auth method for $BHR_workEmail." -Severity Warning }
                        if ((($methodsdata[$i]).Values) -like "*microsoftAuthenticatorAuthenticationMethod*") { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $BHR_workEmail -MicrosoftAuthenticatorAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed auth app method for $BHR_workEmail." -Severity Warning }
                        if ((($methodsdata[$i]).Values) -like "*windowsHelloForBusinessAuthenticationMethod*") { Remove-MgUserAuthenticationFido2Method -UserId $BHR_workEmail -Fido2AuthenticationMethodId ($methodID[$i]).id ; Write-Log -Message "Removed PIN auth method for $BHR_workEmail." -Severity Warning }
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
                        Write-Log -Message "Account $BHR_workEmail marked as inactive in BambooHR but found active in Azure AD. Disabled AzAD account, revoked sessions and removing auth methods. IF condition result is True" -Severity Warning              
                        $error.Clear()
                    }

                }
          
                if ($BHR_AccountEnabled -eq $True -and $azAD_status -eq $False) {
                    # As the account is marked "Active" in BHR and "Inactive" in AzAD, enable the AzAD account
                
                    Update-MgUser -UserId $BHR_workEmail -AccountEnabled:$BHR_AccountEnabled
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not activate the user account. Details below. `
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
                        Write-Log -Message "Account $BHR_workEmail marked as Active in BambooHR but found Inactive in AzAD. Enabled AzAD account for sign-in. IF condition result is True." -Severity Warning
                        $error.Clear()
                    }                   
                }

                # Compare user employeeId with BambooHR and set it if not correct
                if ($BHR_employeeNumber -ne $azAD_employeeNumber) {
                    # Setting the Employee ID found in BHR to the user in AzAD
                    Update-MgUser -UserId $BHR_workEmail -EmployeeId $BHR_employeeNumber             
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the EmployeeID. Error on command. Details below. `
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
                        Write-Log -Message "The ID $BHR_employeeNumber has been set to $BHR_workEmail AzAD account. IF condition result is True." -Severity Warning
                        $error.Clear()
                    }
                }
            
                # Set Company name to "Company Name"
                if ($azAD_CompanyName -ne "Company Name") {
                    # Setting Comany Name as "Company Name" to the employee, if not already set
               
                    Update-MgUser -UserId $BHR_workEmail -CompanyName "Company Name"
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the Company Name of $BHR_workEmail. Error details below. `
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
                        Write-Log -Message "The $BHR_workEmail employee Company attribute has been set to: Company Name. IF condition result is True." -Severity Warning
                    }
                }

                # Set LastModified from BambooHR to ExtensionAttribute1 in AzAD
                if ($UPN_ExtensionAttribute1 -ne $BHR_lastChanged) {
                    # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AzAD
                
                    Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $BHR_lastChanged } -UserId $BHR_workEmail
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the ExtensionAttribute1. Error details below. `
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
                        Write-Log -Message "The $BHR_workEmail employee LastChanged attribute set to extensionAttribute1 as $BHR_lastChanged." -Severity Warning
                    }
                }

                $error.clear()             
            }
            
            # IF - Attributes correction (excluding name/firstname/displayname/emailaddress/UPN)
   
            # Handle name change situations 
            if ($azAD_UPN_OBJdetails.id -ne $azAD_EID_OBJdetails.id -and $azAD_EID_OBJdetails.EmployeeID -eq $BHR_employeeNumber -and $historystatus -notlike "*suspended*") {
                $azAD_UPN = $azAD_EID_OBJdetails.UserPrincipalName
                $azAD_ObjectID = $azAD_EID_OBJdetails.id
                $azAD_workemail = $azAD_EID_OBJdetails.Mail
                $azAD_employeeNumber = $azAD_EID_OBJdetails.EmployeeID
                $azAD_displayname = $azAD_EID_OBJdetails.displayname
                $azAD_firstName = $azAD_EID_OBJdetails.GivenName
                $azAD_lastName = $azAD_EID_OBJdetails.Surname

                Write-Log -Message "Initiated name changing procedure for AzAD user object found with the following attributes:
    Firstname: $azAD_firstName
    Lastname: $azAD_lastName
    Display Name: $azAD_displayname
    Work Email: $azAD_workemail
    UserPrincipalName: $azAD_UPN
    Employee Number: $azAD_employeeNumber" -Severity Information
           
                $error.Clear()
                # Set LastModified from BambooHR to ExtensionAttribute1 in AzAD
                if ($EID_ExtensionAttribute1 -ne $BHR_lastChanged) {
                    # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AzAD
                    Write-Log -Message "The $BHR_workEmail employee LastChanged attribute set to extensionAttribute1 as $BHR_lastChanged. If returned True." -Severity Information
                    Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $BHR_lastChanged } -UserId $azAD_ObjectID
                        
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace

                        Write-Log -Message "Could not change the ExtensionAttribute1. Error on command. Details below. `
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
                        Write-Log -Message "ExtensionAttribute1 changed to: $BHR_lastChanged for employee $BHR_workEmail." -Severity Warning
                    }
                }
                # Change last name in Azure
            
                if ($azAD_lastName -ne $BHR_lastName) {
                    Write-Log -Message "Changing the last name of $BHR_workEmail from $azAD_lastName to $BHR_lastName. IF condition returned true." -Severity Information
                    Update-MgUser -UserId $azAD_ObjectID -Surname $BHR_lastName
                        
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                               
                        Write-Log -Message "Could not change the Last Name. Error details below. `
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
                        Write-Log -Message "Successfully changed the last name of $BHR_workEmail from $azAD_lastName to $BHR_lastName." -Severity Warning
                    }
                }
            
                # Change First Name
                if ($azAD_firstName -ne $BHR_firstName) {
                    Update-MgUser -UserId $azAD_ObjectID -GivenName $BHR_firstName
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the First Name of $azAD_ObjectID. Error details below. `
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
                        Write-Log -Message "Successfully changed the first name of user with objID: $azAD_ObjectID." -Severity Warning
                    }       
                }# Change first name script block closure
           
                # Change display name
                if ($azAD_displayname -ne $BHR_displayName) {
                    Update-MgUser -UserId $azAD_ObjectID -DisplayName $BHR_displayName

                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the Display Name. Error details below. `
                                Exception: `
                                $err_msg `
                                Target object: `
                                $err_TargetObject `
                                Details: `
                                $err_details `
                                StackTrace: `
                                $err_trace" -Severity Error
                        $error.Clear()
                    }# Change display name - Error logging
                    else {
                        Write-Log "The current Display Name: $azAD_displayname of $azAD_ObjectID has been changed to $BHR_displayName. If condition true." -Severity Warning
                    }        
                }
                # Change Display Name script block closure

                # Change Email Address
                if ($azAD_workemail -ne $BHR_workEmail) {
                    Update-MgUser -UserId $azAD_ObjectID -Mail $BHR_workEmail
                        
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the Email Address. Error details below. `
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
                        # Change Email Address error logging
                        Write-Log "The current Email Address: $azAD_workemail of $azAD_ObjectID has been changed to $BHR_workEmail. If condition true." -Severity Warning
                    }             
                }

                # Change UserPrincipalName and send the details via email to the User
                if ($azAD_UPN -ne $BHR_workEmail) {

                    Update-MgUser -UserId $azAD_ObjectID -UserPrincipalName $BHR_workEmail
                       
                    if (!$?) {
                        $err_msg = $Error.exception
                        $err_TargetObject = $error.TargetObject
                        $err_details = $error.ErrorDetails
                        $err_trace = $error.ScriptStackTrace
                             
                        Write-Log -Message "Could not change the UPN for $azAD_ObjectID. Error details below. `
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
                        Write-Log -Message "Changed the current UPN:$azAD_UPN of $azAD_ObjectID to $BHR_workEmail. IF Condition true." -Severity Warning
                        $params = @{
                            Message         = @{
                                Subject      = "Login details change for $BHR_displayName"
                                Body         = @{
                                    ContentType = "text"
                                    Content     = "Hello, `
`
As your name has been changed, we have also changed the login name for you. `
Your new login name is: $BHR_workEmail `
Please use the new login name when authenticating . `
`
Regards, `
Automated User Account Management Service"
                                }
                                ToRecipients = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $BHR_workEmail, $AdminEmailAddress
				                                    }
                                    }
                                )
                            }
                            SaveToSentItems = "True"
                        }

                        Send-MgUserMail -BodyParameter $params -UserId "Impersonated_Mailbox@$($companyEmailDomain)"
                    }
                }
            }
            
            # CREATE ACCOUNT FOR NEW EMPLOYEE   
            if ($azAD_UPN_OBJdetails.Capacity -eq 0 -and $azAD_EID_OBJdetails.Capacity -eq 0 -and $BHR_AccountEnabled -eq $True) {
                # Create AzAD account, as it doesn't have one, if user hire date is less than 14 days in the future, or is in the past
                Write-Log -Message "The Employee $BHR_workEmail does not have an AzAD account and Hire date is less than 14 days from present time or in the past." -Severity Information
        
                $PasswordProfile = @{
                    Password = (Get-NewPassword)
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
                    -MailNickname ($BHR_workEmail -replace "@$($companyEmailDomain)", "") `
                    -UsageLocation RO `
                    -OnPremisesExtensionAttributes @{extensionAttribute1 = $BHR_lastChanged 
                    }

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
                                        Address = $AdminEmailAddress
                                    }
                                }
                            )
                        }
                        SaveToSentItems = "True"
                    }

                    Send-MgUserMail -BodyParameter $params -UserId "Impersonated_Mailbox@$($companyEmailDomain)"

                }
                else {
                    # If "User account creation succeded" closure
                    $creationerror = $error
                    $full_Error_Details = $creationerror | Select-Object *
                    $creationerror_exception = $creationerror.Exception.Message
                    $creationerror_category = $creationerror.CategoryInfo
                    $creationerror_errID = $creationerror.FullyQualifiedErrorId
                    $creationerror_stack = $creationerror.ScriptStackTrace    

                    Write-Log -Message "Account $BHR_workEmail creation failed. New-Mguser cmdlet returned error. Full error details: `
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
                                        Address = $AdminEmailAddress
                                    }
                                }
                            )
                        }
                        SaveToSentItems = "True"
                    }

                    # Send Mail Message parameters definition closure
                    Send-MgUserMail -BodyParameter $params -UserId "Impersonated_Mailbox@$($companyEmailDomain)"

                }# else "User account creation failed" closure

            }

                                                              
        }
        else {
            # If Hire Date is less than 14 days in the future or in the past closure
            # The user account does not need to be created as it does not satisfy the condition of the HireDate being 14 days or less in the future
            Write-Log -Message "The Employee $BHR_workEmail hire date is more than 14 days in the future. Will be created when HireDate is 14 days or less in the future." -Severity Information
        }


    }

}
# RUNTIME BLOCK CLOSURE 
$runtime_seconds = $runtime.TotalSeconds
Write-Log -Message "Total time the script ran: $runtime_seconds" -Severity Information
Disconnect-MgGraph
Exit
#Script End
