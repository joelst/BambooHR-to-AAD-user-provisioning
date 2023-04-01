#Requires -Module ExchangeOnlineManagement,Microsoft.Graph
<#

IMPORTANT: This is a sample solution and should be used by those comfortable testing and retesting and validating before thinking about using it in production. 
All content is provided AS IS with no guarantees or assumptions of quality of functionality. 

If you are using employee information there is much that can go wrong! 

You are responsible for complying with all applicable laws and regulations. 
With great power comes great responsibility. 
Friends don't let friends run untested directory scripts in production.
Don't take any wooden nickels


.SYNOPSIS
Script to synchronize employee information from BambooHR to Azure Active Directory. It does not support on premises Active Directory.

.DESCRIPTION
Extracts employee data from BambooHR and performs one of the following for each user extracted:

	1. Attribute corrections - if the user has an existing account, is an active employee, and the last changed time in Azure AD differs from BambooHR, then this first block will compare each of the AAD User object attributes with the data extracted from BHR and correct them if necessary
	2. Name change - If the user has an existing account, but the name does not match with the one from BHR, then, this block will run and correct the user Name, UPN,	emailaddress
	3. New employee, and there is no account in AAD for him, this script block will create a new user with the data extracted from BHR

.PARAMETER BambooHrApiKey
 Specifies the BambooHR API key as a string. It will be converted to the proper format.

.PARAMETER AdminEmailAddress 
Specifies the email address to receive notifications

.PARAMETER CompanyName 
Specifies the company name to use in the employee information

.PARAMETER BHRCompanyName 
Specifies the BambooHR company name used in the URL

.PARAMETER TenantID 
Specifies the Microsoft tenant name (company.onmicrosoft.com)

.PARAMETER AADCertificateThumbprint 
Specifies the certificate thumbprint for the Azure AD client application.

.PARAMETER AzureClientAppId 
Specifies the Azure AD client application id

.PARAMETER LogPath
Location to save logs

.PARAMETER UsageLocation
A two letter country code (ISO standard 3166) to set AAD usage location. 
Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries. Examples include: US, JP, and GB.

.PARAMETER TestOnly
Specify when you do not want to make any changes.

.PARAMETER CurrentOnly
Specify to only pull current employees from BambooHR. Default is to retrieve future employees.

.PARAMETER NotificationEmailAddress
Specifies an additional email address to send any notification emails to. 

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
    $BHRCompanyName,
    [Parameter()]
    [string]
    $CompanyName,
    [Parameter()]
    [string]
    $TenantId,
    [Parameter()]
    [string]
    $AADCertificateThumbprint,
    [parameter()]
    [string]
    $AzureClientAppId,
    [parameter()]
    [string]
    $LogPath = (Get-Location),
    [parameter()]
    [string]
    $UsageLocation = "US",
    [parameter()]
    [switch]
    $TestOnly,
    [parameter()]
    [switch]
    $CurrentOnly,
    [parameter()]
    [int]
    $DaysAhead = 7,
    [parameter()]
    [string]
    $NotificationEmailAddress,
    [parameter()]
    [string]
    $EmailSignature = "`n Regards, `n`n $CompanyName Automated User Management `n`n`nFor additional information, please create a service request.`n"
)

# ERROR LOGGING FUNCTION
$logFileName = "BhrAadSync-" + (Get-Date -Format yyyyMMdd-HHmm) + ".csv"
$logFilePath = Join-Path $LogPath $logFileName
$Script:logContent = ""

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Debug', 'Information', 'Warning', 'Error', 'Test')]
        [string]
        $Severity = "Information"
    )
    
    [pscustomobject]@{
        Time     = (Get-Date -Format "yyyy/MM/dd HH:mm:ss")
        Message  = ($Message.Replace("`n", '').Replace("`t", '').Replace("``", ''))
        Severity = $Severity
    } | Export-Csv -Path $logFilePath -Append -NoTypeInformation -Force

    switch ($Severity) {
        Debug { 
            Write-Verbose $Message
        }
        Warning { 
            Write-Warning $Message
        }
        Error { 
            Write-Error $Message
            $Script:logContent += "`n$Message`n"
        }
        Information { 
            Write-Host $Message
            $Script:logContent += "`n$Message`n"
        }
        Test { 
            Write-Host " [TestOnly] $Message" -ForegroundColor Green
        } 
        Default { 
            Write-Output $Message
        }
    }
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
        [int]$PasswordLength = 14,
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

# Used to simplify logic of creating or updating later.
$noAADAccount = $true

# Check if variables are not set. If there is an environment variable, set its value to the variable. Used as an Azure Function
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

if ([string]::IsNullOrWhiteSpace($BHRCompanyName) -and [string]::IsNullOrWhiteSpace($env:BHRCompanyName)) {
    Write-Log "BambooHR company name not defined" -Severity Error
    exit
}
elseif ([string]::IsNullOrWhiteSpace($BHRCompanyName) -and (-not [string]::IsNullOrWhiteSpace($env:BHRCompanyName))) {
    $BHRCompanyName = $env:BHRCompanyName
}

if ([string]::IsNullOrWhiteSpace($CompanyName) -and [string]::IsNullOrWhiteSpace($env:CompanyName)) {
    Write-Log "Company name not defined" -Severity Error
    exit
}
elseif ([string]::IsNullOrWhiteSpace($CompanyName) -and (-not [string]::IsNullOrWhiteSpace($env:CompanyName))) {
    $CompanyName = $env:CompanyName
}


if ([string]::IsNullOrWhiteSpace($TenantID) -and [string]::IsNullOrWhiteSpace($env:TenantID)) {
    Write-Log "TenantId not defined" -Severity Error
    exit
}
elseif ([string]::IsNullOrWhiteSpace($TenantID) -and (-not [string]::IsNullOrWhiteSpace($env:TenantID))) {
    $TenantID = $env:TenantID
    $env:AZURE_CLIENT_ID = $TenantId
}

if ([string]::IsNullOrWhiteSpace($AADCertificateThumbprint) -and [string]::IsNullOrWhiteSpace($env:AADCertificateThumbprint)) {
    Write-Log "AAD Certificate thumbprint not defined" -Severity Error
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
    $env:AZURE_CLIENT_ID = $AzureClientAppId
}

#$graphClientAppId = $AzureClientAppId.Replace("-", "")
#$aadCustomDataExtensionName = "extension_$($graphClientAppId)_bhrLastUpdated" 
$companyEmailDomain = $AdminEmailAddress.Split("@")[1]
$bhrRootUri = "https://api.bamboohr.com/api/gateway.php/$($BHRCompanyName)/v1"
$bhrReportsUri = $bhrRootUri

# If you only want to retrieve the current employees, not future ones use -CurrentOnly.
if ($CurrentOnly.IsPresent) {
    $bhrReportsUri = "$($bhrRootUri)/reports/custom?format=json&onlyCurrent=true"
}
else {
    $bhrReportsUri = "$($bhrRootUri)/reports/custom?format=json&onlyCurrent=false"
}

$runtime = Measure-Command -Expression {
    Write-Log -Message "Starting BambooHR to AAD synchronization at $(Get-Date)" -Severity Information
    # Provision users to AAD using the employee details from BambooHR
    Write-Log -Message "Executing Connect-MgGraph -TenantId $TenantID  ..." -Severity Debug
    Connect-MgGraph -TenantId $TenantID -CertificateThumbprint $AADCertificateThumbprint -ClientId $AzureClientAppId
    # Getting all users details from BambooHR and passing the extracted info to the variable $employees
    $headers = @{}
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Authorization", "Basic $([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($BambooHrApiKey):x")))")
    $error.clear()
    try {
        Invoke-RestMethod -Uri $bhrReportsUri -Method POST -Headers $headers -ContentType 'application/json' `
            -Body '{"fields":["status","hireDate","department","employeeNumber","firstName","lastName","displayName","jobTitle","supervisorEmail","workEmail","lastChanged","employmentHistoryStatus","bestEmail","location","workPhone","preferredName","homeEmail"]}' `
            -OutVariable response 

    } 
    catch {
        # If error returned, the API call to BambooHR failed and no usable employee data has been returned, write to log file and exit script
             
        Write-log -Message "Error calling BambooHr API for user information. `nEXCEPTION MESSAGE: $($_.Exception.Message) `n CATEGORY: $($_.CategoryInfo.Category) `n SCRIPT STACK TRACE: $($_.ScriptStackTrace)" -Severity Error
        #Send email alert with the generated error
        $params = @{
            Message         = @{
                Subject      = "User creation automation: Connection to BambooHR has failed"
                Body         = @{
                    ContentType = "text"
                    Content     = "Hello, `nConnection to BambooHR endpoint has failed. `n EXCEPTION MESSAGE: $($_.Exception.Message) `nCATEGORY: $($_.CategoryInfo.Category) `n SCRIPT STACK TRACE: $($_.ScriptStackTrace) `n $EmailSignature"
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

        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress

        Exit
    }

    # If no error returned, it means that the script was not interrupted by the "Exit" command within the "Catch" block. Write info below to log file and continue
    Write-Log -Message "Successfully extracted employee information from BambooHR." -Severity Information
  
    # Saving only the employee data to $employees variable and eliminate $response variable to save memory
    $employees = $response.employees
    $response = $null

    # Connect to AAD using PS Graph Module, authenticating as the configured service principal for this operation, with certificate auth
    $error.Clear()

    if ($?) {
        # If no error returned, write to log file and continue
        Write-Log -Message "Successfully connected to AAD: $TenantId." -Severity Information
    }
    else {

        # If error returned, write to log file and exit script
        Write-Log -Message "Connection to AAD failed.`n EXCEPTION: $($error.Exception) `n CATEGORY: $($error.CategoryInfo) `n ERROR ID: $($error.FullyQualifiedErrorId) `n SCRIPT STACK TRACE: $($error.ScriptStackTrace)" -Severity Error

        # Send email alert with the generated error 
        $params = @{
            Message         = @{
                Subject      = "User creation automation: MGGraph Azure AD connection failed"
                Body         = @{
                    ContentType = "text"
                    Content     = "Hello, `n`nAAD connection failed.`nEXCEPTION: $($error.Exception) `n CATEGORY:$($error.CategoryInfo) `n ERROR ID: $($error.FullyQualifiedErrorId) `nSCRIPT STACK TRACE: $mgErrStack `n $EmailSignature"
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

        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress

        Disconnect-MgGraph
        Exit
    }
        
    Write-Log -Message "Looping through $($employees.Count) users." -Severity Debug
    Write-Log -Message "Removing employee records that do not have a company email address of $companyEmailDomain" -Severity Debug
    
    # Only select employees with a company email.
    $employees | Sort-Object -Property LastName | Where-Object { $_.workEmail -like "*$companyEmailDomain" } | ForEach-Object {
        $error.Clear()
        # On each loop, pass all employee data from BambooHR to variables, to be compared one by one with the user data from Azure AD and set them, if necessary
        $bhrlastChanged = "$($_.lastChanged)"
        # Hire date as listed in Bamboo HR
        $bhrHireDate = "$($_.hireDate)"
        # Employee number as listed in Bamboo HR
        $bhremployeeNumber = "$($_.employeeNumber)"
        # Job title as listed in Bamboo HR
        $bhrJobTitle = "$($_.jobTitle)"
        # Department as listed in Bamboo HR
        $bhrDepartment = "$($_.department)"
        # Supervisor email address as listed in Bamboo HR
        $bhrSupervisorEmail = "$($_.supervisorEmail)"
        # Work email address as listedin Bamboo HR
        $bhrWorkEmail = "$($_.workEmail)"
        # Current status of the employee: Active, Terminated and if contains "Suspended" is in "maternity leave"
        $bhrEmploymentStatus = "$($_.employmentHistoryStatus)"
        # Translating user "status" from BambooHR to boolean, to match and compare with the AAD user account status
        $bhrStatus = "$($_.status)"
        if ($bhrStatus -eq "Inactive")
        { $bhrAccountEnabled = $False }
        if ($bhrStatus -eq "Active")
        { $bhrAccountEnabled = $True }
        $bhrOfficeLocation = "$($_.location)"
        $bhrPreferredName = "$($_.preferredName)"
        $bhrWorkPhone = "$($_.workPhone)"
        $bhrBestEmail = "$($_.bestEmail)"
        $bhrFirstName = $_.firstName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t" -creplace "  ", " "
        $bhrFirstName = (get-culture).textinfo.ToTitleCase($bhrFirstName).Trim()
        # First name of employee in Bamboo HR
        $bhrLastName = $_.lastName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t" -creplace "  ", " "
        $bhrFirstName = (get-culture).textinfo.ToTitleCase($bhrFirstName).Trim()
        # The Display Name of the user in BambooHR
        $bhrDisplayName = $_.displayName -creplace "Ă", "A" -creplace "ă", "a" -creplace "â", "a" -creplace "Â", "A" -creplace "Î", "I" -creplace "î", "i" -creplace "Ș", "S" -creplace "ș", "s" -creplace "Ț", "T" -creplace "ț", "t" -creplace "  ", " "
        $bhrFirstName = (get-culture).textinfo.ToTitleCase($bhrFirstName).Trim()
        $bhrLastName = (get-culture).textinfo.ToTitleCase($bhrLastName).Trim()
        $bhrDisplayName = (get-culture).textinfo.ToTitleCase($bhrDisplayName).Trim()
        $bhrHomeEmail = "$($homeEmail)"
 
        if ($bhrPreferredName -ne $bhrFirstName -and [string]::IsNullorWhitespace($bhrPreferredName) -eq $false) {
            Write-Log -Message "User preferred first name of $bhrPreferredName instead of legal name $bhrFirstName" -Severity Debug
            $bhrFirstName = $bhrPreferredName
            $bhrDisplayName = "$bhrPreferredName $bhrLastName"
        }

        
        Write-Log -Message "BambooHR employee: $bhrFirstName $bhrLastName ($bhrDisplayName) $bhrWorkEmail" -Severity Debug
        Write-Log -Message "Department: $bhrDepartment, Title: $bhrJobTitle, Manager: $bhrSupervisorEmail HireDate: $bhrHireDate" -Severity Debug
        Write-Log -Message "EmployeeId: $bhrEmployeeNumber, Status: $bhrStatus, Employee Status: $bhrEmploymentStatus" -Severity Debug
        Write-Log -Message "Location: $bhrOfficeLocation, PreferredName: $bhrPreferredName, BestEmail: $bhrBestEmail HomeEmail: $bhrHomeEmail, WorkPhone: $bhrWorkPhone" -Severity Debug
        
        $aadUpnObjDetails = $null
        $aadEidObjDetails = $null
        $isAADUser = $true
        
        <#
            If the user start date is in the past, or in less than -DaysAhead days from current time, we can begin processing the user: 
            create AAD account or correct the attributes in AAD for the employee, else, the employee found on BambooHR will not be processed
        #>

        if (([datetime]$bhrHireDate) -le (Get-Date).AddDays($DaysAhead)) {
            
            $error.clear()
    
            # Check if the user exists in AAD and if there is an account with the EmployeeID of the user checked in the current loop
            Write-Log -Message "Validating $bhrWorkEmail AAD account." -Severity Debug
            $aadUpnObjDetails = Get-MgUser -UserId $bhrWorkEmail -Property id, userprincipalname, Department, EmployeeId, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, EmployeeHireDate, OfficeLocation, BusinessPhones, OnPremisesExtensionAttributes, AdditionalProperties -ExpandProperty manager -ErrorAction SilentlyContinue
            $aadEidObjDetails = Get-MgUser -Filter "employeeID eq '$bhrEmployeeNumber'" -Property employeeid, userprincipalname, Department, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, OfficeLocation, BusinessPhones , EmployeeHireDate, OnPremisesExtensionAttributes, AdditionalProperties -ExpandProperty manager -ErrorAction SilentlyContinue
            $error.clear()
            
            if ([string]::IsNullOrEmpty($aadUpnObjDetails) -eq $false) {
                $UpnExtensionAttribute1 = ($aadUpnObjDetails | Select-Object @{Name = 'bhrLastUpdated'; Expression = { $_.AdditionalProperties.bhrLastUpdated } }).bhrLastUpdated
            }

            if ([string]::IsNullOrEmpty($aadEidObjDetails) -eq $false) {
                #$EID_ExtensionAttribute1 = ($aadEidObjDetails | Select-Object @{Name = 'bhrLastUpdated'; Expression = { $_.AdditionalProperties.bhrLastUpdated } }).bhrLastUpdated
            }

            # Saving AAD attributes to be compared one by one with the details pulled from BambooHR
            $aadWorkEmail = "$($aadUpnObjDetails.Mail)"
            $aadJobTitle = "$($aadUpnObjDetails.JobTitle)"
            $aadDepartment = "$($aadUpnObjDetails.Department)"
            $aadStatus = "$($aadUpnObjDetails.AccountEnabled)"
            $aadEmployeeNumber = "$($aadUpnObjDetails.EmployeeId)"
            $aadEmployeeNumber2 = "$($aadEidObjDetails.EmployeeId)"
            $aadSupervisorEmail = "$(($aadUpnObjDetails | Select-Object @{Name = 'Manager'; Expression = { $_.Manager.AdditionalProperties.mail } }).manager)"
            $aadDisplayname = "$($aadUpnObjDetails.displayName)"
            $aadFirstName = "$($aadUpnObjDetails.GivenName)"
            $aadLastName = "$($aadUpnObjDetails.Surname)"
            $aadCompanyName = "$($aadUpnObjDetails.CompanyName)"
            $aadWorkPhone = "$($aadUpnObjDetails.BusinessPhones)"
            $aadOfficeLocation = "$($aadUpnObjDetails.OfficeLocation)"
            # $aadEmployeeType = ()
            $noAADAccount = $false

            if ($aadUpnObjDetails.EmployeeHireDate) {
                $aadHireDate = $aadUpnObjDetails.EmployeeHireDate.AddHours(12).ToString("yyyy-MM-dd") 
            }
            
            
            if (([string]::IsNullOrWhiteSpace($aadUpnObjDetails) -eq $false) -or ([string]::IsNullOrWhiteSpace($aadEidObjDetails) -eq $false)) {
                Write-Log -Message "AAD user: $aadFirstName $aadLastName ($aadDisplayName) $aadWorkEmail" -Severity Debug 
                Write-Log -Message "Department: $aadDepartment, Title: $aadJobTitle, Manager: $aadSupervisorEmail, HireDate: $aadHireDate" -Severity Debug
                Write-Log -Message "EmployeeId: $aadEmployeeNumber, Enabled: $aadStatus OfficeLocation: $aadOfficeLocation, WorkPhone: $aadWorkPhone" -Severity Debug

                <# If empID of object returned by UPN or by empID is equal, and if object ID is the same as object ID of object returned by UPN and EmpID and
            if UPN = workemail from bamboo  AND if the last changed date from BambooHR is NOT equal to the last changed date saved in ExtensionAttribute1 in AAD, 
            check each attribute and set them correctly, according to BambooHR
            #>

                if (($aadEmployeeNumber -eq $aadEmployeeNumber2 -and `
                            $aadEidObjDetails.UserPrincipalName -eq $aadUpnObjDetails.UserPrincipalName -eq $bhrWorkEmail -and `
                            $aadUpnObjDetails.id -eq $aadEidObjDetails.id -and `
                            $bhrLastChanged -ne $UpnExtensionAttribute1 -and `
                            $aadEidObjDetails.Capacity -ne 0 -and `
                            $aadUpnObjDetails.Capacity -ne 0 -and `
                            $bhrEmploymentStatus -notlike "*suspended*") -and $isAadUser -eq $true) { 
                
                    Write-Log  -Message "$bhrWorkEmail is a valid AAD Account, with matching EmployeeId and UPN in AAD and BambooHR, but different last modified date." -Severity Debug

                    $error.clear() 
                
                    # Checking JobTitle if correctly set, if not, configure the JobTitle as set in BambooHR
                    if ($aadJobTitle.Trim() -ne $bhrJobTitle.Trim()) {
                        Write-Log -Message "AAD Job Title $aadJobTitle does not match BHR Job Title $bhrJobTitle" -Severity Debug
                    
                        if ($TestOnly.IsPresent) {
                        
                            Write-Log -Message " Executing: Update-MgUser -UserId $bhrWorkEmail -JobTitle $bhrJobTitle" -Severity Test   
                        }
                        else {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -JobTitle '$bhrJobTitle'" -Severity Debug
                            if ([string]::IsNullOrWhiteSpace($bhrJobTitle) -eq $false) {
                                Update-MgUser -UserId $bhrWorkEmail -JobTitle $bhrJobTitle
                            }
                            else {
                                Update-MgUser -UserId $bhrWorkEmail -JobTitle $null 
                            }

                            if (!$?) {

                                Write-Log -Message "Error changing Job Title of $bhrWorkEmail.`nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                $error.Clear()
                                Write-Log -Message "JobTitle for $bhrWorkEmail in AAD set from  $aadjobTitle to $bhrjobTitle." -Severity Information
                            }
                        }
                    }

                    # Checking department if correctly set, if not, configure the Department as set in BambooHR
                    if ($aadDepartment.Trim() -ne $bhrDepartment.Trim()) {
                        Write-Log -Message "AAD department '$aadDepartment' does not match BambooHR department '$($bhrDepartment.Trim())'" -Severity Debug
                        if ($TestOnly.IsPresent) {                   
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -Department $bhrDepartment" -Severity Test
                        }
                        else {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -Department $bhrDepartment" -Severity Debug
                            Update-MgUser -UserId $bhrWorkEmail -Department "$bhrDepartment"
                            if (!$?) {

                                Write-Log -Message "Error changing Department of $bhrWorkEmail `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                $error.Clear()
                                Write-Log -Message "Department for $bhrWorkEmail in AAD set from $aadDepartment to $bhrDepartment." -Severity Information
                            }
                        }
                    }
                    else {
                        Write-Log "AAD and BHR department already matches $aadDepartment" -Severity Debug
                    }
        
                    # Checking the manager if correctly set, if not, configure the manager as set in BambooHR
                    if ($aadSupervisorEmail -ne $bhrSupervisorEmail -and ([string]::IsNullOrWhiteSpace($bhrSupervisorEmail) -eq $false)) {
                        Write-Log -Message "Manager in AAD '$aadSupervisorEmail' does not match BHR manager '$bhrSupervisorEmail'" -Severity Debug
                    
                        $aadManagerID = (Get-MgUser -UserId $bhrSupervisorEmail | Select-Object id ).id 
                        $newManager = @{
                            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$aadmanagerID"
                        }

                        if ($TestOnly.IsPresent) {                
                            Write-Log -Message "Executing: Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $newManager"
                        }
                        else {

                            Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $newManager
                            if (!$?) {

                                Write-Log -Message "Error changing manager of $bhrWorkEmail. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                $error.Clear()
                                Write-Log -Message "Manager of $bhrWorkEmail in AAD '$aadsupervisorEmail' and in BambooHR '$bhrsupervisorEmail'. Setting the Manager found on BambooHR to the Azure User Object." -Severity Warning
                            }
                        }
                    }
                    else {
                        Write-Log -Message "Supervisor email already correct $aadSuperVisorEmail" -Severity Debug
                    }

                    # Check and set the Office Location
                    if ($aadOfficeLocation.Trim() -ne $bhrOfficeLocation.Trim()) {
                        Write-Log -Message "AAD office location '$aadOfficeLocation' does not match BHR hire data '$bhrOfficeLocation'" -Severity Debug
                        if ($TestOnly.IsPresent) {
                       
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -OfficeLocation $($bhrOfficeLocation.Trim())" -Severity Test
                        }
                        else {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -OfficeLocation $($bhrOfficeLocation.Trim())" -Severity Debug
                            Update-MgUser -UserId $bhrWorkEmail -OfficeLocation $bhrOfficeLocation.Trim()
                            if (!$?) {

                                Write-Log -Message "Error changing employee office location. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                $error.Clear()
                                Write-Log -Message "Office location of $bhrWorkEmail matches in AAD and BHR $bhrOfficeLocation." -Severity Information
                            }
                        }
                    }
                    else {
                        Write-Log -Message "Office Location correct $aadOfficeLocation" -Severity Debug
                    }

                    # Check and set the Employee Hire Date
                    if ($aadHireDate -ne $bhrHireDate) {
                        Write-Log -Message "AAD hire date '$aadHireDate' does not match BHR hire data '$bhrHireDate'" -Severity Debug
                        if ($TestOnly.IsPresent) {
                                       
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeHireDate $bhrHireDate" -Severity Test
                        }
                        else {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeHireDate $bhrHireDate" -Severity Debug
                            Update-MgUser -UserId $bhrWorkEmail -EmployeeHireDate $bhrHireDate
                            if (!$?) {
                
                                Write-Log -Message "Error changing employee hire date. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                $error.Clear()
                                Write-Log -Message "Hire date of $bhrWorkEmail changed from $aadHireDate in AAD and BHR $bhrHireDate." -Severity Information
                            }
                        }
                    }
                    else {
                        Write-Log -Message "Hire date already correct $aadHireDate" -Severity Debug
                    }

                    # Check and set the workphone
                    if ($aadWorkPhone -ne $bhrWorkPhone) {
                        Write-Log -Message "AAD work phone '$aadWorkPhone' does not match BHR '$bhrWorkPhone'" -Severity Debug
                        if ([string]::IsNullOrWhiteSpace($bhrWorkPhone)) {
                            $bhrWorkPhone = $null
                        }

                        if ($TestOnly.IsPresent) {
                                       
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone" -Severity Test
                        }
                        else {
                            if ([string]::IsNullOrWhiteSpace($bhrWorkPhone)) {
                                $bhrWorkPhone = $null
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $null" -Severity Debug
                                Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $null -ErrorAction Continue
                            }
                            else {
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone" -Severity Debug
                                Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone -ErrorAction Continue
                            }

                            if (!$?) {
                
                                Write-Log -Message "Error changing employee work phone. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                $error.Clear()
                                Write-Log -Message "Work Phone for '$bhrWorkEmail' changed from '$aadWorkPhone' to '$bhrWorkPhone'" -Severity Information
                            }
                        }
                    }
                    else {
                        Write-Log -Message "Work phone correct $aadWorkPhone" -Severity Debug
                    }

                    # Check if user is active in BambooHR, and set the status of the account as it is in BambooHR (active or inactive)
         
                    if ($bhrAccountEnabled -eq $false -and $bhrEmploymentStatus.Trim() -eq "Terminated" -and $aadStatus -eq $true ) {
                        Write-Log -Message "Account is marked Inactive in BHR and Active in AAD, block sign-in, revoke sessions, change pass, remove auth methods"
                        # The account is marked "Inactive" in BHR and "Active" in AAD, block sign-in, revoke sessions, change pass, remove auth methods
                        $error.clear()
                        if ($TestOnly.IsPresent) {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Test
                            Write-Log -Message "Executing: Revoke-MgUserSign -UserId $bhrWorkEmail" -Severity Test
                        }
                        else {
                            Write-Log -Message "Executing: Revoke-MgUserSign -UserId $bhrWorkEmail" -Severity Debug
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Debug
                            Revoke-MgUserSign -UserId $bhrWorkEmail
                            Start-Sleep 10
                            Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled
                            
                        }

                        #Change to a random pass
                        $params = @{
                            PasswordProfile = @{
                                ForceChangePasswordNextSignIn = $true
                                Password                      = (Get-NewPassword)
                            }
                        }

                        if ($TestOnly.IsPresent) {
                            Write-Log -Message " Executing: Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params"  -Severity Test
                            Write-Log -Message " Executing: Get-MgUserMemberOf -UserId $bhrWorkEmail"  -Severity Test

                            Write-Log -Message " Convert Mailbox to shared..." -Severity Test
                            Write-Log -Message " Executing: Set-Mailbox -Identity $bhrWorkEmail -Type Shared" -Severity Test
                            Write-Log -Message " Removing licenses..." -Severity Test
                            Write-Log -Message " Executing: Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }" -Severity Test
                        
                            Write-Log -Message " Remove group memberships" -Severity Test
                            Write-Log -Message " Executing: Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $aadUpnObjDetails.id } " -Severity Test

                            Write-Log -Message " Remove MFA auth for user" -Severity Test

                        }
                        else {
                            Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params
                            Get-MgUserMemberOf -UserId $bhrWorkEmail

                            # Convert mailbox to shared
                            Connect-ExchangeOnline -CertificateThumbprint $AADCertificateThumbprint -AppId $AzureClientAppId -Organization $TenantId 
                            Set-Mailbox -Identity $bhrWorkEmail -Type Shared
                            Disconnect-ExchangeOnline -Confirm:$False
                            # Remove Licenses

                            Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }
                            # Remove groups
                            Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $aadUpnObjDetails.id } 
                            $methodID = Get-MgUserAuthenticationMethod -UserId $bhrWorkEmail | Select-Object id 
                            $methodsdata = Get-MgUserAuthenticationMethod -UserId $bhrWorkEmail | Select-Object -ExpandProperty AdditionalProperties
                            $methods_count = ($methodID | Measure-Object | Select-Object count).count
                            
                            # Pass through each authentication method and remove them
                            $error.Clear() 

                            for ($i = 0 ; $i -lt $methods_count ; $i++) {
       
                                if ((($methodsdata[$i]).Values) -like "*phoneAuthenticationMethod*") { Remove-MgUserAuthenticationPhoneMethod -UserId $bhrWorkEmail -PhoneAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed phone auth method for $bhrWorkEmail." -Severity Warning }
                                if ((($methodsdata[$i]).Values) -like "*microsoftAuthenticatorAuthenticationMethod*") { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $bhrWorkEmail -MicrosoftAuthenticatorAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed auth app method for $bhrWorkEmail." -Severity Warning }
                                if ((($methodsdata[$i]).Values) -like "*windowsHelloForBusinessAuthenticationMethod*") { Remove-MgUserAuthenticationFido2Method -UserId $bhrWorkEmail -Fido2AuthenticationMethodId ($methodID[$i]).id ; Write-Log -Message "Removed PIN auth method for $bhrWorkEmail." -Severity Warning }
                            }

                            # Update-MgUser -EmployeeId "LVR" -UserId $bhrWorkEmail.Trim()                                            
                            if ($error.Count -ne 0) {
                                $error | ForEach-Object {
                                    $err_Exception = $_.Exception
                                    $err_Target = $_.TargetObject
                                    $errCategory = $_.CategoryInfo
                                    Write-Log "Could not remove authentication details. `
 Exception: $err_Exception `n Target Object: $err_Target `n Error Category: $errCategory " -Severity Error
                                }
                            }
                            else {
                                Write-Log -Message "Account $bhrWorkEmail marked as inactive in BambooHR AAD account has been disabled, sessions revoked and removed MFA." -Severity Information              
                                $error.Clear()
                            }
                        }
                    }
                    else {
                        Write-Log "Account is not disabled or terminated." -Severity Debug
                    }
          
                    if ($bhrAccountEnabled -eq $true -and $aadstatus -eq $false) {
                        # The account is marked "Active" in BHR and "Inactive" in AAD, enable the AAD account
                        Write-Log -Message "$bhrWorkEmail is marked Active in BHR and Inactive in AAD" -Severity Debug
                        if ($TestOnly.IsPresent) {
                            Write-Log -Message " Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Test
                        }
                        else {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Debug
                            Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled
                            if (!$?) {
    
                                Write-Log -Message "Could not activate the user account. ``nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                Write-Log -Message "Account $bhrWorkEmail marked as Active in BambooHR but Inactive in AAD. Enabled AAD account for sign-in." -Severity Information
                                $error.Clear()
                            }                   
                        }
                    }
                    else {
                        Write-Log -Message "Account is not enabled in BHR and disabled in AAD" -Severity Debug
                    }
               
                    # Compare user employee id with BambooHR and set it if not correct
                    if ($bhrEmployeeNumber.Trim() -ne $aadEmployeeNumber.Trim()) {
                        Write-Log -Message "BHR employee number $bhrEmployeeNumber does not match AAD employee id $aadEmployeeNumber" -Severity Debug
                        if ($TestOnly.IsPresent) {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber  "
                        }
                        else {
                            # Setting the Employee ID found in BHR to the user in AAD
                            Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber.Trim()             
                            if (!$?) {
                                $err_msg = $Error.exception
                                $err_TargetObject = $error.TargetObject
                                $err_details = $error.ErrorDetails
                                $err_trace = $error.ScriptStackTrace

                                Write-Log -Message "Error changing EmployeeId. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                Write-Log -Message "The ID $bhremployeeNumber has been set to $bhrWorkEmail AAD account. IF condition result is True." -Severity Warning
                                $error.Clear()
                            }
                        }
                    }
                    else {
                        Write-Log -Message "Employee ID matched $bhrEmployeeNumber and $aadEmployeeNumber" -Severity Debug
                    }

                    # Set Company name to $CompanyName"
                    if ($aadCompanyName.Trim() -ne $CompanyName.Trim()) {
                        Write-Log -Message "AAD company name '$aadCompany' does not match '$CompanyName'" -Severity Debug
                        if ($TestOnly.IsPresent) {
                    
                            Write-Log -Message " Executing: Update-MgUser -UserId $bhrWorkEmail -CompanyName $($CompanyName.Trim())" -Severity Information -Severity Test
                        }
                        else {
                            # Setting Company Name as $CompanyName to the employee, if not already set
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -CompanyName $($CompanyName.Trim())" -Severity Debug
                            Update-MgUser -UserId $bhrWorkEmail -CompanyName $CompanyName.Trim()
                            if (!$?) {

                                Write-Log -Message "Could not change the Company Name of $bhrWorkEmail. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                Write-Log -Message "The $bhrWorkEmail employee Company attribute has been set to: $CompanyName." -Severity Information
                            }
                        }
                    }
                    else {
                        Write-Log -Message "Company name already matched in AAD and BHR $aadCompanyName" -Severity Debug
                    }

                    # Set LastModified from BambooHR to ExtensionAttribute1 in AAD
<#
                     if ($upnExtensionAttribute1 -ne $bhrLastChanged) {
                        # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AAD
                        Write-Log -Message "AAD Extension Attribute '$upnExtensionAttribute1' does not match BHR last changed '$bhrLastChanged'" -Severity Debug
                        Write-Log -Message "Set LastModified from BambooHR to ExtensionAttribute1 in AAD" -Severity Debug
                        
                        if ($TestOnly.IsPresent) {
                    
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AdditionalProperties @{'bhrLastChanged' = '$bhrLastChanged' }" -Severity Test
                        }
                        else {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AdditionalProperties @{'bhrLastChanged' = '$bhrLastChanged' }" -Severity Debug
                            # TODO DOES NOT WORK
                            #Update-MgUser -UserId $bhrWorkEmail -AdditionalProperties @{'extension_bhrLastChanged' = "$bhrLastChanged" } 
                            if (!$?) {
                                $err_msg = $Error.exception
                                $err_TargetObject = $error.TargetObject
                                $err_details = $error.ErrorDetails
                                $err_trace = $error.ScriptStackTrace

                                Write-Log -Message "Error changing ExtensionAttribute1. `n Exception: $err_msg `n Target object: $err_TargetObject `n Details: $err_details `n StackTrace: $err_trace" -Severity Error
                                $error.Clear()
                            }
                            else {
                                Write-Log -Message "$bhrWorkEmail LastChanged attribute set from '$upnExtensionAddtribute' to '$bhrlastChanged'." -Severity Information
                            }
                        }

                        $error.clear()             
                    }
                    else {
                        Write-Log -Message "Attribute already matched last changed of $bhrLastChanged" -Severity Debug
                    } 
#>
                }
            }
            else {
                Write-Log -Message "No AAD user found for $bhrWorkEmail" -Severity Debug
                $noAADAccount = $true
                
                # This might not be needed anymore
                $aadWorkEmail = ""
                $aadJobTitle = ""
                $aadDepartment = ""
                $aadStatus = ""
                $aadEmployeeNumber = ""
                $aadSupervisorEmail = ""
                $aadDisplayname = ""
                $aadHireDate = ""
                $aadFirstName = ""
                $aadLastName = ""
                $aadCompanyName = ""
                $aadWorkPhone = ""
                $aadOfficeLocation = ""
                $isAadUser = $false
            }

            # Handle name changes 
            if ($aadEmployeeNumber -eq $bhremployeeNumber -and $historystatus -notlike "*inactive*") {
                $aadUPN = $aadEidObjDetails.UserPrincipalName
                $aadObjectID = $aadEidObjDetails.id
                $aadworkemail = $aadEidObjDetails.Mail
                $aademployeeNumber = $aadEidObjDetails.EmployeeID
                $aaddisplayname = $aadEidObjDetails.displayname
                $aadfirstName = $aadEidObjDetails.GivenName
                $aadlastName = $aadEidObjDetails.Surname

                Write-Log -Message "Initiated AAD name change for $aadfirstName $aadlastName ($aaddisplayname) `n`t Work Email: $aadWorkEmail UserPrincipalName: $aadUpn EmployeeId: $aadEmployeeNumber" -Severity Debug
           
                $error.Clear()
                # 3/31/23 Is this required here or should it be handled after the name change or the next sync after the name change?
                #                 # Set LastModified from BambooHR to ExtensionAttribute1 in AAD
                #                 if ($EID_ExtensionAttribute1 -ne $bhrlastChanged) {

                #                     if ($TestOnly.IsPresent) {
                #                         #Write-Log -Message "The $bhrWorkEmail employee LastChanged attribute set to extensionAttribute1 as $bhrlastChanged." -Severity Information -Severity Test
                #                         Write-Log -Message "Executing: Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } -UserId $aadObjectID" -Severity Test
                #                     }
                #                     else {
                #                         # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AAD
                #                         # Write-Log -Message "The $bhrWorkEmail employee LastChanged attribute set to extensionAttribute1 as $bhrlastChanged." -Severity Information
                #                         Write-Log -Message "Executing: Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } -UserId $aadObjectID" -Severity Debug

                #                         # TODO
                #                         Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } -UserId $aadObjectID
                        
                #                         if (!$?) {
                #                             $err_msg = $Error.exception
                #                             $err_TargetObject = $error.TargetObject
                #                             $err_details = $error.ErrorDetails
                #                             $err_trace = $error.ScriptStackTrace

                #                             Write-Log -Message "Error changing ExtensionAttribute1.`
                # Exception: $err_msg `n Target object: $err_TargetObject`n Details: $err_details `n StackTrace: $err_trace" -Severity Error
                #                             $error.Clear()
                #                         }
                #                         else {
                #                             Write-Log -Message "ExtensionAttribute1 changed to: $bhrlastChanged for employee $bhrWorkEmail." -Severity Information
                #                         }
                #                     }
                #                 }

                # Change last name in Azure         
                if ($aadLastName -ne $bhrLastName) {
                    Write-Log -Message "Last name in AAD $aadLastName does not match in BHR $bhrLastName" -Severity Debug
                    Write-Log -Message "Changing the last name of $bhrWorkEmail from $aadLastName to $bhrLastName." -Severity Debug
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Surname $bhrLastName" -Severity Test
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Surname $bhrLastName" -Severity Debug
                        Update-MgUser -UserId $aadObjectID -Surname $bhrLastName
                        
                        if (!$?) {
                            $err_msg = $Error.exception
                            $err_TargetObject = $error.TargetObject
                            $err_details = $error.ErrorDetails
                            $err_trace = $error.ScriptStackTrace
                               
                            Write-Log -Message "Error changing AAD Last Name.`
Exception: $err_msg `n Target object: $err_TargetObject `n Details: $err_details `nStackTrace: $err_trace" -Severity Error
                            $error.Clear()
                        }
                        else {
                            Write-Log -Message "Successfully changed the last name of $bhrWorkEmail from $aadLastName to $bhrLastName." -Severity Information
                        }
                    }
                }
            
                # Change First Name
                if ($aadfirstName -ne $bhrfirstName) {
                    Write-Log "AAD First Name $aadfirstName is not equal to BHR First Name $bhrFirstName" -Severity Debug
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -GivenName $bhrFirstName" -Severity Test
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -GivenName $bhrFirstName" -Severity Debug
                        Update-MgUser -UserId $aadObjectID -GivenName $bhrFirstName
                        if (!$?) {
                            $err_msg = $Error.exception
                            $err_TargetObject = $error.TargetObject
                            $err_details = $error.ErrorDetails
                            $err_trace = $error.ScriptStackTrace
                             
                            Write-Log -Message "Could not change the First Name of $aadObjectID. Error details below. `
 Exception: $err_msg `n Target object: $err_TargetObject `n Details: $err_details `n StackTrace: $err_trace" -Severity Error
                            $error.Clear()
                        }   
                        else {
                            Write-Log -Message "Successfully changed $aadObjectID first name from $aadFirstName to $bhrFirstName." -Severity Information
                        }       
                    } 
                }
           
                # Change display name
                if ($aadDisplayname -ne $bhrDisplayName) {
                    Write-Log -Message "AAD Display Name $aadDisplayname is not equal to BHR $bhrDisplayName" -Severity Debug
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -DisplayName $bhrdisplayName" -Severity Test
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -DisplayName $bhrdisplayName" -Severity Debug
                        Update-MgUser -UserId $aadObjectID -DisplayName $bhrdisplayName

                        if (!$?) {
                            $err_msg = $Error.exception
                            $err_TargetObject = $error.TargetObject
                            $err_details = $error.ErrorDetails
                            $err_trace = $error.ScriptStackTrace
                             
                            Write-Log -Message "Could not change the Display Name. Error details below. `
 Exception: $err_msg `n Target object: $err_TargetObject `n Details: $err_details `n StackTrace: $err_trace" -Severity Error
                            $error.Clear()
                        }# Change display name - Error logging
                        else {
                            Write-Log "Display name $aadDisplayName of $aadObjectID changed to $bhrDisplayName." -Severity Information
                        }        
                    }
                }

                # Change Display Name script block closure

                # Change Email Address
                if ($aadWorkEmail -ne $bhrWorkEmail) {
                    Write-Log -Message "aadWorkEmaill $aadWorkEmail is not equal to bhrWorkEmail $bhrWorkEmail"
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Mail $bhrWorkEmail"
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Mail $bhrWorkEmail"
                        Update-MgUser -UserId $aadObjectID -Mail $bhrWorkEmail
                        
                        if (!$?) {
                            $err_msg = $Error.exception
                            $err_TargetObject = $error.TargetObject
                            $err_details = $error.ErrorDetails
                            $err_trace = $error.ScriptStackTrace
                             
                            Write-Log -Message "Error changing Email Address. `
                        Exception: $err_msg `
                        Target object: $err_TargetObject `
                        Details: $err_details `
                        StackTrace: $err_trace" -Severity Error
                            $error.Clear()
                        }
                        else {
                            # Change Email Address error logging
                            Write-Log "The current Email Address: $aadworkemail of $aadObjectID has been changed to $bhrWorkEmail. If condition true." -Severity Warning
                        }             
                    }
                }

                # Change UserPrincipalName and send the details via email to the User
                if ($aadUpn -ne $bhrWorkEmail) {
                    Write-Log -Message "aadUPN $aadUpn does not match bhrWorkEmail $bhrWorkEmail" -Severity Debug
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -UserPrincipalName $bhrWorkEmail" -Severity Test
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -UserPrincipalName $bhrWorkEmail" -Severity Debug
                        Update-MgUser -UserId $aadObjectID -UserPrincipalName $bhrWorkEmail
                       
                        if (!$?) {
                            $err_msg = $Error.exception
                            $err_TargetObject = $error.TargetObject
                            $err_details = $error.ErrorDetails
                            $err_trace = $error.ScriptStackTrace
                             
                            Write-Log -Message "Error changing UPN for $aadObjectID. `
Exception: $err_msg `n Target object: $err_TargetObject `n Details: $err_details `n StackTrace: $err_trace" -Severity Error
                            $error.Clear()
                        } 
                        else {
                            Write-Log -Message "Changed the current UPN:$aadUPN of $aadObjectID to $bhrWorkEmail. IF Condition true." -Severity Warning
                            $params = @{
                                Message         = @{
                                    Subject      = "Login details change for $bhrdisplayName"
                                    Body         = @{
                                        ContentType = "text"
                                        Content     = "Hello, `
`
Your name has been changed, therefore your user account has also changed. `
Use your new user name: $bhrWorkEmail `n Your password has not been modified.`n $EmailSignature"
                                    }
                                    ToRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $bhrWorkEmail
                                            }
                                        }
                                    )
                                    CCRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $AdminEmailAddress, $NotificationsEmailAddress, $bhrSupervisorEmail
                                            }
                                        }
                                    )
                                }
                                SaveToSentItems = "True"
                            }

                            Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress
                        }
                    }
                }
            

            }   
            # CREATE ACCOUNT FOR NEW EMPLOYEE   
            if (($noAADAccount -eq $true) -and ($bhrAccountEnabled -eq $true)) {
                Write-Log -Message "No AAD account exist but employee in bhr is $bhrAccountEnabled)" -Severity Debug
            
                $PasswordProfile = @{
                    Password = (Get-NewPassword)
                }
                $error.clear() 
            
                if ($TestOnly.IsPresent) {
                    # Write logging here
                    Write-Log -Message "Executing New-MgUser `
                                -Department $bhrdepartment `
                                -EmployeeId $bhremployeeNumber `
                                -JobTitle $bhrjobTitle `
                                -CompanyName $CompanyName `
                                -Surname $bhrlastName `
                                -GivenName $bhrfirstName `
                                -DisplayName $bhrdisplayName `
                                -AccountEnabled `
                                -Mail $bhrWorkEmail `
                                -EmployeeHireDate $bhrHireDate `
                                -UserPrincipalName $bhrWorkEmail `
                                -PasswordProfile $PasswordProfile `
                                -MailNickname ($bhrWorkEmail -replace '@', '' -replace $companyEmailDomain, '') `
                                -UsageLocation $UsageLocation `
                                -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged " -Severity Debug
                                
                }
                else {
                    # Create AAD account, as it doesn't have one, if user hire date is less than 14 days in the future, or is in the past
                    Write-Log -Message "$bhrWorkEmail does not have an AAD account and hire date is less than 14 days from present time or in the past." -Severity Information
                    
                    Write-Log -Message "Executing New-MgUser -EmployeeId $bhremployeeNumber `
                                -CompanyName $CompanyName -Surname $bhrlastName `
                                -GivenName $bhrfirstName -DisplayName $bhrdisplayName `
                                -AccountEnabled -Mail $bhrWorkEmail `
                                -EmployeeHireDate $bhrHireDate -UserPrincipalName $bhrWorkEmail `
                                -PasswordProfile $PasswordProfile `
                                -MailNickname ($bhrWorkEmail -replace '@', '' -replace $companyEmailDomain, '' ) `
                                -UsageLocation $UsageLocation `
                                -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged " -Severity Debug
            
                    New-MgUser -EmployeeId $bhrEmployeeNumber `
                        -CompanyName $CompanyName -Surname $bhrlastName `
                        -GivenName $bhrfirstName -DisplayName $bhrdisplayName `
                        -AccountEnabled -Mail $bhrWorkEmail `
                        -EmployeeHireDate $bhrHireDate -UserPrincipalName $bhrWorkEmail `
                        -PasswordProfile $PasswordProfile `
                        -MailNickname ($bhrWorkEmail -replace '@', '' -replace $companyEmailDomain, '') `
                        -UsageLocation $UsageLocation `
                        -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged 
                    }
                            
                    if ($? -eq $true) {
                        
                        if ([string]::IsNullOrWhiteSpace($bhrSupervisorEmail) -eq $false) {
                            Write-Log -Message "Account $bhrWorkEmail successfully created." -Severity Information
                            $aadmanagerID = (Get-MgUser -UserId $bhrSupervisorEmail | Select-Object id).id
                        
                            $newManager = @{
                                "@odata.id" = "https://graph.microsoft.com/v1.0/users/$aadmanagerID"
                            }
                            Start-Sleep -Seconds 8
                        
                            Write-Log -Message "Setting the manager of the newly created user $bhrWorkEmail." -Severity Debug
                            Write-Log -Message "Executing: Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $NewManager" -Severity Debug
                            Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $NewManager
                            $params = @{
                                Message         = @{
                                    Subject      = "User creation automation: $bhrdisplayName"
                                    Body         = @{
                                        ContentType = "text"
                                        Content     = "Hello, `n`t New employee user account created for $bhrDisplayName.`
    Please review the new user setup guide for additional information.`n`t User name: $bhrWorkEmail `n`t Password: $($PasswordProfile.Values) `n `n $EmailSignature"
                                    }
                                    ToRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $bhrSupervisorEmail
                                            }
                                        }
                                    )
                                    CcRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $AdminEmailAddress
                                            }
                                        }
                                    )
                                }
                                SaveToSentItems = "True"
                            }
                        }
                        else {
                            $params = @{
                                Message         = @{
                                    Subject      = "User creation automation: $bhrdisplayName"
                                    Body         = @{
                                        ContentType = "text"
                                        Content     = "Hello, `n`t New employee user account created for $bhrDisplayName.`
    No manager account is currently active for this account so this info is being sent to the default location.`n
    Please reference the new user setup guide for additional information.`n`t User name: $bhrWorkEmail `n`t Password: $($PasswordProfile.Values) `n $EmailSignature"
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
                        }

                        #Assigning the user to BambooHR enterprise app
                        #$uid = (get-mguser -UserId $bhrWorkEmail | Select-Object ID).id
                        #New-MgUserAppRoleAssignment -UserId $uid -PrincipalId $uid `
                        #    -ResourceId 6b419818-6e25-4f9c-9268-1f0d2ef78700 `
                        #    -AppRoleId a8972ac9-2341-4879-9028-2a9d979844a0
            
                        # Send mail with credentials of the newly created user
                        
            
                        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress 
            
                    }
                    else {
                        # If "User account creation succeded" closure
                        $creationerror = $error
                        $full_Error_Details = $creationerror | Select-Object *
                        $creationerror_exception = $creationerror.Exception.Message
                        $creationerrorCategory = $creationerror.CategoryInfo
                        $creationerror_errID = $creationerror.FullyQualifiedErrorId
                        $creationerrorStack = $creationerror.ScriptStackTrace    
            
                        Write-Log -Message "Account $bhrWorkEmail creation failed. New-Mguser cmdlet returned error. `n $full_Error_Details"-Severity Error
                        $params = @{
                            Message         = @{
                                Subject      = "FAILURE: User creation automation $bhrdisplayName"
                                Body         = @{
                                    ContentType = "text"
                                    Content     = "Hello, `
            `
            Account creation for user: $bhrWorkEmail has failed. Please check the log: $logFileName for further details. Error message summary below. `
            Error Message: $creationerror_exception `n Error Category: $creationerrorCategory `n Error ID: $creationerror_errID `n Stack: $creationerrorStack `n $EmailSignature"
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
                        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress
            
                    }# else "User account creation failed" closure
                }
            }                    
        }
        else {
            # If Hire Date is less than 14 days in the future or in the past closure
            # The user account does not need to be created as it does not satisfy the condition of the HireDate being 14 days or less in the future
            Write-Log -Message "The Employee $bhrWorkEmail hire date is more than 14 days in the future. Will be created when HireDate is 14 days or less in the future." -Severity Information
        }
    }



}


$params = @{
    Message         = @{
        Subject      = "$($logFileName.Split(".")[0])"
        Body         = @{
            ContentType = "text"
            Content     = "The following was completed by the sync. `n $Script:LogContent `n $EmailSignature"
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

Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress

# RUNTIME BLOCK CLOSURE 
$runtime_seconds = $runtime.TotalSeconds
Write-Log -Message "Total time the script ran: $runtime_seconds" -Severity Information
Disconnect-MgGraph | Out-Null
Exit
#Script End
