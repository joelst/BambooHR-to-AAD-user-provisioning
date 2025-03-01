#Requires -Module ExchangeOnlineManagement,PSTeams,Microsoft.Graph.Users,Microsoft.Graph.Authentication,Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Identity.SignIns
<#

IMPORTANT: This is a sample solution and should be used by those comfortable testing and retesting and validating before using it in production. 
All content is provided AS IS with no guarantees or assumptions of quality or functionality. 

If you are using employee information there is much that can go wrong! 

YOU are responsible for complying with applicable laws and regulations for handling PII. 
Remember, with great power comes great responsibility. 
Friends don't let friends run untested scripts in production.
Don't take any wooden nickels

.SYNOPSIS
Script to synchronize employee information from BambooHR to Azure Active Directory (Entra Id). It does not support on premises Active Directory.

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

.PARAMETER HelpDeskEmailAddress
Email address for help desk

.PARAMETER EmailSignature 
Signature to add to the bottom of all sent email messages

.PARAMETER WelcomeUserText
Sentence to add to new user email messages specific to finding the IT helpdesk FAQ.

.PARAMETER LogPath
Location to save logs

.PARAMETER UsageLocation
A two letter country code (ISO standard 3166) to set AAD usage location. 
Required for users that will be assigned licenses due to legal requirement to check for availability of services in countries. Examples include: US, JP, and GB.

.PARAMETER DaysAhead
Number of days to look ahead for the employee to start.

.PARAMETER TestOnly
Specify when you do not want to make any changes.

.PARAMETER EnableMobilePhoneSync
Use this to synchronize mobile phone numbers from BHR to AAD.

.PARAMETER CurrentOnly
Specify to only pull current employees from BambooHR. Default is to retrieve future employees.

.PARAMETER NotificationEmailAddress
Specifies an additional email address to send any notification emails to. 

.PARAMETER ForceSharedMailboxPermissions
When specified shared mailbox permissions are updated

.PARAMETER LicenseId
When specified with a valid license id it will make sure there are still unassigned licenses before creating a new user.

.NOTES
More documentation available in project README

#>

[CmdletBinding()]
param (
    [Parameter()]
    [String]
    $BambooHrApiKey = $env:BambooHrApiKey,
    [Parameter()]
    [String]
    $AdminEmailAddress = "bhr-sync@companydomain.com",
    [Parameter()]
    [string]
    $BHRCompanyName = $env:BHRCompanyName,
    [Parameter()]
    [string]
    $CompanyName = $env:CompanyName,
    [Parameter()]
    [string]
    $TenantId = $env:TenantId,
    [parameter()]
    [string]
    $LogPath = "C:\temp\bamboo",
    [parameter()]
    [string]
    $UsageLocation = "US",
    # Number of days before the user starts to provision their account
    [parameter()]
    [int]
    $DaysAhead = 7,
    [parameter()]
    [string]
    $NotificationEmailAddress = "hr@companydomain.com",
    [parameter()]
    [string]
    $HelpDeskEmailAddress = "helpdesk@companydomain.com",
    [parameter()]
    [string]
    $EmailSignature = "<br/>Regards, <br/> $CompanyName Automated User Management <br/><br/>For additional information, please review the <a href='https://bing.com'>IT FAQ.</a><br/>",
    [parameter()]
    [string]
    $WelcomeUserText = "Please review the <a href='https://bing.com'>new user setup guide</a> for getting started. We also recommend printing the <a href='https://bing.com'>IT Quick Reference Guide</a> so can find help when technology isn't cooperating.",
    [parameter()]
    [string]
    $DefaultProfilePicPath = (Join-Path Get-Location "DefaultProfilePic.jpg"),
    [parameter()]
    [string]
    $TeamsCardUri = $TeamsCardUri,
    [parameter()]
    [switch]
    $TestOnly,
    [parameter()]
    [switch]
    $EnableMobilePhoneSync,
    [parameter()]
    [switch]
    $CurrentOnly,
    [parameter()]
    [switch]
    $ForceSharedMailboxPermissions,
    $MailboxDelegationParams = @(
        @{
            Group           = "CG-SharedMailboxDelegatedAccessMailbox1"
            DelegateMailbox = "Mailbox1"
        }
        @{
            Group           = "CG-SharedMailboxDelegatedAccessMailbox2"
            DelegateMailbox = "Mailbox2"
        }
        @{
            Group           = "CG-SharedMailboxDelegatedAccessMailbox3"
            DelegateMailbox = "Mailbox3"
        }                            
    ),
    [parameter()]
    [string]
    # Business Premium License Id - replace with the license id that you are assigning
    $LicenseId = "21546e07-2132-445d-b21e-9ca36bf847c2_cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"
)

$AzureAutomate = $true
# Logging Function
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
    
    if ($AzureAutomate -eq $false) {
        [pscustomobject]@{
            Time     = (Get-Date -Format "yyyy/MM/dd HH:mm:ss")
            Message  = ($Message.Replace("`n", '').Replace("`t", '').Replace("``", ''))
            Severity = $Severity
        } | Export-Csv -Path $logFilePath -Append -NoTypeInformation -Force
    }

    switch ($Severity) {
        Debug { 
            Write-Verbose $Message
        }
        Warning { 
            Write-Warning $Message
            $Script:logContent += "<p>$Message</p>`n"
        }
        Error { 
            Write-Error $Message
            $Script:logContent += "<p>$Message</p>`n"
        }
        Information { 
            Write-Output $Message
            $Script:logContent += "<p>$Message</p>`n"
        }
        Test { 
            Write-Host " [TestOnly] $Message" -ForegroundColor Green
            
        } 
        Default { 
            Write-Host $Message
        }
    }
} 

function Get-LicenseStatus {
    <#
    .SYNOPSIS
    Get the license status for the specified license id.
    .PARAMETER LicenseId
    The license id to check for availability.
    .PARAMETER TeamsCardUri
    The URI to send the adaptive card to.
    .NOTES
    This function is used to check if there are any available licenses for the specified license id.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $LicenseId,
        [Parameter()]
        [string]
        $TeamsCardUri,
        [Parameter()]
        [int]
        $MaximumExtraLicenses = 4,
        [Parameter()]
        [switch]
        $NewUser
    )

    $licenses = Get-MgSubscribedSku -SubscribedSkuId $LicenseId | Select-Object SkuPartNumber, SkuId, ConsumedUnits, @{Name = 'EnabledUnits'; Expression = { $_.PrepaidUnits.Enabled } }, @{Name = 'SuspendedUnits'; Expression = { $_.PrepaidUnits.Suspended } }, @{Name = 'WarningUnits'; Expression = { $_.PrepaidUnits.Warning } }
    $licensesConsumed = $licenses.ConsumedUnits
    $licensesEnabled = $licenses.EnabledUnits
    $licensesEnabled = $licensesEnabled
    if ($NewUser.IsPresent) {
        $licensesConsumed++
    }

    $licensesAvailable = $licensesEnabled - $licensesConsumed

    if ($licensesAvailable -lt 0 -and $NewUser.IsPresent) {
        Write-Log -Message "`n There are no licenses available for a newly created user!" -Severity Error
        $params = @{
            Message         = @{
                Subject      = "BhrAadSync: There are no licenses available for a newly created user!"
                Body         = @{
                    ContentType = "html"
                    Content     = "No licenses available for the new user. <br/> There are $($licensesConsumed) of $($licensesEnabled) assigned. $EmailSignature"
                }
                ToRecipients = @(
                    @{
				                    EmailAddress = @{
                            Address = $HelpDeskEmailAddress
                        }
                    }
                )
            }
            SaveToSentItems = "True"
        }

        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose
        New-AdaptiveCard {
            New-AdaptiveTextBlock -Text "There are $($licensesConsumed) of $($licensesEnabled) assigned!" -HorizontalAlignment Center -Weight Bolder -Size ExtraLarge
            New-AdaptiveTextBlock -Text "The number of licenses should be increased" -Wrap -Weight Bolder -Size ExtraLarge
        } -Uri $TeamsCardUri -Speak "There are no licenses left for a new user!"
    }
    elseif ($licensesAvailable -le 0) {
        Write-Log -Message "`n There are no additional licenses available!" -Severity Error
        $params = @{
            Message         = @{
                Subject      = "BhrAadSync: No additional licenses available"
                Body         = @{
                    ContentType = "html"
                    Content     = "No additional licenses are available. <br/> There are $($licensesConsumed) of $($licensesEnabled) assigned. $EmailSignature"
                }
                ToRecipients = @(
                    @{
				                    EmailAddress = @{
                            Address = $HelpDeskEmailAddress
                        }
                    }
                )
            }
            SaveToSentItems = "True"
        }

        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose
        New-AdaptiveCard {
            New-AdaptiveTextBlock -Text "There are $($licensesConsumed) of $($licensesEnabled) assigned!" -HorizontalAlignment Center -Weight Bolder -Size ExtraLarge
            New-AdaptiveTextBlock -Text "The number of licenses should be increased" -Wrap -Weight Bolder -Size ExtraLarge
        } -Uri $TeamsCardUri -Speak "There are no licenses left for a new user!"
    }
    elseif ($licenses.ConsumedUnits -lt ($licensesEnabled - $MaximumExtraLicenses)) {
        Write-Log -Message "`n There are too many licenses left!" -Severity Error
        $params = @{
            Message         = @{
                Subject      = "BhrAadSync: Too many extra licenses"
                Body         = @{
                    ContentType = "html"
                    Content     = "Too many extra licenses. <br/> There are $($licensesConsumed) of $($licensesEnabled) assigned. $EmailSignature"
                }
                ToRecipients = @(
                    @{
				                    EmailAddress = @{
                            Address = $HelpDeskEmailAddress
                        }
                    }
                )
            }
            SaveToSentItems = "True"
        }

        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose
        New-AdaptiveCard {
            New-AdaptiveTextBlock -Text "There are $($licensesConsumed) of $($licensesEnabled) assigned!" -HorizontalAlignment Center -Weight Bolder
            New-AdaptiveTextBlock -Text "Consider reducing the number of licenses" -Wrap -Weight Bolder -Size ExtraLarge
        } -Uri $TeamsCardUri -Speak "Too many extra licenses!"
    }
    else {
        Write-Log -Message "`n $($licensesConsumed) of $($licensesEnabled) licenses $LicenseId have been assigned." -Severity Information
    }

    return $licenses
}

function Get-NewPassword {
    <#
        .DESCRIPTION
            Generate a random password with the configured number of characters and special characters.
            Does not return characters that are commonly confused like 0 and O and 1 and l. Also removes characters that cause issues in PowerShell scripts.
        .EXAMPLE
            Get-NewPassword -PasswordLength 13 -SpecialChars 4
            Returns a password that is 13 characters long and includes 4 special characters.
    
        .PARAMETER PasswordLength
            Specifies the total length of password to generate
        .PARAMETER SpecialChars
            Specifies the number of special characters to include in the generated password.
        
        .NOTES
            Inspired by: http://blog.oddbit.com/2012/11/04/powershell-random-passwords/
        #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [int]$PasswordLength = 14,    
        [int]$SpecialChars = 2
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

function Sync-GroupMailboxDelegation {
    <#
    .SYNOPSIS
    You can assign a group as a mailbox delegate to allow all users delegate access to the mailbox. However, when a group is assigned,
    Outlook for Windows users will not get these delegate mailboxes automapped. The user must manually add the mailbox to their Outlook profile.
    If users are accessing mail using Outlook for web or Mac, automapping is not supported, so you can simply assign a group delegated permissions.
 
    .DESCRIPTION
    This script will add and remove delegates to an Exchange Online mailbox. Specify the group name and the mailbox for which to provide access.

    .PARAMETER Group
    The Azure AD Group or Distribution group members to apply permissions

    .PARAMETER DelegateMailbox
    Mailbox to delegate access to

    .PARAMETER LeaveExistingDelegates
    Do not remove any of the existing delegates

    .PARAMETER Permissions
    Provide list of permissions to delegate. Default includes FullAccess and SendAs

    .PARAMETER DoNotConnect
    Specify when the PowerShell session is already properly authenticated with ExchangeOnline. Then it will not be connected again inside the function.
#>

    [CmdletBinding()]
    param (   
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Group,
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $DelegateMailbox,
        [switch]
        $LeaveExistingDelegates,
        [string[]]
        $Permissions = @("FullAccess", "SendAs"),
        [Parameter()]
        [string]
        $TenantId = $TenantId,
        [switch]$DoNotConnect

    )

    function Get-MgGroupMemberRecursively {
        param([Parameter()][string]$GroupId, 
            [Parameter()][string]$GroupDisplayName
        ) 
        if ([string]::IsNullOrWhiteSpace($GroupId)) {
            $GroupId = (Get-MgGroup -Filter "DisplayName eq '$GroupDisplayName'" -ErrorAction SilentlyContinue).Id
        }

        $output = @()
        if ($GroupId) {
            Get-MgGroupMember -GroupId $GroupId -All | ForEach-Object { 
                if ($_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user") {
                    $output += $_
                }
                if ($_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group") {
                    $output += @(Get-MgGroupMemberRecursively -GroupId $_.Id)
                }
    
            }
        }
        return $output
    }

    if ($DoNotConnect.IsPresent -eq $false) {
        Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false | Out-Null
    }

    # Find the shared mailbox
    if ([string]::IsNullOrWhiteSpace($DelegateMailbox) -eq $false) {
        $mObj = Get-ExoMailbox -anr $DelegateMailbox
    }

    if ($null -eq $mObj) {
        Write-Log " Shared mailbox $DelegateMailbox not found!" -Severity Error
        exit 1
    }

    Write-Log "`t$DelegateMailbox matched with $($mObj) $($mObj.Identity) " -Severity Debug
    Connect-MgGraph -Identity -NoWelcome
    $gMembers = Get-MgGroupMemberRecursively -GroupDisplayName $Group | Sort-Object -Property Id -Unique
    Write-Log " $group member count: $($gMembers.Count)" -Severity Debug

    if ($Permissions -contains "FullAccess") {

        $existingFullAccessPermissions = Get-ExoMailboxPermission -Identity $mObj.identity | Sort-Object -Property User -Unique | Where-Object { $_.User -notlike "*SELF" } | Sort-Object -Unique -Property User | Foreach-object { Get-MgUser -UserId $_.User }
        if ($gMembers) {
            $cPermissions = Compare-Object -ReferenceObject $existingFullAccessPermissions -DifferenceObject $gMembers -Property Id -ErrorAction SilentlyContinue
        }
        $missingPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ "=>"
        Write-Log " Missing perms: $($missingPermissions.Count + 0)" -Severity Debug
        $extraPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ "<="
        Write-Log " Extra perms: $($extraPermissions.Count + 0)" -Severity Debug
        
        # if need to add FullAccess
        if (($missingPermissions.Count + 0) -gt 0) {
            Write-Log "Adding $($missingPermissions.Count) missing permission(s) based on group membership" -Severity Information
        
            foreach ($missing in $missingPermissions) {
                $u = Get-MgUser -UserId $missing.id
                Write-Log "`tAdding Full Access permissions for $($u.DisplayName) to $($mObj.Identity) $DelegateMailbox..." -Severity Debug
                Add-MailboxPermission -Identity $mObj.Identity -User $missing.Id -AccessRights ‘FullAccess’ -Automapping:$true –inheritancetype All | Out-Null
            }
        }
        else {
            Write-Log "No Full Access permissions added to $($mObj.Identity)" -Severity Debug
        }
    
        if (($LeaveExistingDelegates.IsPresent -eq $false) -and (($extraPermissions.Count + 0) -gt 0)) {
        
            Write-Log "Removing $($extraPermissions.Count) extra permission(s) based on group membership" -Severity Debug
            foreach ($extra in $extraPermissions) {
                $u = Get-MgUser -UserId $extra.id
                Write-Log "`tRemoving Full Access $($u.DisplayName) permissions from $($mObj.Identity) $DelegateMailbox..." -Severity Debug
                Remove-MailboxPermission -Identity $mObj.identity -User $extra.Id -Confirm:$false -AccessRights "FullAccess" | Out-Null
            } 
        }
        else {
            Write-Log "No Full Access permissions removed from $($mObj.Identity)." -Severity Debug
        }
       
    }

    # If need to add SendAs
    if ($Permissions -contains "SendAs") {
        
        $existingSendAsPermissions = Get-ExoRecipientPermission -Identity $mObj.identity | Where-Object { $_.Trustee -like "*@*" -and $_.AccessControlType -eq "Allow" -and $_.AccessRights -contains "SendAs" } | Sort-Object -Property Trustee -Unique | ForEach-Object { Get-MgUser -UserId $_.Trustee }
        if ($gMembers) {
            $cPermissions = Compare-Object -ReferenceObject $existingSendAsPermissions -DifferenceObject $gMembers -Property Id -ErrorAction SilentlyContinue
        }
        $missingPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ "=>"
        $extraPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ "<="
        if (($missingPermissions.Count + 0) -gt 0) {
            Write-Log "Adding $($missingPermissions.Count) missing permission(s) based on group membership" -Severity Information
        
            foreach ($missing in $missingPermissions) {
                $u = Get-MgUser -UserId $missing.id
                Write-Log "`tAdding SendAs permissions for $($u.DisplayName) to $($mObj.Identity) $DelegateMailbox..." -Severity Debug
                Add-RecipientPermission -Identity $mObj.Id -Trustee $missing.Id -AccessRights 'SendAs' -Confirm:$false | Out-Null
            }
        }
        else {
            # Write-Log "No Send As permissions added to $DelegateMailbox" -Severity Debug
        }
    
        if (($LeaveExistingDelegates.IsPresent -eq $false) -and (($extraPermissions.Count + 0) -gt 0)) {
    
            Write-Log "Removing $($extraPermissions.Count) extra permission(s) based on group membership" -Severity Information
            foreach ($extra in $extraPermissions) {
                $u = Get-MgUser -UserId $extra.id
                Write-Log "`tRemoving SendAs permissions for $($u.DisplayName) to $($mObj.Identity) $DelegateMailbox..." -Severity Debug
                Remove-RecipientPermission -Identity $mObj.identity -Trustee $extra.Id -Confirm:$false -AccessRights "SendAs" | Out-Null
            } 
        }
        else {
            # Write-Log "No Send As permissions removed from $DelegateMailbox." -Severity Debug
        }
    }
}

#Import-Module Microsoft.Graph.Users | Out-Null
#Import-Module Microsoft.Graph.Users.Actions | Out-Null
#Import-Module Microsoft.Graph.Identity.DirectoryManagement | Out-Null
#Import-Module ExchangeOnlineManagement
#Import-Module PSTeams | Out-Null

Write-Log -Message "Executing Connect-MgGraph -TenantId $TenantID" -Severity Debug

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process

# Connect to Azure with system-assigned managed identity
$AzureContext = (Connect-AzAccount -Identity).context

# Set and store context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext

# Connect to Microsoft Graph
Connect-MgGraph -Identity -NoWelcome
$testUser = Get-MgUser -UserId $AdminEmailAddress
if ([string]::IsNullOrWhiteSpace($testUser)) {
    Write-Log -Message "Unable to obtain user information using Get-MgUser -UserId $AdminEmailAddress" -Severity Error
    exit 1
}

if ([string]::IsNullOrWhiteSpace($BambooHrApiKey)) {  
    $BambooHrApiKey = Get-AutomationVariable -Name 'BambooHrApiKey'
}
if ([string]::IsNullOrWhiteSpace($BhrCompanyName)) {  
    $BhrCompanyName = Get-AutomationVariable -Name 'BhrCompanyName'
}
if ([string]::IsNullOrWhiteSpace($CompanyName)) {  
    $CompanyName = Get-AutomationVariable -Name 'CompanyName'
}
if ([string]::IsNullOrWhiteSpace($TeamsCardUri)) {  
    $TeamsCardUri = Get-AutomationVariable -Name 'TeamsCardUri'
}
if ([string]::IsNullOrWhiteSpace($TenantId)) {  
    $TenantId = Get-AutomationVariable -Name 'TenantId'
}

# Check if variables are not set. If there is an environment variable, set its value to the variable. Used as an Azure Function
if ([string]::IsNullOrWhiteSpace($BambooHrApiKey) -and [string]::IsNullOrWhiteSpace($env:BambooHrApiKey)) {
    Write-Log "BambooHR API Key not defined" -Severity Error
    New-AdaptiveCard {  
        New-AdaptiveTextBlock -Text "BambooHR API Key not defined" -Weight Bolder -Wrap     
    } -Uri $TeamsCardUri -Speak "BhrAadSync error: BambooHR API Key not defined"
    exit
}
elseif ([string]::IsNullOrWhiteSpace($AdminEmailAddress) -and (-not [string]::IsNullOrWhiteSpace($env:BambooHrApiKey))) {
    $BambooHrApiKey = $env:BambooHrApiKey
}

if ([string]::IsNullOrWhiteSpace($AdminEmailAddress) -and [string]::IsNullOrWhiteSpace($env:AdminEmailAddress)) {
    Write-Log "Admin Email Address not defined" -Severity Error
    New-AdaptiveCard {  
        New-AdaptiveTextBlock -Text "Admin Email Address not defined" -Weight Bolder -Wrap    
    } -Uri $TeamsCardUri -Speak "BhrAadSync error: Admin Email Address not defined"
    exit
}
elseif ([string]::IsNullOrWhiteSpace($AdminEmailAddress) -and (-not [string]::IsNullOrWhiteSpace($env:AdminEmailAddress))) {
    $AdminEmailAddress = $env:AdminEmailAddress
}

if ([string]::IsNullOrWhiteSpace($BHRCompanyName) -and [string]::IsNullOrWhiteSpace($env:BHRCompanyName)) {
    Write-Log "BambooHR company name not defined" -Severity Error
    New-AdaptiveCard {  
        New-AdaptiveTextBlock -Text "BambooHR company name not defined" -Weight Bolder -Wrap    
    } -Uri $TeamsCardUri -Speak "BhrAadSync error: BambooHR company name not defined"
    exit
}
elseif ([string]::IsNullOrWhiteSpace($BHRCompanyName) -and (-not [string]::IsNullOrWhiteSpace($env:BHRCompanyName))) {
    $BHRCompanyName = $env:BHRCompanyName
}

if ([string]::IsNullOrWhiteSpace($CompanyName) -and [string]::IsNullOrWhiteSpace($env:CompanyName)) {
    Write-Log "Company name not defined" -Severity Error
    New-AdaptiveCard {  
        New-AdaptiveTextBlock -Text "Company name not defined" -Weight Bolder -Wrap    
    } -Uri $TeamsCardUri -Speak "BhrAadSync error: Company name not defined"
    exit
}
elseif ([string]::IsNullOrWhiteSpace($CompanyName) -and (-not [string]::IsNullOrWhiteSpace($env:CompanyName))) {
    $CompanyName = $env:CompanyName
}

if ([string]::IsNullOrWhiteSpace($TenantID) -and [string]::IsNullOrWhiteSpace($env:TenantID)) {
    Write-Log "TenantId not defined" -Severity Error
    New-AdaptiveCard {  
        New-AdaptiveTextBlock -Text "TenantId not defined" -Weight Bolder -Wrap    
    } -Uri $TeamsCardUri -Speak "BhrAadSync error: TenantId not defined"
    exit
}
elseif ([string]::IsNullOrWhiteSpace($TenantID) -and (-not [string]::IsNullOrWhiteSpace($env:TenantID))) {
    $TenantID = $env:TenantID
    $env:AZURE_TENANT_ID = $TenantId
}

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
    Write-Log -Message "Starting BambooHR to Entra AD synchronization at $(Get-Date)" -Severity Debug
    # Provision users to AAD using the employee details from BambooHR
    

    #Connect-MgGraph -TenantId $TenantID -CertificateThumbprint $AADCertificateThumbprint -ClientId $AzureClientAppId
    # Getting all users details from BambooHR and passing the extracted info to the variable $employees
    
    $headers = @{}
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Authorization", "Basic $([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($BambooHrApiKey):x")))")
    $error.clear()

    try {
        Invoke-RestMethod -Uri $bhrReportsUri -Method POST -Headers $headers -ContentType 'application/json' `
            -Body '{"fields":["status","hireDate","department","employeeNumber","firstName","lastName","displayName","jobTitle","supervisorEmail","workEmail","lastChanged","employmentHistoryStatus","bestEmail","location","workPhone","preferredName","homeEmail","mobilePhone"]}' `
            -OutVariable response 
    } 
    catch {
        # If error returned, the API call to BambooHR failed and no usable employee data has been returned, write to log file and exit script
             
        Write-log -Message "Error calling BambooHr API for user information. `nEXCEPTION MESSAGE: $($_.Exception.Message) `n CATEGORY: $($_.CategoryInfo.Category) `n SCRIPT STACK TRACE: $($_.ScriptStackTrace)" -Severity Error
        #Send email alert with the generated error
        $params = @{
            Message         = @{
                Subject      = "BhrAadSync error: BambooHR connection failed"
                Body         = @{
                    ContentType = "html"
                    Content     = "BambooHR connection failed. <br/> EXCEPTION MESSAGE: $($_.Exception.Message) <br/>CATEGORY: $($_.CategoryInfo.Category) <br/> SCRIPT STACK TRACE: $($_.ScriptStackTrace) `n $EmailSignature"
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

        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose
        New-AdaptiveCard {
            New-AdaptiveTextBlock -Text "BambooHR API connection failed!" -Weight Bolder -Wrap -Color Red
            New-AdaptiveTextBlock -Text "Exception Message $($_.Exception.Message)" -Wrap
            New-AdaptiveTextBlock -Text "Category: $($_.CategoryInfo.Category)" -Wrap
            New-AdaptiveTextBlock -Text "SCRIPT STACK TRACE: $($_.ScriptStackTrace)" -Wrap
        } -Uri $TeamsCardUri -Speak "BhrAadSync error: BambooHR connection failed"
        #Send-TeamsCard -CardSubject "BhrAadSync error: BambooHR connection failed"  -Message "BambooHR connection failed. <br/> EXCEPTION MESSAGE: $($_.Exception.Message) <br/>CATEGORY: $($_.CategoryInfo.Category) <br/> SCRIPT STACK TRACE: $($_.ScriptStackTrace) `n $EmailSignature"
        exit
    }

    # If no error returned, it means that the script was not interrupted by the "Exit" command within the "Catch" block. Write info below to log file and continue
    Write-Log -Message "Successfully extracted employee information from BambooHR." -Severity Debug

    # Saving only the employee data to $employees variable and eliminate $response variable to save memory
    $employees = $response.employees
    $response = $null

    # Connect to AAD using PS Graph Module, authenticating as the configured service principal for this operation, with certificate auth
    $error.Clear()

    if ($?) {
        # If no error returned, write to log file and continue
        Write-Log -Message "Successfully connected to AAD: $TenantId." -Severity Debug
    }
    else {

        # If error returned, write to log file and exit script
        Write-Log -Message "Connection to AAD failed.`n EXCEPTION: $($error.Exception) `n CATEGORY: $($error.CategoryInfo) `n ERROR ID: $($error.FullyQualifiedErrorId) `n SCRIPT STACK TRACE: $($error.ScriptStackTrace)" -Severity Error

        # Send email alert with the generated error 
        $params = @{
            Message         = @{
                Subject      = "BhrAadSync error: Graph connection failed"
                Body         = @{
                    ContentType = "html"
                    Content     = "<br/><br/>AAD connection failed.<br/>EXCEPTION: $($error.Exception) <br/> CATEGORY:$($error.CategoryInfo) <br/> ERROR ID: $($error.FullyQualifiedErrorId) <br/>SCRIPT STACK TRACE: $mgErrStack <br/> $EmailSignature"
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

        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose

        New-AdaptiveCard {  

            New-AdaptiveTextBlock -Text "AAD Connection Failed" -Weight Bolder -Wrap
            New-AdaptiveTextBlock -Text "Exception Message $($_.Exception.Message)" -Wrap
            New-AdaptiveTextBlock -Text "Category: $($_.CategoryInfo.Category)" -Wrap
            New-AdaptiveTextBlock -Text "ERROR ID: $($error.FullyQualifiedErrorId)" -Wrap
            New-AdaptiveTextBlock -Text "SCRIPT STACK TRACE: $($_.ScriptStackTrace)" -Wrap        
        
        } -Uri $TeamsCardUri -Speak "BhrAadSync error: Graph connection failed"

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
        $bhrEmployeeId = "$($_.id)"
        # Translating user "status" from BambooHR to boolean, to match and compare with the AAD user account status
        $bhrStatus = "$($_.status)"
        if ($bhrStatus -eq "Inactive")
        { $bhrAccountEnabled = $False }
        if ($bhrStatus -eq "Active")
        { $bhrAccountEnabled = $True }
        $bhrOfficeLocation = "$($_.location)"
        $bhrPreferredName = "$($_.preferredName)"
        $bhrWorkPhone = "$($_.workPhone)"
        $bhrMobilePhone = "$($_.mobilePhone)"
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
        
        <#
            If the user start date is in the past, or in less than -DaysAhead days from current time, we can begin processing the user: 
            create AAD account or correct the attributes in AAD for the employee, else, the employee found on BambooHR will not be processed
        #>

        if (([datetime]$bhrHireDate) -le (Get-Date).AddDays($DaysAhead)) {

            $error.clear()

            # Check if the user exists in AAD and if there is an account with the EmployeeID of the user checked in the current loop
            Write-Log -Message "Validating $bhrWorkEmail AAD account." -Severity Debug
            Get-MgUser -UserId $bhrWorkEmail -Property id, userprincipalname, Department, EmployeeId, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, EmployeeHireDate, OfficeLocation, BusinessPhones, MobilePhone, OnPremisesExtensionAttributes, AdditionalProperties -ExpandProperty manager -ErrorAction SilentlyContinue -OutVariable aadUpnObjDetails
            Get-MgUser -Filter "employeeID eq '$bhrEmployeeNumber'" -Property employeeid, userprincipalname, Department, JobTitle, CompanyName, Surname, GivenName, DisplayName, MobilePhone, AccountEnabled, Mail, OfficeLocation, BusinessPhones , EmployeeHireDate, OnPremisesExtensionAttributes, AdditionalProperties -ExpandProperty manager -ErrorAction SilentlyContinue -OutVariable aadEidObjDetails
            $error.clear()

            if ([string]::IsNullOrEmpty($aadUpnObjDetails) -eq $false) {
                $UpnExtensionAttribute1 = ($aadUpnObjDetails | Select-Object @{Name = 'ExtensionAttribute1'; Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } } -ErrorAction SilentlyContinue).ExtensionAttribute1
            }

            if ([string]::IsNullOrEmpty($aadEidObjDetails) -eq $false) {
                $EIDExtensionAttribute1 = ($aadEidObjDetails | Select-Object @{Name = 'ExtensionAttribute1'; Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } } -ErrorAction SilentlyContinue).ExtensionAttribute1
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
            $aadMobilePhone = "$($aadUpnObjDetails.MobilePhone)"
            $aadOfficeLocation = "$($aadUpnObjDetails.OfficeLocation)"

            # Clean up phone info to make it easier to work with
            [string]$bhrWorkPhone = [int64]($bhrWorkPhone -replace '[^0-9]', '') -replace '^1', ''
            [string]$aadWorkPhone = [int64]($aadWorkPhone -replace '[^0-9]', '') -replace '^1', ''
            [string]$bhrMobilePhone = [int64]($bhrMobilePhone -replace '[^0-9]', '') -replace '^1', ''
            [string]$aadMobilePhone = [int64]($aadMobilePhone -replace '[^0-9]', '') -replace '^1', ''

            if ($aadUpnObjDetails.EmployeeHireDate) {
                $aadHireDate = $aadUpnObjDetails.EmployeeHireDate.AddHours(12).ToString("yyyy-MM-dd") 
            }        

            Write-Log -Message "AAD Upn Obj Details: '$([string]::IsNullOrEmpty($aadUpnObjDetails) -eq $false)' AadEidObjDetails: $([string]::IsNullOrEmpty($aadEidObjDetails) -eq $false) = $(([string]::IsNullOrEmpty($aadUpnObjDetails) -eq $false) -or ([string]::IsNullOrEmpty($aadEidObjDetails) -eq $false))" -Severity Debug

            if (([string]::IsNullOrEmpty($aadUpnObjDetails) -eq $false) -or ([string]::IsNullOrEmpty($aadEidObjDetails) -eq $false)) {
                Write-Log -Message "Entra Id user: $aadFirstName $aadLastName ($aadDisplayName) $aadWorkEmail" -Severity Debug 
                Write-Log -Message "Department: $aadDepartment, Title: $aadJobTitle, Manager: $aadSupervisorEmail, HireDate: $aadHireDate" -Severity Debug
                Write-Log -Message "EmployeeId: $aadEmployeeNumber, Enabled: $aadStatus OfficeLocation: $aadOfficeLocation, WorkPhone: $aadWorkPhone" -Severity Debug

                <# If empID of object returned by UPN or by empID is equal, and if object ID is the same as object ID of object returned by UPN and EmpID and
            if UPN = workemail from bamboo AND if the last changed date from BambooHR is NOT equal to the last changed date saved in ExtensionAttribute1 in AAD, 
            check each attribute and set them correctly, according to BambooHR
            #>

                Write-Log -Message "Entra Id Employee Number: $aadEmployeeNumber -eq $aadEmployeeNumber2 = $($aadEmployeeNumber -eq $aadEmployeeNumber2) -and `
            $($aadEidObjDetails.UserPrincipalName) -eq $($aadUpnObjDetails.UserPrincipalName) -eq $bhrWorkEmail = $($aadEidObjDetails.UserPrincipalName -eq $aadUpnObjDetails.UserPrincipalName -eq $bhrWorkEmail) -and `
            $($aadUpnObjDetails.id) -eq $($aadEidObjDetails.id) = $($aadUpnObjDetails.id -eq $aadEidObjDetails.id) -and `
            $bhrLastChanged -ne $UpnExtensionAttribute1 = $($bhrLastChanged -ne $UpnExtensionAttribute1) -and `
            $($aadEidObjDetails.Capacity) -ne 0 -and $($aadUpnObjDetails.Capacity) -ne 0 = $($aadEidObjDetails.Capacity -ne 0 -and $aadUpnObjDetails.Capacity -ne 0) -and `
            $bhrEmploymentStatus -notlike '*suspended*' = $($bhrEmploymentStatus -notlike "*suspended*") " -Severity Debug

                # This may not be needed anymore.
                if (($aadEmployeeNumber -ne $bhrEmployeeNumber) -and ($aadUpnObjDetails.UserPrincipalName -eq $bhrWorkEmail) -and `
                        $bhrEmploymentStatus -notlike "*suspended*" -and $bhrLastChanged -ne $UpnExtensionAttribute1) { 
                    # Employee number in Entra Id does not match the one in BambooHR, but the UPN matches. Update the employee number in AAD.
                    Write-Log -Message "Entra Id Employee number $aadEmployeeNumber does not match BambooHR $bhrEmployeeNumber, but the UPN matches. Update the employee number in AAD." -Severity Debug
                    $error.clear()
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber" -Severity Test
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber" -Severity Debug
                        Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber
                        $aadEmployeeNumber = $bhrEmployeeNumber
                    }
                }

                if ($aadEmployeeNumber -eq $aadEmployeeNumber2 -or `
                    (($aadEidObjDetails.UserPrincipalName -eq $bhrWorkEmail) -or 
                        ($aadUpnObjDetails.UserPrincipalName -eq $bhrWorkEmail)) -and `
                        #$aadUpnObjDetails.id -eq $aadEidObjDetails.id -and `
                        $bhrLastChanged -ne $UpnExtensionAttribute1 -and `
                    ($aadEidObjDetails.Capacity -ne 0) -or ($aadUpnObjDetails.Capacity -ne 0) -and `
                        $bhrEmploymentStatus -notlike "*suspended*" ) { 

                    Write-Log -Message "$bhrWorkEmail is a valid AAD Account, with matching EmployeeId and UPN in AAD and BambooHR, but different last modified date." -Severity Debug
                    $error.clear() 

                    # Check if user is active in BambooHR, and set the status of the account as it is in BambooHR (active or inactive)
                    if ($bhrAccountEnabled -eq $false -and $bhrEmploymentStatus.Trim() -eq "Terminated" -and $aadStatus -eq $true ) {
                        Write-Log -Message "$bhrWorkEmail is marked 'Inactive' in BHR and 'Active' in Entra Id (AAD). Blocking sign-in, revoking sessions, changing pw, removing auth methods"
                        # The account is marked "Inactive" in BHR and "Active" in AAD, block sign-in, revoke sessions, change pass, remove auth methods
                        $error.clear()
                        if ($TestOnly.IsPresent) {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Test
                            Write-Log -Message "Executing: Revoke-MgUserSignInSession -UserId $bhrWorkEmail" -Severity Test
                        }
                        else {
                            Write-Log -Message "Executing: Revoke-MgUserSignInSession -UserId $bhrWorkEmail" -Severity Debug
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Debug
                            Revoke-MgUserSignInSession -UserId $bhrWorkEmail
                            Start-Sleep 10
                            Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled

                        }

                        # Change to a random password that is not known to the user.
                        $params = @{
                            PasswordProfile = @{
                                ForceChangePasswordNextSignIn = $true
                                Password                      = (Get-NewPassword)

                            }
                        }

                        if ($TestOnly.IsPresent) {
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params"  -Severity Test
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -Department 'Not Active' -JobTitle 'Not Active $(Get-Date)' -OfficeLocation 'Not Active' -BusinessPhones '0' -MobilePhone '0' -CompanyName 'Not Active'"  -Severity Test
                            Write-Log -Message "Executing: Get-MgUserMemberOf -UserId $bhrWorkEmail"  -Severity Test
                            Write-Log -Message "Executing: $null = Update-MgUser -UserId $bhrWorkEmail -OnPremisesExtensionAttributes @{extensionAttribute1 = '$bhrLastChanged' }" -Severity Test

                            Write-Log -Message "Converting $bhrWorkEmail to a shared mailbox..." -Severity Test
                            Write-Log -Message "Executing: Set-Mailbox -Identity $bhrWorkEmail -Type Shared" -Severity Test
                            # Give permissions to converted mailbox to previous manager
                            Write-log "Executing: Add-MailboxPermission -Identity $($mObj.Id) -User $aadSupervisorEmail -AccessRights ‘FullAccess’ -Automapping:$true –inheritancetype All" -Severity Test                            

                            Write-Log -Message "Removing licenses..." -Severity Test
                            Write-Log -Message "Executing: Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }" -Severity Test

                            Write-Log -Message "Remove group memberships" -Severity Test
                            Write-Log -Message "Executing: Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $aadUpnObjDetails.id } " -Severity Test

                            Write-Log -Message "Remove MFA auth for user" -Severity Test

                        }
                        else {

                            Write-Log -Message "User $bhrWorkEmail is no longer active in BambooHR, disabling Entra Id (AAD) account." -Severity Information
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params"  -Severity Debug
                            Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params
                            Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -Department 'Not Active' -JobTitle 'Not Active' -OfficeLocation 'Not Active' -BusinessPhones '0' -MobilePhone '0' -CompanyName '$(Get-Date -Uformat %D)'"  -Severity Debug
                            Update-MgUser -UserId $bhrWorkEmail -Department "Not Active" -JobTitle "Not Active" -OfficeLocation "Not Active" -BusinessPhones "0" -MobilePhone "0" -CompanyName "$(Get-Date -Uformat %D)"
                            Get-MgUserMemberOf -UserId $bhrWorkEmail

                            # TODO: Does not work for on premises synced accounts. Not a problem with Entra Id (AAD) native.
                            $null = Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrLastChanged } -UserId $bhrWorkEmail -ErrorAction SilentlyContinue | Out-Null

                            if (!$?) {
                                #Write-Log -Message "Error changing ExtensionAttribute1. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nTrace: $($error.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                            }
                            else {
                                Write-Log -Message "$bhrWorkEmail LastChanged attribute set from '$upnExtensionAttribute1' to '$bhrlastChanged'." -Severity Information
                            }

                            # Convert mailbox to shared
                            Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false | Out-Null

                            Write-Log -Message "Executing: Set-Mailbox -Identity $bhrWorkEmail -Type Shared" -Severity Debug
                            Write-Log -Message "Converting $bhrWorkEmail to a shared mailbox..." -Severity Debug
                            Set-Mailbox -Identity $bhrWorkEmail -Type Shared
                            # Wait for mailbox to be converted
                            Start-Sleep 60
                            # Give permissions to converted mailbox to previous manager
                            $mObj = Get-ExoMailbox -anr $bhrWorkEmail
                            Write-Log "`t$($aadSupervisorEmail) being given permissions to $bhrWorkEmail now..." -Severity Information
                            Write-log "Executing: Add-MailboxPermission -Identity $($mObj.Id) -User $aadSupervisorEmail -AccessRights ‘FullAccess’ -Automapping:$true –inheritancetype All" -Severity Debug
                            Add-MailboxPermission -Identity $mObj.Id -User $aadSupervisorEmail -AccessRights ‘FullAccess’ -Automapping:$true –inheritancetype All | Out-Null
                            Disconnect-ExchangeOnline -Confirm:$False

                            # Move OneDrive for Business content to archive location based on department
                            # TODO

                            # Set Out of Office for user
                            # TODO

                            # Cancel Meetings
                            # TODO

                            # If user was a group owner, reassign ownership to someone else
                            # TODO

                            # Reset/wipe the employees device(s)
                            # TODO

                            # Remove Licenses
                            Write-Log -Message "Removing licenses..." -Severity Debug

                            Write-Log -Message "Executing: Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }" -Severity Debug
                            Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }

                            # Remove groups
                            Write-Log -Message "Removing group memberships" -Severity Debug
                            Write-Log -Message "Executing: Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $aadUpnObjDetails.id } " -Severity Debug

                            Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $aadUpnObjDetails.id -ErrorAction SilentlyContinue; Start-Sleep 10 } 
                            $methodID = Get-MgUserAuthenticationMethod -UserId $bhrWorkEmail | Select-Object id 
                            $methodsdata = Get-MgUserAuthenticationMethod -UserId $bhrWorkEmail | Select-Object -ExpandProperty AdditionalProperties
                            $methods_count = ($methodID | Measure-Object | Select-Object count).count

                            # Loop through and remove each authentication method
                            $error.Clear() 

                            for ($i = 0 ; $i -lt $methods_count ; $i++) {
       
                                if ((($methodsdata[$i]).Values) -like "*phoneAuthenticationMethod*") { Remove-MgUserAuthenticationPhoneMethod -UserId $bhrWorkEmail -PhoneAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed phone auth method for $bhrWorkEmail." -Severity Warning }
                                if ((($methodsdata[$i]).Values) -like "*microsoftAuthenticatorAuthenticationMethod*") { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $bhrWorkEmail -MicrosoftAuthenticatorAuthenticationMethodId ($methodID[$i]).id; Write-Log -Message "Removed auth app method for $bhrWorkEmail." -Severity Warning }
                                if ((($methodsdata[$i]).Values) -like "*windowsHelloForBusinessAuthenticationMethod*") { Remove-MgUserAuthenticationFido2Method -UserId $bhrWorkEmail -Fido2AuthenticationMethodId ($methodID[$i]).id ; Write-Log -Message "Removed PIN auth method for $bhrWorkEmail." -Severity Warning }
                            }

                            # Remove Manager
                            Write-Log -Message "Removing Manager..." -Severity Debug   
                            Write-Log -Message "Executing: Remove-MgUserManagerByRef -UserId $bhrWorkEmail" -Severity Debug
                            Remove-MgUserManagerByRef -UserId $bhrWorkEmail

                            Write-Log -Message "Executing: Update-MgUser -EmployeeId 'LVR' -UserId $bhrWorkEmail" -Severity Debug
                            Update-MgUser -EmployeeId 'LVR' -UserId $bhrWorkEmail
                            Write-Log -Message "Updating shared mailbox settings..."

                            if ($error.Count -ne 0) {
                                $error | ForEach-Object {
                                    $err_Exception = $_.Exception
                                    $err_Target = $_.TargetObject
                                    $errCategory = $_.CategoryInfo
                                    Write-Log " Could not remove authentication details. `n Exception: $err_Exception `n Target Object: $err_Target `n Error Category: $errCategory " -Severity Error
                                }
                            }
                            else {
                                Write-Log -Message " Account $bhrWorkEmail marked as inactive in BambooHR AAD account has been disabled, sessions revoked and removed MFA." -Severity Information              
                                $error.Clear()
                            }
                        }
                    }
                    elseif ($bhrAccountEnabled -eq $false -and $bhrEmploymentStatus.Trim() -eq "Terminated" -and $aadStatus -eq $false ) {
                        #Account is disabled and there is nothing else to do
                    }
                    else {
                        Write-Log "User account active, looking for user updates." -Severity Debug
  
                        if ($bhrAccountEnabled -eq $true -and $aadstatus -eq $false) {
                            # The account is marked "Active" in BHR and "Inactive" in AAD, enable the AAD account
                            Write-Log -Message "$bhrWorkEmail is marked Active in BHR and Inactive in AAD" -Severity Debug

                            #Change to a random pass
                            $newPas = (Get-NewPassword)
                            $params = @{
                                PasswordProfile = @{
                                    ForceChangePasswordNextSignIn = $true
                                    Password                      = $newPas
                                }
                            }

                            if ($TestOnly.IsPresent) {
                            
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Test
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params" -Severity Test
                                Write-Log -Message "Executing: Set-Mailbox -Identity $bhrWorkEmail -Type Regular" -Severity Test
                                Write-log "Executing: Remove-MailboxPermission -Identity $($mObj.Id) -ResetDefault" -Severity Debug
                            }
                            else {
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Debug                         
                                Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params" -Severity Debug
                                Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params
                            
                                # Convert mailbox from shared to user mailbox
                                Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false | Out-Null

                                Write-Log -Message "Executing: Set-Mailbox -Identity $bhrWorkEmail -Type Regular" -Severity Debug
                                Write-Log -Message "Converting $bhrWorkEmail to a user mailbox..." -Severity Debug
                                Set-Mailbox -Identity $bhrWorkEmail -Type Regular

                                # Wait for mailbox to be converted
                                Start-Sleep 60

                                # Remove permissions to converted mailbox to previous manager
                                $mObj = Get-ExoMailbox -anr $bhrWorkEmail
                                Write-Log "`tShared permissions being revoked for $bhrWorkEmail..." -Severity Information
                                Write-log "Executing: Remove-MailboxPermission -Identity $($mObj.Id) -ResetDefault" -Severity Debug
                                Remove-MailboxPermission -Identity $mObj.Id -ResetDefault | Out-Null
                                Disconnect-ExchangeOnline -Confirm:$False

                                $params = @{
                                    Message         = @{
                                        Subject      = "User Account Re-enabled: $bhrdisplayName"
                                        Body         = @{
                                            ContentType = "html"
                                            Content     = "<br/>One of your direct report's user account has been re-enabled. Please securely share this information with them so that they can login.<br/> User name: $bhrWorkEmail <br/> Temporary Password: $newPas.`n<br/><br/> $EmailSignature"
                                        }
                                        ToRecipients = @(
                                            @{
                                                EmailAddress = @{
                                                    Address = $bhrSupervisorEmail
                                                }
                                            }
                                        
                                            @{
                                                EmailAddress = @{
                                                    Address = $NotificationEmailAddress
                                                }
                                            }
                                            @{
                                                EmailAddress = @{
                                                    Address = $AdminEmailAddress
                                                }        
                                            }
                                        )
                                    }
                                    SaveToSentItems = "True"
                                }

                                Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose
                               
                                New-AdaptiveCard {  

                                    New-AdaptiveTextBlock -Text "User Account $bhrWorkEmail Re-enabled" -HorizontalAlignment Center -Wrap -Weight Large
                                    New-AdaptiveTextBlock -Text "User name: $bhrWorkEmail" -Wrap
                                    New-AdaptiveTextBlock -Text "Temporary Password: $newPas" -Wrap     
        
                                } -Uri $TeamsCardUri -Speak "User Account Re-enabled: $bhrdisplayName"
            

                                if (!$?) {
    
                                    Write-Log -Message " Could not activate user account. `n`nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                }
                                else {
                                    Write-Log -Message " Account $bhrWorkEmail marked as Active in BambooHR but Inactive in AAD. Enabled AAD account for sign-in." -Severity Information
                                    $error.Clear()
                                }                   
                            }
                        }
                        else {
                            Write-Log -Message "Account is in the correct state: Enabled in both BHR and Entra Id (AAD)" -Severity Debug
                        }
                    
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
                                    Write-Log -Message "JobTitle for $bhrWorkEmail in AAD set from '$aadjobTitle' to '$bhrjobTitle'." -Severity Information
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
                                    Write-Log -Message "Department for $bhrWorkEmail in AAD set from '$aadDepartment' to '$bhrDepartment'." -Severity Information
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
                                    Write-Log -Message "Manager of $bhrWorkEmail in AAD '$aadsupervisorEmail' and in BambooHR '$bhrsupervisorEmail'. Setting new manager to the Azure User Object." -Severity Information
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
                                    Write-Log -Message "Office location of $bhrWorkEmail in AAD changed from '$aadOfficeLocation' to '$bhrOfficeLocation'." -Severity Information
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
                
                                    Write-Log -Message "Error changing $bhrWorkEmail hire date. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                }
                                else {
                                    $error.Clear()
                                    Write-Log -Message "Hire date of $bhrWorkEmail changed from '$aadHireDate' in AAD and BHR '$bhrHireDate'." -Severity Information
                                }
                            }
                        }
                        else {
                            Write-Log -Message "Hire date already correct $aadHireDate" -Severity Debug
                        }
                        
                        
                        # Check and set the work phone ignoring formatting
                        if (($aadWorkPhone) -ne ($bhrWorkPhone)) {

                            Write-Log -Message "AAD work phone '$aadWorkPhone' does not match BHR '$bhrWorkPhone'" -Severity Debug
                            if ([string]::IsNullOrWhiteSpace($bhrWorkPhone)) {
                                $bhrWorkPhone = "0"
                            }

                            if ($TestOnly.IsPresent) {
                                       
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone" -Severity Test
                            }
                            else {

                                if ([string]::IsNullOrWhiteSpace($bhrWorkPhone)) {
                                    $bhrWorkPhone = "0"
                                    Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones '$bhrWorkPhone'" -Severity Debug
                                    Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone -ErrorAction SilentlyContinue | Out-Null

                                }
                                else {
                                    [string]$bhrWorkPhone = [int64]($bhrWorkPhone -replace '[^0-9]', '') -replace '^1', ''
                                    Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone" -Severity Debug
                                    Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone -ErrorAction SilentlyContinue | Out-Null
                                    
                                }

                                if (!$?) {
                
                                    Write-Log -Message "Error changing work phone for $bhrWorkEmail. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                }
                                else {
                                    $error.Clear()
                                    Write-Log -Message "Work Phone for '$bhrWorkEmail' changed from '$aadWorkPhone' to '$bhrWorkPhone'" -Severity Information
                                }
                            }
                        }
                        else {
                            Write-Log -Message "Work phone correct $aadWorkEmail $aadWorkPhone" -Severity Debug
                        }

                        if ($EnableMobilePhoneSync.IsPresent) {
                            [string]$aadMobilePhone = $aadMobilePhone -replace '[^0-9]', ''
                            [string]$bhrMobilePhone = $bhrMobilePhone -replace '[^0-9]', ''
                            # Check and set the mobile phone ignoring formatting
                            if ($aadMobilePhone -ne $bhrMobilePhone) {

                                Write-Log -Message "AAD mobile phone '$aadWorkPhone' does not match BHR '$bhrMobilePhone'" -Severity Debug
                        
                                if ($TestOnly.IsPresent) {
                                                               
                                    Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -MobilePhone $bhrMobilePhone" -Severity Test
                                }
                                else {
                                    if ([string]::IsNullOrWhiteSpace($bhrMobilePhone)) {
                                        $bhrMobilePhone = "0"
                                        Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -MobilePhone '$bhrMobilePhone'" -Severity Debug
                                        Update-MgUser -UserId $bhrWorkEmail -MobilePhone $bhrMobilePhone -ErrorAction Continue
                                    }
                                    else {
                                        $bhrMobilePhone = ($bhrMobilePhone -replace '[^0-9]', '' ) -replace '^1', ''
                                        $bhrMobilePhone = "{0:(###) ###-####}" -f [int64]$bhrMobilePhone
                                        Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -MobilePhone $bhrMobilePhone" -Severity Debug
                                        if ($bhrWorkEmail -notlike "rherndon*") {
                                            Update-MgUser -UserId $bhrWorkEmail -MobilePhone $bhrMobilePhone -ErrorAction Continue
                                        }
                                    }
                        
                                    if (!$?) {
                                        
                                        Write-Log -Message "Error changing $bhrWorkEmail mobile phone. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                        $error.Clear()
                                    }
                                    else {
                                        $error.Clear()
                                        Write-Log -Message "Work Mobile Phone for '$bhrWorkEmail' changed from '$aadMobilePhone' to '$bhrMobilePhone'" -Severity Deb
                                    }
                                }
                            }
                            else {
                                Write-Log -Message "Mobile phone correct for $aadWorkEmail $aadMobilePhone" -Severity Debug
                            }
                        }

                        # Compare user employee id with BambooHR and set it if not correct
                        if ($bhrEmployeeNumber.Trim() -ne $aadEmployeeNumber.Trim()) {
                            Write-Log -Message " BHR employee number $bhrEmployeeNumber does not match AAD employee id $aadEmployeeNumber" -Severity Debug
                            if ($TestOnly.IsPresent) {
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber  "

                            }
                            else {
                                # Setting the Employee ID found in BHR to the user in AAD
                                Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber.Trim()             
                                if (!$?) {

                                    Write-Log -Message " Error changing EmployeeId. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                }
                                else {
                                    Write-Log -Message " The ID $bhremployeeNumber has been set to $bhrWorkEmail AAD account." -Severity Warning
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
                    
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -CompanyName $($CompanyName.Trim())" -Severity Information -Severity Test
                            }
                            else {
                                # Setting Company Name as $CompanyName to the employee, if not already set
                                Write-Log -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -CompanyName $($CompanyName.Trim())" -Severity Debug
                                Update-MgUser -UserId $bhrWorkEmail -CompanyName $CompanyName.Trim()
                                if (!$?) {

                                    Write-Log -Message " Could not change the Company Name of $bhrWorkEmail. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                }
                                else {
                                    Write-Log -Message " The $bhrWorkEmail employee Company attribute has been set to: $CompanyName." -Severity Information
                                }
                            }
                        }
                        else {
                            Write-Log -Message "Company name already matched in AAD and BHR $aadCompanyName" -Severity Debug
                        }

                        # Set LastModified from BambooHR to ExtensionAttribute1 in AAD

                        if ($upnExtensionAttribute1 -ne $bhrLastChanged) {
                            # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AAD
                            Write-Log -Message "AAD Extension Attribute '$upnExtensionAttribute1' does not match BHR last changed '$bhrLastChanged'" -Severity Debug
                            Write-Log -Message "Set LastModified from BambooHR to ExtensionAttribute1 in AAD" -Severity Debug
                        
                            if ($TestOnly.IsPresent) {
                    
                                Write-Log -Message "Executing: $null = Update-MgUser -UserId $bhrWorkEmail -OnPremisesExtensionAttributes @{extensionAttribute1 = = '$bhrLastChanged' }" -Severity Test
                            }
                            else {
                                Write-Log -Message "Executing: $null = Update-MgUser -UserId $bhrWorkEmail -OnPremisesExtensionAttributes @{extensionAttribute1 = '$bhrLastChanged' }" -Severity Debug
                                # TODO: Does not work for on premises synched accounts. Not a problem with AAD native.
                                $null = Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrLastChanged } -UserId $bhrWorkEmail -ErrorAction SilentlyContinue | Out-Null

                                if (!$?) {
                                    #Write-Log -Message "Error changing ExtensionAttribute1. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                }
                                else {
                                    Write-Log -Message "$bhrWorkEmail LastChanged attribute set from '$upnExtensionAttribute1' to '$bhrlastChanged'." -Severity Information
                                }
                            }

                            $error.clear()             
                        }
                        else {
                            Write-Log -Message "Attribute already matched last changed of $bhrLastChanged" -Severity Debug
                        } 

                    }
                }
            }
            else {
                Write-Log -Message "No AAD user found for $bhrWorkEmail" -Severity Debug
                
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

            }

            # Handle name changes 
            if (($aadEmployeeNumber2 -eq $bhremployeeNumber) -and ($historystatus -notlike "*inactive*") -and ($aadUpnObjDetails.id -eq $aadEidObjDetails.id)) {

                $aadUPN = $aadEidObjDetails.UserPrincipalName
                $aadObjectID = $aadEidObjDetails.id
                $aadworkemail = $aadEidObjDetails.Mail
                $aademployeeNumber = $aadEidObjDetails.EmployeeID
                $aaddisplayname = $aadEidObjDetails.displayname
                $aadfirstName = $aadEidObjDetails.GivenName
                $aadlastName = $aadEidObjDetails.Surname

                Write-Log -Message "Evaluating if AAD name change is required for $aadfirstName $aadlastName ($aaddisplayname) `n`t Work Email: $aadWorkEmail UserPrincipalName: $aadUpn EmployeeId: $aadEmployeeNumber" -Severity Debug
           
                $error.Clear()
                    
                # 3/31/2023 Is this required here or should it be handled after the name change or the next sync after the name change?
                # Set LastModified from BambooHR to ExtensionAttribute1 in AAD
                if ($EIDExtensionAttribute1 -ne $bhrlastChanged) {
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "The $bhrWorkEmail employee LastChanged attribute set to extensionAttribute1 as $bhrlastChanged." -Severity Test
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } " -Severity Test
                    }
                    else {
                        # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in AAD
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } " -Severity Debug
                        # This does not work for AD on premises synced accounts.
                        $null = Update-MgUser -UserId $aadObjectID -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } -ErrorAction SilentlyContinue | Out-Null
                                            
                    }
                }

                # Change last name in AAD         
                if ($aadLastName -ne $bhrLastName) {
                    Write-Log -Message " Last name in AAD $aadLastName does not match in BHR $bhrLastName" -Severity Debug
                    Write-Log -Message " Changing the last name of $bhrWorkEmail from $aadLastName to $bhrLastName." -Severity Debug
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Surname $bhrLastName" -Severity Test
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Surname $bhrLastName" -Severity Debug
                        Update-MgUser -UserId $aadObjectID -Surname $bhrLastName
                        
                        if (!$?) {
                              
                            Write-Log -Message "Error changing AAD Last Name.`n`nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                        }
                        else {
                            Write-Log -Message " Successfully changed the last name of $bhrWorkEmail from $aadLastName to $bhrLastName." -Severity Information
                        }
                    }
                }
            
                # Change First Name in AAD
                if ($aadfirstName -ne $bhrfirstName) {
                    Write-Log "AAD first name '$aadfirstName' is not equal to BHR first name '$bhrFirstName'" -Severity Debug
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -GivenName $bhrFirstName" -Severity Test
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -GivenName $bhrFirstName" -Severity Debug
                        Update-MgUser -UserId $aadObjectID -GivenName $bhrFirstName
                        if (!$?) {
                             
                            Write-Log -Message "Could not change the First Name of $aadObjectID. Error details below. `n`nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                        }   
                        else {
                            Write-Log -Message " Successfully changed $aadObjectID first name from $aadFirstName to $bhrFirstName." -Severity Information
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
                             
                            Write-Log -Message " Could not change the Display Name. Error details below. `n`nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                        }# Change display name - Error logging
                        else {
                            Write-Log " Display name $aadDisplayName of $aadObjectID changed to $bhrDisplayName." -Severity Information
                        }        
                    }
                }

                # Change Email Address
                if ($aadWorkEmail -ne $bhrWorkEmail) {
                    Write-Log -Message "AAD work email $aadWorkEmail does not match BHR work email $bhrWorkEmail"
                    if ($TestOnly.IsPresent) {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Mail $bhrWorkEmail"
                    }
                    else {
                        Write-Log -Message "Executing: Update-MgUser -UserId $aadObjectID -Mail $bhrWorkEmail"
                        Update-MgUser -UserId $aadObjectID -Mail $bhrWorkEmail
                        if (!$?) {
                            
                            Write-Log -Message "Error changing Email Address. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                        }
                        else {
                            # Change Email Address error logging
                            Write-Log "The current Email Address: $aadworkemail of $aadObjectID has been changed to $bhrWorkEmail." -Severity Warning
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
                             
                            Write-Log -Message " Error changing UPN for $aadObjectID. `n Exception: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                        } 
                        else {
                            Write-Log -Message " Changed the current UPN:$aadUPN of $aadObjectID to $bhrWorkEmail." -Severity Warning
                            $params = @{
                                Message         = @{
                                    Subject       = "Login changed for $bhrdisplayName"
                                    Body          = @{
                                        ContentType = "HTML"
                                        Content     = "
<p>Your email address was changed in the $CompanyName BambooHR. Your user account has been changed accordingly.</p><ui><li>Use your new user name: $bhrWorkEmail</li><li>Your password has not been modified.</li></ul><br/><p>$EmailSignature</p>"
                                    }
                                    ToRecipients  = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $bhrWorkEmail
                                            }
                                        }
                                    )
                                    CCRecipients  = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $bhrSupervisorEmail
                                            }
                                        }
                                    )
                                    BCCRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $NotificationEmailAddress
                                            }
                                        }
                                    )
                                }
                                SaveToSentItems = "True"
                            }
                                
                            Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose

                            New-AdaptiveCard {  

                                New-AdaptiveTextBlock -Text "Login changed for $bhrdisplayName" -HorizontalAlignment Center -Weight Bolder -Wrap
                                New-AdaptiveTextBlock -Text "An email address was changed in the $CompanyName BambooHR. Your user account has been changed accordingly." -Wrap
                                New-AdaptiveTextBlock -Text "The user should use the new user name: $bhrWorkEmail" -Wrap     
                                New-AdaptiveTextBlock -Text "The user's password has not been modified." -Wrap   
                            } -Uri $TeamsCardUri -Speak "Login changed for $bhrdisplayName"

                        }
                    }
                }

            }   

            # Create new employee account
            if ($aadUpnObjDetails.Capacity -eq 0 -and $aadEidObjDetails.Capacity -eq 0 -and ($bhrAccountEnabled -eq $true)) {
                Write-Log -Message "No AAD account exist but employee in bhr is $bhrAccountEnabled" -Severity Debug

                if ([string]::IsNullOrWhiteSpace($LicenseId) -eq $false) {

                    Get-LicenseStatus -LicenseId $LicenseId -TeamsCardUri $TeamsCardUri -NewUser
                }

                $PasswordProfile = @{
                    Password = (Get-NewPassword)
                }

                $error.clear() 
            
                if ($TestOnly.IsPresent) {
                    # Write logging here
                    Write-Log -Message "Executing New-MgUser -EmployeeId $bhremployeeNumber -Department $bhrDepartment -CompanyName $CompanyName -Surname $bhrlastName -GivenName $bhrfirstName -DisplayName $bhrdisplayName -AccountEnabled -Mail $bhrWorkEmail -OfficeLocation $bhrOfficeLocation `
                                -EmployeeHireDate $bhrHireDate -UserPrincipalName $bhrWorkEmail -PasswordProfile $PasswordProfile -JobTitle $bhrjobTitle -MailNickname ($bhrWorkEmail -replace '@', '' -replace $companyEmailDomain, '' ) -UsageLocation $UsageLocation -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged} " -Severity Test
                                
                }
                else {
                    # Create AAD account, as it doesn't have one, if user hire date is less than $DaysAhead days in the future, or is in the past
                    Write-Log -Message "$bhrWorkEmail does not have an AAD account and hire date ($bhrHireDate) is less than $DaysAhead days from now." -Severity Information
                    
                    Write-Log -Message "Executing New-MgUser -EmployeeId $bhremployeeNumber -Department $bhrDepartment -CompanyName $CompanyName -Surname $bhrlastName -GivenName $bhrfirstName -DisplayName $bhrdisplayName -AccountEnabled -Mail $bhrWorkEmail -OfficeLocation $bhrOfficeLocation `
                        -EmployeeHireDate $bhrHireDate -UserPrincipalName $bhrWorkEmail -PasswordProfile $PasswordProfile -JobTitle $bhrjobTitle -MailNickname ($bhrWorkEmail -replace '@', '' -replace $companyEmailDomain, '' ) -UsageLocation $UsageLocation -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged }" -Severity Debug
            
                    New-MgUser -EmployeeId $bhrEmployeeNumber -Department $bhrDepartment -CompanyName $CompanyName -Surname $bhrlastName -GivenName $bhrfirstName -DisplayName $bhrdisplayName `
                        -AccountEnabled -Mail $bhrWorkEmail -OfficeLocation $bhrOfficeLocation -EmployeeHireDate $bhrHireDate -UserPrincipalName $bhrWorkEmail -PasswordProfile $PasswordProfile `
                        -JobTitle $bhrjobTitle -MailNickname ($bhrWorkEmail -replace '@', '' -replace $companyEmailDomain, '') `
                        -UsageLocation $UsageLocation -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged }
                    
                    # Since we are setting up a new account lets use the image from the BambooHR profile and add it to the AAD account
                    Write-Log -Message "Retrieving user photo from BambooHR..." -Severity Information
                    $bhrEmployeePhotoUri = "$($bhrRootUri)/employees/$bhrEmployeeId/photo/large"
                    $profilePicPath = Join-Path -Path $env:temp -ChildPath "bhr-$($bhrEmployeeId).jpg"
                    $aadProfilePicPath = Join-Path -Path $env:temp -ChildPath "aad-$($bhrEmployeeId).jpg"
                    Remove-Item -Path $profilePicPath -ErrorAction SilentlyContinue -Force | Out-Null
                    Write-Log -Message "Executing: Invoke-RestMethod -Uri $bhrRep -Method POST -Headers $headers -ContentType 'application/json' -OutFile $profilePicPath" -Severity Debug
                    $null = Invoke-RestMethod -Uri $bhrEmployeePhotoUri -Method GET -Headers $headers -ContentType 'application/json' -OutFile $profilePicPath -ErrorAction SilentlyContinue | Out-Null

                    Write-Log "Reconnecting to Microsoft Graph..." -Severity Debug
                    $null = Disconnect-Mggraph | Out-Null
                    Connect-MgGraph -Identity -NoWelcome
                    #Connect-MgGraph -TenantId $TenantID -CertificateThumbprint $AADCertificateThumbprint -ClientId $AzureClientAppId | Out-Null
                    Write-Log "Updating user account with BambooHR profile picture..." -Severity Information
                    $user = Get-MgUser -UserId $bhrWorkEmail -ErrorAction SilentlyContinue
                    Start-Sleep 120
                    if ((Test-Path $profilePicPath -PathType Leaf -ErrorAction SilentlyContinue) -eq $false -and (Test-Path $DefaultProfilePicPath)) { 
                        $profilePicPath = $DefaultProfilePicPath
                    }

                    if (Test-Path $profilePicPath -PathType Leaf -ErrorAction SilentlyContinue) {
                        Write-Log "Executing: Set-MgUserPhotoContent -UserId $($user.Id) -InFile $profilePicPath" -Severity Debug
                        Get-MgUserPhotoContent -UserId $user.Id -OutFile $aadProfilePicPath -ErrorAction SilentlyContinue
                        Set-MgUserPhotoContent -UserId $user.Id -InFile $profilePicPath -ErrorAction Continue
                    
                    }
                    
                    if ($profilePicPath -ne $DefaultProfilePicPath) {
                        Write-Log "Executing: Remove-Item -Path $profilePicPath -ErrorAction SilentlyContinue | Out-Null" -Severity Debug
                        #Remove-Item -Path $profilePicPath -ErrorAction SilentlyContinue -Force | Out-Null
                    }

                    if ($? -eq $true) {
                        
                        if ([string]::IsNullOrWhiteSpace($bhrSupervisorEmail) -eq $false) {
                            Write-Log -Message "Account $bhrWorkEmail successfully created." -Severity Information
                            $aadmanagerID = (Get-MgUser -UserId $bhrSupervisorEmail | Select-Object id).id
                        
                            $newManager = @{
                                "@odata.id" = "https://graph.microsoft.com/v1.0/users/$aadmanagerID"
                            }
                            Start-Sleep -Seconds 8
                        
                            Write-Log -Message "Setting manager for newly created user $bhrWorkEmail." -Severity Debug
                            Write-Log -Message "Executing: Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $NewManager" -Severity Debug
                            Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $NewManager
                            $params = @{
                                Message         = @{
                                    Subject       = "User account created for: $bhrdisplayName"
                                    Body          = @{
                                        ContentType = "html"
                                        Content     = "<br/><br/><p>A new user account was created for $bhrDisplayName with hire date of $bhrHireDate. </p><p> $WelcomeUserText <ul><li>User name: $bhrWorkEmail</li><li>Password: $($PasswordProfile.Values)</li></ul><br/><p>$EmailSignature</p>"
                                    }
                                    ToRecipients  = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $bhrSupervisorEmail
                                            }
                                        }
                                    )
                                    BCCRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $NotificationEmailAddress
                                            }
                                        }
                                    )
                                }
                                SaveToSentItems = "True"
                            }
                            Write-Log -Message "Sending $bhrSupervisorEmail new employee information for $bhrDisplayName in email." -Severity Information
                            Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose

                            New-AdaptiveCard {  

                                New-AdaptiveTextBlock -Text 'New user account created' -HorizontalAlignment Center -Weight Bolder -Wrap
                                New-AdaptiveTextBlock -Text "User name: $bhrWorkEmail" -Wrap
                                #New-AdaptiveTextBlock -Text "Password: $($PasswordProfile.Values)" -Wrap
                            } -Uri $TeamsCardUri -Speak "New User $bhrDisplayName account created"

                            # Todo input these and an array and loop through only if needed.

                            # Give a little time for the mailbox to be setup so that it can receive the message.
                            Write-Output "Waiting for mailbox setup before continuing"
                            Start-Sleep -Seconds 180
                            Write-Output "Evaluating shared mailbox permissions"
                            #Connect-ExchangeOnline -CertificateThumbprint $AADCertificateThumbprint -AppId $ExchangeClientAppId -Organization $TenantId -ShowBanner:$false | Out-Null
                            Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false | Out-Null
                            foreach ($params in $MailboxDelegationParams) {
                                Sync-GroupMailboxDelegation @params -DoNotConnect
                            }

                            $newUserWelcomeEmailParams = @{
                                Message         = @{
                                    Subject       = "Welcome, $bhrFirstName!"
                                    Body          = @{
                                        ContentType = "html"
                                        Content     = "<br/><br/><p>Welcome to $CompanyName, $bhrFirstName!</p><br/>`
                                        <p> $WelcomeUserText</p><br/>`
                                        <p> Your manager will provide more details about working with your team.</p>`
                                        <p>Additionally, below you will find some helpful links to get you started.</p>`
                                        <ul>`
                                        <li><a href='https://support.microsoft.com/en-us/office/manage-meetings-ba44d0fd-da3c-4541-a3eb-a868f5e2b137'>Managing Teams Meetings</a></li>`
                                        <li><a href='https://passwordreset.microsoftonline.com/'>Password reset without a computer</a></li>`
                                        <li><a href='https://go.namedpipes.net/MfaReset'>Request Multifactor authentication reset</a></li>`
                                        <li><a href='https://outlook.office.com'>Accessing Outlook (Email) via Browser</a></li>`
                                        <li><a href='https://teams.microsoft.com'>Accessing Teams via Browser</a></li>`
                                        </ul><br/><p>$EmailSignature</p>"
                                    }
                                    ToRecipients  = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $bhrWorkEmail
                                            }
                                        }
                                    )
                                    BCCRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $bhrSupervisorEmail
                                            }
                                        }
                                    )
                                }
                                SaveToSentItems = "True"
                            }

                            Write-Output "Sending welcome email to $bhrWorkEmail"
                            Send-MgUserMail -BodyParameter $newUserWelcomeEmailParams -UserId $AdminEmailAddress -Verbose
                            
                        }
                        else {
                            $params = @{
                                Message         = @{
                                    Subject      = "User creation automation: $bhrdisplayName"
                                    Body         = @{
                                        ContentType = "html"
                                        Content     = "<br/><p>New employee user account created for $bhrDisplayName. No manager account is currently active for this account so this info is being sent to the default location.`
                                        <p> $WelcomeUserText <ul><li>User name: $bhrWorkEmail</li><li>Password: $($PasswordProfile.Values)</li></ul></p><p>$EmailSignature</p>"
                                    }
                                    ToRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $HelpDeskEmailAddress
                                            }
                                        }
                                    )
                                    CCRecipients = @(
                                        @{
                                            EmailAddress = @{
                                                Address = $NotificationEmailAddress
                                            }
                                        }
                                    )
                                }
                                SaveToSentItems = "True"
                            }
                            Write-Log -Message "Sending new employee information to default notification email because no manager was defined." -Severity Information
                            Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose

                            New-AdaptiveCard {  

                                New-AdaptiveTextBlock -Text 'New user account created without assigned manager' -HorizontalAlignment Center -Weight Bolder -Wrap
                                New-AdaptiveTextBlock -Text "No manager account is currently active for this account so this info is being sent to the default location." -Wrap
                                New-AdaptiveTextBlock -Text "User name: $bhrWorkEmail" -Wrap
                                New-AdaptiveTextBlock -Text "Password: $($PasswordProfile.Values)" -Wrap
                            } -Uri $TeamsCardUri -Speak "New User $bhrDisplayName Account Created"

                        }

                    }
                    else {
  
                        Write-Log -Message "Account $bhrWorkEmail creation failed. New-Mguser cmdlet returned error. `n $($error | Select-Object *)"-Severity Error
                        $params = @{
                            Message         = @{
                                Subject      = "BhrAadSync error: User creation automation $bhrdisplayName"
                                Body         = @{
                                    ContentType = "html"
                                    Content     = "<p>Hello,</p><br/><p>Account creation for user: $bhrWorkEmail has failed. Please check the log: $logFileName for further details.`
                                         The error information is  below. <ul><li>Error Message: $($error.Exception.Message)</li><li>Error Category: $($error.CategoryInfo)</li><li>Error ID: $($error.FullyQualifiedErrorId)</li><li>Stack: $($error.ScriptStackTrace)</li></ul></p><p>$EmailSignature</p>"
                                }
                                ToRecipients = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $HelpDeskEmailAddress
                                        }
                                    }
                                )
                                CCRecipients = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $NotificationEmailAddress
                                        }
                                    }
                                )
                            }
                            SaveToSentItems = "True"
                        }
            
                        # Send Mail Message parameters definition closure
                        Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose
                        New-AdaptiveCard {  

                            New-AdaptiveTextBlock -Text "Account creation for user: $bhrWorkEmail failed." -HorizontalAlignment Center -Weight Bolder -Wrap
                            New-AdaptiveTextBlock -Text "Error Message: $($error.Exception.Message)" -Wrap
                            New-AdaptiveTextBlock -Text "Error Category: $($error.CategoryInfo)" -Wrap
                            New-AdaptiveTextBlock -Text "Error ID: $($error.FullyQualifiedErrorId)" -Wrap
                            New-AdaptiveTextBlock -Text "Stack: $($error.ScriptStackTrace)" -Wrap
                        } -Uri $TeamsCardUri -Speak "BHR-Sync Account Creation Error"

                    }
                }
            }                    
        }
        else {
            # If Hire Date is less than $days days in the future or in the past closure
            # The user account does not need to be created as it does not satisfy the condition of the HireDate being $DaysAhead days or less in the future
            if ($bhrAccountEnabled) {
                Write-Log -Message "$bhrWorkEmail's hire date ($bhrHireDate) is more than $DaysAhead days from now." -Severity Information
            }
            else {
                Write-Log -Message "$bhrWorkEmail has been terminated, the account will not be created." -Severity Debug
            }
            
        }
    
    }
}
if (($TestOnly.IsPresent -eq $false ) -and ([string]::IsNullOrWhiteSpace($Script:logContent)) -eq $false) {
    
    if ([string]::IsNullOrWhiteSpace($LicenseId) -eq $false) {

        Get-LicenseStatus -LicenseId $LicenseId -TeamsCardUri $TeamsCardUri
    }

    Write-Log -Message "`n Completed sync at $(Get-Date) and ran for $($runtime.Totalseconds) seconds" -Severity Information
    
    New-AdaptiveCard {  
        New-AdaptiveTextBlock -Text "BHR Sync Successful" -Wrap -Weight Bolder
        $Script:logContent | ForEach-Object { $atb = $_.Replace("<p>", "").Replace("</p>", "").Replace("<br/>", "").Replace("<br>", ""); New-AdaptiveTextBlock -Text $atb -Wrap }
    } -Uri $TeamsCardUri -Speak "BambooHR to AAD sync ran successfully!"

    #Send-TeamsCard -Message "<p>BambooHR to AAD sync ran with the following results:</p><br/>$Script:LogContent<br/><br/>"
    Start-Sleep 30
    # Todo input these and an array and loop through only if needed.

    Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false | Out-Null
    foreach ($params in $MailboxDelegationParams) {
        Sync-GroupMailboxDelegation @params -DoNotConnect
 
    }

}
else { 

    Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false | Out-Null
    foreach ($params in $MailboxDelegationParams) {
        Sync-GroupMailboxDelegation @params -DoNotConnect
    }

    if ([string]::IsNullOrWhiteSpace($LicenseId) -eq $false) {

        Get-LicenseStatus -LicenseId $LicenseId -TeamsCardUri $TeamsCardUri
    }
    
    Write-Log -Message "`n Completed sync at $(Get-Date) and ran for $($runtime.Totalseconds) seconds" -Severity Information

    New-AdaptiveCard {  
        New-AdaptiveTextBlock -Text "BHR Sync Successful" -Wrap -Weight Bolder 
        $Script:logContent | ForEach-Object { 
            $atb = $_.Replace("<p>", "").Replace("</p>", "").Replace("<br/>", "").Replace("<br>", "");
            New-AdaptiveTextBlock -Text $atb -Wrap
        }
    } -Uri $TeamsCardUri -Speak "BambooHR to AAD sync ran successfully!"

    Write-Log "No log content to share, no message sent" -Severity Debug

    #Send-TeamsCard -Message "<p>BambooHR to AAD sync ran with the following results:</p><br/>$Script:LogContent<br/><br/>"
    if ($ForceSharedMailboxPermissions.IsPresent) {    
        Connect-ExchangeOnline -ManagedIdentity -Organization $TenantId -ShowBanner:$false | Out-Null
        foreach ($params in $MailboxDelegationParams) {
            Sync-GroupMailboxDelegation @params -DoNotConnect
        }
    }
}

#Script End
exit 0