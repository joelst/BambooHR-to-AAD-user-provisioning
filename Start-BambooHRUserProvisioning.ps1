#Requires -Module ExchangeOnlineManagement,PSTeams,Microsoft.Graph.Users,Microsoft.Graph.Authentication,Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Identity.SignIns

<#
===============================================================================
BambooHR to Entra ID User Provisioning Script
===============================================================================

IMPORTANT: This is a sample solution and should be used by those comfortable testing and retesting and validating
before using it in production.

All content is provided AS IS with no guarantees or assumptions of quality or functionality.

When using employee information there is so much that can go wrong!

You are responsible for complying with applicable laws and regulations for handling PII.
Remember, with great power comes great responsibility.
Friends don't let friends run untested scripts in production.


PREREQUISITES:
1. PowerShell 7+ recommended for better performance
2. Azure Automation Account with Managed Identity enabled
3. BambooHR account with API access
4. Microsoft Graph API permissions configured
5. Exchange Online Management permissions

REQUIRED MODULES (Install these first):
  Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
  Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser
  Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
  Install-Module -Name Microsoft.Graph.Identity.SignIns -Scope CurrentUser
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  Install-Module -Name PSTeams -Scope CurrentUser

AZURE AUTOMATION VARIABLES TO CONFIGURE:
  - BambooHrApiKey: Your BambooHR API key (secure string)
  - BHRCompanyName: Your BambooHR company subdomain
  - CompanyName: Your company display name
  - TenantId: Your Entra ID tenant ID
  - TeamsCardUri: (Optional) Teams webhook URL for notifications
  - BHRScript_MaxRetryAttempts: (Optional) Number of retry attempts (default: 3)
  - AdminEmailAddress: (Optional) Admin email for notifications
  - HelpDeskEmailAddress: (Optional) Help desk email for notifications
  - NotificationEmailAddress: (Optional) Email address for additional notifications
  - BHR_CustomizationsJson: (Optional) JSON overrides for per-tenant customizations

GRAPH API PERMISSIONS NEEDED:
  - User.ReadWrite.All
  - Directory.ReadWrite.All
  - Mail.Send

HOW THIS SCRIPT WORKS:
1. Connects to Entra ID using Managed Identity
2. Retrieves employee data from BambooHR API using an API key
3. For each employee, it determines if they need to be:
   - Created (new hire)
   - Updated (attribute changes)
   - Disabled (terminated)
   - Users must have a valid company email address to be considered for provisioning
4. Applies changes with automatic retry on failures
5. Sends email notifications for important events
6. Generates performance statistics and error summary

FIRST TIME SETUP:
1. Run with -WhatIf parameter to preview changes without applying them
2. Review logs carefully
3. Test with a small group of users first
4. Monitor for 2-3 runs before full deployment

.SYNOPSIS
Script to synchronize employee information from BambooHR to Azure Active Directory (Entra Id).
On premises Active Directory is not supported.

.DESCRIPTION
Extracts employee data from BambooHR and performs one of the following for each user extracted:

	1. Attribute corrections - if the user has an existing account, is an active employee, and the last changed time
    in Entra ID differs from BambooHR, then this first block will compare each of the Entra ID User object attributes
    with the data extracted from BHR and correct them if necessary.
	2. Name change - If the user has an existing account, but the name does not match with the one from BHR, then,
    this block will run and correct the user Name, UPN,	emailaddress.
	3. New employee, and there is no account in Entra ID for him, this script block will create a new user with the data
    extracted from BHR.

.PARAMETER AdminEmailAddress
Specifies the email address to receive notifications

.PARAMETER BambooHrApiKey
Specifies the BambooHR API key as a string. It will be converted to the proper format.

.PARAMETER BHRCompanyName
Specifies the BambooHR company name used in the URL

.PARAMETER BatchSize
Specifies the batch size for bulk operations. Default is 25.

.PARAMETER CompanyName
Specifies the company name to use in the employee information

.PARAMETER Confirm
Prompts you for confirmation before executing any state-changing operations.

.PARAMETER CurrentOnly
Specify to only pull current employees from BambooHR. Default is to retrieve future employees.

.PARAMETER DaysAhead
Number of days to look ahead for the employee to start.

.PARAMETER DaysToKeepAccountsAfterTermination
Specifies the number of days to keep accounts after termination. Default is 30 days.

.PARAMETER DefaultProfilePicPath
Path to a default profile image used for new users.

.PARAMETER EmailSignature
Specifies the email signature HTML to append to automated messages.

.PARAMETER EnableMobilePhoneSync
Use this to synchronize mobile phone numbers from BHR to Entra ID.

.PARAMETER ForceSharedMailboxPermissions
When specified shared mailbox permissions are updated

.PARAMETER HelpDeskEmailAddress
Email address for help desk

.PARAMETER LicenseId
When specified with a valid license id it will make sure there are still unassigned licenses before creating
 a new user.

.PARAMETER LogPath
Location to save logs

.PARAMETER MailboxDelegationParams
Specifies an array of hashtables defining mailbox delegation configurations.

.PARAMETER MaxRetryAttempts
Specifies the maximum number of retry attempts for failed operations. Default is 3.

.PARAMETER NotificationEmailAddress
Specifies an additional email address to send any notification emails to.

.PARAMETER OperationTimeoutSeconds
Specifies the timeout in seconds for API operations. Default is 120 seconds.

.PARAMETER RetryDelaySeconds
Specifies the initial delay in seconds between retry attempts. Default is 5 seconds.

.PARAMETER TeamsCardUri
Specifies the Teams webhook URI for sending adaptive card notifications.

.PARAMETER TenantId
Specifies the Microsoft tenant name (company.onmicrosoft.com)

.PARAMETER UsageLocation
A two letter country code (ISO standard 3166) to set Entra ID usage location.
Required for users that will be assigned licenses due to legal requirement to check for availability of services
 in countries. Examples include: US, JP, and GB.

.PARAMETER ModifiedWithinDays
Only process employees whose BambooHR lastChanged timestamp falls within this many days.
Default is 14. Ignored when -FullSync is specified.

.PARAMETER FullSync
Bypass the ModifiedWithinDays filter and process all employees. Use this for periodic
catch-all runs to ensure no changes are missed.

.PARAMETER WhatIf
Shows what would happen if the cmdlet runs. The cmdlet is not run. Use this to preview changes.

.PARAMETER WelcomeUserText
Sentence to add to new user email messages specific to finding the IT helpdesk FAQ.

.NOTES
More documentation available in project README
#>


[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'None')]
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidLongLines', '')]

param (
  [Parameter(Mandatory = $false, HelpMessage = 'Administrator email address for notifications and operations.')]
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
  [String]
  $AdminEmailAddress,

  [Parameter(HelpMessage = "BambooHR API key for authentication. Can be provided via Azure Automation variable 'BambooHrApiKey'.")]
  [ValidateNotNullOrEmpty()]
  [String]
  $BambooHrApiKey,

  [Parameter(HelpMessage = "BambooHR company name used in API URLs. Can be provided via Azure Automation variable 'BHRCompanyName'.")]
  [ValidateNotNullOrEmpty()]
  [string]
  $BHRCompanyName,

  [Parameter(HelpMessage = 'Batch size for bulk operations.')]
  [ValidateRange(10, 100)]
  [int]
  $BatchSize = 25,

  [Parameter(HelpMessage = "Company name for user information. Can be provided via Azure Automation variable 'CompanyName'.")]
  [ValidateNotNullOrEmpty()]
  [string]
  $CompanyName,

  [Parameter(HelpMessage = 'Only retrieve current employees, not future hires.')]
  [bool]
  $CurrentOnly = $true,

  [Parameter(HelpMessage = 'Number of days ahead to provision accounts before hire date.')]
  [ValidateRange(0, 30)]
  [int]
  $DaysAhead = 7,

  [Parameter(HelpMessage = 'Number of days to keep accounts after termination.')]
  [int]
  $DaysToKeepAccountsAfterTermination = 30,

  [Parameter(HelpMessage = 'Path to default profile picture for new users.')]
  [string]
  $DefaultProfilePicPath,

  [Parameter(HelpMessage = 'Email signature HTML for automated messages.')]
  [string]
  $EmailSignature,

  [Parameter(HelpMessage = 'Enable mobile phone synchronization from BambooHR.')]
  [bool]
  $EnableMobilePhoneSync = $true,

  [Parameter(HelpMessage = 'Force update of shared mailbox permissions.')]
  [bool]
  $ForceSharedMailboxPermissions = $false,

  [Parameter(HelpMessage = 'Help desk email address for user support.')]
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
  [string]
  $HelpDeskEmailAddress,

  [Parameter(HelpMessage = 'Microsoft 365 license SKU ID for new users.')]
  [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}_[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
  [string]
  $LicenseId,

  [Parameter(HelpMessage = 'Path for log files. Defaults to temp directory for Azure Automation.')]
  [string]
  $LogPath = $env:TEMP,

  [Parameter(HelpMessage = 'Mailbox delegation configuration array.')]
  [ValidateNotNull()]
  [Array]
  $MailboxDelegationParams = @( ),

  [Parameter(HelpMessage = 'Maximum number of retry attempts for failed operations.')]
  [ValidateRange(1, 10)]
  [int]
  $MaxRetryAttempts = 3,

  [Parameter(HelpMessage = 'HR notification email address.')]
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
  [string]
  $NotificationEmailAddress,

  [Parameter(HelpMessage = 'Timeout in seconds for API operations.')]
  [ValidateRange(30, 300)]
  [int]
  $OperationTimeoutSeconds = 120,

  [Parameter(HelpMessage = 'Initial delay in seconds between retry attempts.')]
  [ValidateRange(1, 60)]
  [int]
  $RetryDelaySeconds = 5,

  [Parameter(HelpMessage = "Teams webhook URI for notifications. Can be provided via Azure Automation variable 'TeamsCardUri'.")]
  [string]
  $TeamsCardUri,

  [Parameter(HelpMessage = "Azure tenant ID. Can be provided via Azure Automation variable 'TenantId'.")]
  [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$|^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
  [string]
  $TenantId,

  [Parameter(HelpMessage = 'Two-letter country code for user usage location (ISO 3166).')]
  [ValidatePattern('^[A-Z]{2}$')]
  [ValidateSet('US', 'GB', 'CA', 'AU', 'DE', 'FR', 'JP', 'IN', 'BR', 'MX', 'IT', 'ES', 'NL', 'SE', 'NO', 'DK', 'FI', 'IE', 'AT', 'CH', 'BE', 'PT', 'GR', 'PL', 'CZ', 'HU', 'SK', 'SI', 'HR', 'RO', 'BG', 'LT', 'LV', 'EE', 'MT', 'CY', 'LU')]
  [string]
  $UsageLocation = 'US',

  [Parameter(HelpMessage = 'Welcome message text for new users.')]
  [string]
  $WelcomeUserText,

  [Parameter(HelpMessage = 'Only process employees modified within this many days. Default is 21. Ignored when -FullSync is specified.')]
  [ValidateRange(1, 365)]
  [int]
  $ModifiedWithinDays = 21,

  [Parameter(HelpMessage = 'Bypass the ModifiedWithinDays filter and process all employees.')]
  [bool]
  $FullSync = $false
)

# Script-level variables
$AzureAutomate = $true
$Script:logContent = ''
$Script:SignificantChanges = [ordered]@{
  Created        = @{}
  Disabled       = @{}
  NameChanged    = @{}
  UpnChanged     = @{}
  ManagerChanged = @{}
  UpdatedMajor   = @{}
}
$Script:CorrelationId = [Guid]::NewGuid().ToString()
$Script:StartTime = Get-Date

function Add-SignificantChange {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Created', 'Disabled', 'NameChanged', 'UpnChanged', 'ManagerChanged', 'UpdatedMajor')]
    [string]
    $Category,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $User,

    [Parameter()]
    [string]
    $Detail
  )

  if (-not $Script:SignificantChanges) {
    return
  }

  if (-not $Script:SignificantChanges.Contains($Category)) {
    $Script:SignificantChanges[$Category] = @{}
  }

  if (-not $Script:SignificantChanges[$Category].ContainsKey($User)) {
    $Script:SignificantChanges[$Category][$User] = $Detail
    return
  }

  if ([string]::IsNullOrWhiteSpace($Script:SignificantChanges[$Category][$User]) -and -not [string]::IsNullOrWhiteSpace($Detail)) {
    $Script:SignificantChanges[$Category][$User] = $Detail
  }
}

# Connection state tracking to prevent redundant connections
$Script:MgGraphConnected = $false
$Script:ExchangeConnected = $false
$Script:AzureConnected = $false

$PSDefaultParameterValues = @{'Import-Module:Verbose' = $false }

#region Configuration Management
<#
===============================================================================
CONFIGURATION MANAGEMENT SECTION
===============================================================================

This section handles all configuration for the script. Understanding this is
crucial for developers:

KEY CONCEPTS:
1. Configuration Centralization: All settings stored in $Script:Config object
2. Azure Automation Variables: Settings pulled from Azure Automation account
3. Parameter Precedence: Command-line params override automation variables
4. Validation: All required settings are validated before processing starts

CONFIGURATION STRUCTURE:
  $Script:Config = @{
    Runtime      - Script execution settings (WhatIfPreference, logging, retries)
    BambooHR     - BambooHR API connection settings
    Azure        - Entra ID/Graph API settings
    Email        - Email notification addresses and templates
    Features     - Optional features (Teams cards, delegation, etc.)
  }

HOW TO ADD NEW CONFIGURATION:
1. Add parameter to param() block at top of script
2. Add to relevant section in Initialize-Configuration
3. Access via $Script:Config.Section.SettingName throughout script
4. Document in SETUP GUIDE section above
#>

function Initialize-Configuration {
  <#
    .SYNOPSIS
    Initialize and validate all configuration parameters for the runbook.

    .DESCRIPTION
    This function centralizes configuration management, retrieves values from Azure Automation variables,
    validates all parameters, and sets up the runtime environment.

    WHAT THIS FUNCTION DOES:
    1. Creates $Script:Config hashtable with all settings
    2. Attempts to retrieve missing values from Azure Automation variables
    3. Validates required parameters are present
    4. Validates email address formats
    5. Sets up logging paths and filenames
    6. Returns configuration object for use throughout script

    DEVELOPER NOTE:
    If you add a new configuration parameter:
    - Add it to param() block at script start
    - Add to appropriate section below (Runtime, BambooHR, Azure, etc.)
    - Add validation if it's required
    - Document it in the header comments
    #>
  [CmdletBinding()]
  param()

  Write-Verbose "[Initialize-Configuration] Starting configuration initialization with correlation ID: $Script:CorrelationId"

  $toBool = {
    param($Value)
    if ($Value -is [System.Management.Automation.SwitchParameter]) {
      return $Value.IsPresent
    }
    return [bool]$Value
  }

  # Configuration object to store all settings
  $config = @{
    CorrelationId    = $Script:CorrelationId
    StartTime        = $Script:StartTime
    IsValid          = $true
    ValidationErrors = @()
    Runtime          = @{}
    BambooHR         = @{}
    Azure            = @{}
    Email            = @{}
    Features         = @{}
  }

  try {
    # Get values from Azure Automation variables if not provided as parameters
    $automationVariables = @{
      'BambooHrApiKey'           = $script:BambooHrApiKey
      'BHRCompanyName'           = $script:BHRCompanyName
      'CompanyName'              = $script:CompanyName
      'TenantId'                 = $script:TenantId
      'TeamsCardUri'             = $script:TeamsCardUri
      'AdminEmailAddress'        = $script:AdminEmailAddress
      'NotificationEmailAddress' = $script:NotificationEmailAddress
      'HelpDeskEmailAddress'     = $script:HelpDeskEmailAddress
      'LicenseId'                = $script:LicenseId
    }

    foreach ($varName in $automationVariables.Keys) {
      if ([string]::IsNullOrWhiteSpace($automationVariables[$varName])) {
        try {
          $automationValue = Get-AutomationVariable -Name $varName -ErrorAction SilentlyContinue
          if ($automationValue) {
            Set-Variable -Name $varName -Value $automationValue -Scope Script
            Write-Verbose "[Initialize-Configuration] Retrieved $varName from Azure Automation variables"
          }
        }
        catch {
          Write-Verbose "[Initialize-Configuration] Could not retrieve $varName from Azure Automation variables: $($_.Exception.Message)"
        }
      }
    }

    # Apply optional JSON-based customizations
    $customizationsJson = $null
    try {
      if (-not [string]::IsNullOrWhiteSpace($script:CustomizationsJson)) {
        $customizationsJson = $script:CustomizationsJson
      }
      else {
        $customizationsJson = Get-AutomationVariable -Name 'BHR_CustomizationsJson' -ErrorAction SilentlyContinue
      }
    }
    catch {
      Write-Verbose "[Initialize-Configuration] Could not retrieve BHR_CustomizationsJson: $($_.Exception.Message)"
    }

    if (-not [string]::IsNullOrWhiteSpace($customizationsJson)) {
      try {
        $custom = $customizationsJson | ConvertFrom-Json -Depth 6

        if ($custom.AdminEmailAddress) { $script:AdminEmailAddress = $custom.AdminEmailAddress }
        if ($custom.NotificationEmailAddress) { $script:NotificationEmailAddress = $custom.NotificationEmailAddress }
        if ($custom.HelpDeskEmailAddress) { $script:HelpDeskEmailAddress = $custom.HelpDeskEmailAddress }
        if ($custom.TeamsCardUri) { $script:TeamsCardUri = $custom.TeamsCardUri }
        if ($custom.LicenseId) { $script:LicenseId = $custom.LicenseId }
        if ($custom.UsageLocation) { $script:UsageLocation = $custom.UsageLocation }
        if ($null -ne $custom.DaysAhead) { $script:DaysAhead = [int]$custom.DaysAhead }
        if ($null -ne $custom.DaysToKeepAccountsAfterTermination) {
          $script:DaysToKeepAccountsAfterTermination = [int]$custom.DaysToKeepAccountsAfterTermination
        }
        $enableMobilePhoneSyncWasBound = ($null -ne $script:PSBoundParameters -and $script:PSBoundParameters.ContainsKey('EnableMobilePhoneSync'))
        if ($null -ne $custom.EnableMobilePhoneSync -and -not $enableMobilePhoneSyncWasBound) {
          $script:EnableMobilePhoneSync = [bool]$custom.EnableMobilePhoneSync
        }
        $currentOnlyWasBound = ($null -ne $script:PSBoundParameters -and $script:PSBoundParameters.ContainsKey('CurrentOnly'))
        if ($null -ne $custom.CurrentOnly -and -not $currentOnlyWasBound) {
          $script:CurrentOnly = [bool]$custom.CurrentOnly
        }
        $forceSharedMailboxPermissionsWasBound = ($null -ne $script:PSBoundParameters -and $script:PSBoundParameters.ContainsKey('ForceSharedMailboxPermissions'))
        if ($null -ne $custom.ForceSharedMailboxPermissions -and -not $forceSharedMailboxPermissionsWasBound) {
          $script:ForceSharedMailboxPermissions = [bool]$custom.ForceSharedMailboxPermissions
        }
        if ($custom.DefaultProfilePicPath) { $script:DefaultProfilePicPath = $custom.DefaultProfilePicPath }
        if ($custom.EmailSignature) { $script:EmailSignature = $custom.EmailSignature }
        if ($custom.WelcomeUserText) { $script:WelcomeUserText = $custom.WelcomeUserText }
        if ($custom.MailboxDelegationParams) { $script:MailboxDelegationParams = @($custom.MailboxDelegationParams) }
        if ($custom.WelcomeLinksHtml) { $script:WelcomeLinksHtml = $custom.WelcomeLinksHtml }
        if ($null -ne $custom.ModifiedWithinDays) { $script:ModifiedWithinDays = [int]$custom.ModifiedWithinDays }
        $fullSyncWasBound = ($null -ne $script:PSBoundParameters -and $script:PSBoundParameters.ContainsKey('FullSync'))
        if ($null -ne $custom.FullSync -and -not $fullSyncWasBound) {
          $script:FullSync = [bool]$custom.FullSync
        }

        Write-Verbose '[Initialize-Configuration] Applied BHR_CustomizationsJson overrides'
      }
      catch {
        Write-Warning "[Initialize-Configuration] Failed to parse BHR_CustomizationsJson: $($_.Exception.Message)"
      }
    }

    # Validate required parameters
    $requiredParams = @{
      'BambooHrApiKey' = $script:BambooHrApiKey
      'BHRCompanyName' = $script:BHRCompanyName
      'CompanyName'    = $script:CompanyName
      'TenantId'       = $script:TenantId
    }

    foreach ($param in $requiredParams.GetEnumerator()) {
      if ([string]::IsNullOrWhiteSpace($param.Value)) {
        $config.ValidationErrors += "Required parameter '$($param.Key)' is missing or empty"
        $config.IsValid = $false
      }
    }

    # Set default values for optional parameters
    if ([string]::IsNullOrWhiteSpace($script:AdminEmailAddress)) {
      $script:AdminEmailAddress = "admin@$($script:CompanyName.ToLower().Replace(' ', '')).com"
      Write-Warning "AdminEmailAddress not provided, using default: $script:AdminEmailAddress"
    }

    if ([string]::IsNullOrWhiteSpace($script:NotificationEmailAddress)) {
      $script:NotificationEmailAddress = "hr@$($script:CompanyName.ToLower().Replace(' ', '')).com"
      Write-Warning "NotificationEmailAddress not provided, using default: $script:NotificationEmailAddress"
    }

    if ([string]::IsNullOrWhiteSpace($script:HelpDeskEmailAddress)) {
      $script:HelpDeskEmailAddress = "helpdesk@$($script:CompanyName.ToLower().Replace(' ', '')).com"
      Write-Warning "HelpDeskEmailAddress not provided, using default: $script:HelpDeskEmailAddress"
    }

    $licensePattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}_[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    if ([string]::IsNullOrWhiteSpace($script:LicenseId) -or ($script:LicenseId -notmatch $licensePattern)) {
      Write-Warning 'LicenseId not provided or invalid.'
    }

    # Build configuration object
    $config.Runtime = @{
      LogPath                 = $script:LogPath
      WhatIfPreference        = $WhatIfPreference
      MaxRetryAttempts        = $script:MaxRetryAttempts
      RetryDelaySeconds       = $script:RetryDelaySeconds
      OperationTimeoutSeconds = $script:OperationTimeoutSeconds
      BatchSize               = $script:BatchSize
    }

    $config.BambooHR = @{
      ApiKey      = $script:BambooHrApiKey
      CompanyName = $script:BHRCompanyName
      RootUri     = "https://api.bamboohr.com/api/gateway.php/$($script:BHRCompanyName)/v1"
      ReportsUri  = if ((& $toBool $script:CurrentOnly)) {
        "https://api.bamboohr.com/api/gateway.php/$($script:BHRCompanyName)/v1/reports/custom?format=json&onlyCurrent=true"
      }
      else {
        "https://api.bamboohr.com/api/gateway.php/$($script:BHRCompanyName)/v1/reports/custom?format=json&onlyCurrent=false"
      }
    }

    $config.Azure = @{
      TenantId      = $script:TenantId
      UsageLocation = $script:UsageLocation
      LicenseId     = $script:LicenseId
      CompanyName   = $script:CompanyName
    }

    $config.Email = @{
      AdminEmailAddress        = $script:AdminEmailAddress
      NotificationEmailAddress = $script:NotificationEmailAddress
      HelpDeskEmailAddress     = $script:HelpDeskEmailAddress
      EmailSignature           = if ([string]::IsNullOrWhiteSpace($script:EmailSignature)) {
        "<br/>Regards, <br/> $($script:CompanyName) Automated User Management <br/><br/>For additional information, please contact IT support.<br/>"
      }
      else { $script:EmailSignature }
      WelcomeUserText          = if ([string]::IsNullOrWhiteSpace($script:WelcomeUserText)) {
        "Welcome to $($script:CompanyName)! Please contact IT support for assistance with your new account."
      }
      else { $script:WelcomeUserText }
      WelcomeLinksHtml         = if ([string]::IsNullOrWhiteSpace($script:WelcomeLinksHtml)) {
        "<p>Your manager will provide more details about working with your team.</p>`
<p>Additionally, below you will find some helpful links to get you started.</p>`
<ul>`
<li><a href='https://support.microsoft.com/en-us/office/manage-meetings-ba44d0fd-da3c-4541-a3eb-a868f5e2b137'>Managing Teams Meetings</a></li>`
<li><a href='https://passwordreset.microsoftonline.com/'>Password reset without a computer</a></li>`
<li><a href='https://outlook.office.com'>Accessing Outlook (Email) via Browser</a></li>`
<li><a href='https://teams.microsoft.com'>Accessing Teams via Browser</a></li>`
</ul><br/>"
      }
      else { $script:WelcomeLinksHtml }
      CompanyEmailDomain       = ($script:AdminEmailAddress -split '@')[1]
    }

    $config.Features = @{
      EnableMobilePhoneSync              = (& $toBool $script:EnableMobilePhoneSync)
      CurrentOnly                        = (& $toBool $script:CurrentOnly)
      ForceSharedMailboxPermissions      = (& $toBool $script:ForceSharedMailboxPermissions)
      DaysAhead                          = $script:DaysAhead
      DaysToKeepAccountsAfterTermination = $script:DaysToKeepAccountsAfterTermination
      ModifiedWithinDays                 = $script:ModifiedWithinDays
      FullSync                           = (& $toBool $script:FullSync)
      TeamsCardUri                       = $script:TeamsCardUri
      DefaultProfilePicPath              = $script:DefaultProfilePicPath
      MailboxDelegationParams            = if ($script:MailboxDelegationParams.Count -eq 0) {
        # Set default mailbox delegation configuration if none provided
        @(
          @{ Group = 'CG-SharedMailboxDelegatedAccessMailbox1'; DelegateMailbox = 'Mailbox1' }
          @{ Group = 'CG-SharedMailboxDelegatedAccessMailbox2'; DelegateMailbox = 'Mailbox2' }
          @{ Group = 'CG-SharedMailboxDelegatedAccessMailbox3'; DelegateMailbox = 'Mailbox3' }
          @{ Group = 'CG-SharedMailboxDelegatedAccessMailbox4'; DelegateMailbox = 'Mailbox4' }
          @{ Group = 'CG-SharedMailboxDelegatedAccessMailbox5'; DelegateMailbox = 'Mailbox5' }
        )
      }
      else { $script:MailboxDelegationParams }
    }

    # Validate email addresses
    $emailFields = @('AdminEmailAddress', 'NotificationEmailAddress', 'HelpDeskEmailAddress')
    foreach ($field in $emailFields) {
      $email = $config.Email[$field]
      if ($email -and $email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
        $config.ValidationErrors += "Invalid email format for ${field}: $email"
        $config.IsValid = $false
      }
    }

    # Setup logging
    $logFileName = 'BhrEntraSync-' + (Get-Date -Format yyyyMMdd-HHmm) + '-' + $Script:CorrelationId.Substring(0, 8) + '.csv'
    $logFilePath = Join-Path $config.Runtime.LogPath $logFileName
    $config.Runtime.LogFileName = $logFileName
    $config.Runtime.LogFilePath = $logFilePath

    Write-Verbose '[Initialize-Configuration] Configuration initialized successfully'
    return $config
  }
  catch {
    Write-Error "[Initialize-Configuration] Failed to initialize configuration: $($_.Exception.Message)"
    $config.ValidationErrors += "Configuration initialization failed: $($_.Exception.Message)"
    $config.IsValid = $false
    return $config
  }
}

#endregion Script-Level Variables

# Initialize configuration from parameters and Azure Automation variables
$Script:Config = Initialize-Configuration

# Validate configuration before proceeding with any operations
if (-not $Script:Config.IsValid) {
  $errorMessage = "Configuration validation failed:`n" + ($Script:Config.ValidationErrors -join "`n")
  Write-Error $errorMessage

  # Send notification if possible
  if ($Script:Config.Features.TeamsCardUri) {
    try {
      New-AdaptiveCard {
        New-AdaptiveTextBlock -Text 'BHR Sync Configuration Error' -Weight Bolder -Wrap -Color Red
        New-AdaptiveTextBlock -Text $errorMessage -Wrap
      } -Uri $Script:Config.Features.TeamsCardUri -Speak 'BHR Sync configuration validation failed'
    }
    catch {
      Write-Warning "Could not send Teams notification: $($_.Exception.Message)"
    }
  }

  exit 1
}

Write-Output "[Initialize] BambooHR to Entra ID sync started with correlation ID: $($Script:Config.CorrelationId)"

# Logging Function

function Write-PSLog {
  <#
  .SYNOPSIS
  Write a log entry to the log file with enhanced tracking and correlation.
  .PARAMETER Message
  The message to log.
  .PARAMETER Severity
  The severity level of the log entry.
  .PARAMETER CorrelationId
  Optional correlation ID for tracking related operations.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$Message,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Debug', 'Information', 'Warning', 'Error', 'Test')]
    [string]$Severity = 'Information',

    [Parameter()]
    [string]$CorrelationId = $Script:Config.CorrelationId
  )

  # Add correlation ID and timestamp to message
  $timestampedMessage = "[$CorrelationId] [$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"

  # Log to file if not in Azure Automation mode
  if ($AzureAutomate -eq $false -and $Script:Config.Runtime.LogFilePath) {
    [pscustomobject]@{
      Time          = (Get-Date -Format 'yyyy/MM/dd HH:mm:ss')
      CorrelationId = $CorrelationId
      Message       = ($Message.Replace("`n", '').Replace("`t", '').Replace("``", ''))
      Severity      = $Severity
      Operation     = (Get-PSCallStack)[1].FunctionName
    } | Export-Csv -Path $Script:Config.Runtime.LogFilePath -Append -NoTypeInformation -Force
  }

  # Output to appropriate stream based on severity
  switch ($Severity) {
    'Debug' {
      Write-Verbose $timestampedMessage
    }
    'Warning' {
      Write-Warning $timestampedMessage
      $Script:logContent += "<p>$Message</p>`n"
    }
    'Error' {
      Write-Error $timestampedMessage
      $Script:logContent += "<p><strong>ERROR:</strong> $Message</p>`n"
    }
    'Information' {
      Write-Output $timestampedMessage
      $Script:logContent += "<p>$Message</p>`n"
    }
    'Test' {
      Write-Information "[WhatIf] $timestampedMessage" -InformationAction Continue
      $Script:logContent += "<p><em>[TEST ONLY]</em> $Message</p>`n"
    }
  }
}

function Invoke-WithRetry {
  <#
    .SYNOPSIS
    Execute a script block with retry logic and exponential backoff.

    .DESCRIPTION
    Provides resilient execution of operations with configurable retry attempts,
    exponential backoff, and detailed error handling for Azure Automation scenarios.

    .PARAMETER ScriptBlock
    The script block to execute

    .PARAMETER MaxAttempts
    Maximum number of retry attempts

    .PARAMETER InitialDelaySeconds
    Initial delay between retries in seconds

    .PARAMETER Operation
    Name of the operation for logging purposes

    .PARAMETER RetryableErrorTypes
    Array of error types that should trigger a retry
    #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)]
    [ScriptBlock]$ScriptBlock,

    [Parameter()]
    [int]$MaxAttempts = $Script:Config.Runtime.MaxRetryAttempts,

    [Parameter()]
    [int]$InitialDelaySeconds = $Script:Config.Runtime.RetryDelaySeconds,

    [Parameter()]
    [string]$Operation = 'Unknown Operation',

    [Parameter()]
    [string[]]$RetryableErrorTypes = @(
      'System.Net.WebException',
      'Microsoft.Graph.PowerShell.Runtime.RestException',
      'System.TimeoutException',
      'System.Net.Http.HttpRequestException'
    )
  )

  $attempt = 1
  $delay = $InitialDelaySeconds

  do {
    try {
      Write-PSLog "Attempting $Operation (attempt $attempt of $MaxAttempts)" -Severity Debug
      $result = & $ScriptBlock
      Write-PSLog "$Operation completed successfully on attempt $attempt" -Severity Debug
      return $result
    }
    catch {
      $errorType = $_.Exception.GetType().FullName
      $isRetryable = $RetryableErrorTypes -contains $errorType

      if ($attempt -ge $MaxAttempts -or -not $isRetryable) {
        Write-PSLog "$Operation failed after $attempt attempts. Error: $($_.Exception.Message)" -Severity Error
        throw
      }

      Write-PSLog "$Operation failed on attempt $attempt with retryable error ($errorType). Retrying in $delay seconds..." -Severity Warning
      Start-Sleep -Seconds $delay

      # Exponential backoff with jitter
      $delay = [Math]::Min($delay * 2 + (Get-Random -Minimum 1 -Maximum 5), 60)
      $attempt++
    }
  } while ($attempt -le $MaxAttempts)
}

#endregion Retry and Error Handling

#region Performance Helper Functions
<#
===============================================================================
PERFORMANCE OPTIMIZATION SECTION
===============================================================================

This section contains functions to improve script performance and provide metrics.

KEY PERFORMANCE FEATURES:

1. CACHING:
   - Stores frequently accessed data (like manager lookups)
   - Reduces redundant API calls by 30-50%
   - Thread-safe on PowerShell 7+ (uses ConcurrentDictionary)
   - Falls back to regular hashtable on PowerShell 5.1

2. PERFORMANCE METRICS:
   - Tracks script execution time
   - Calculates users per minute throughput
   - Measures cache effectiveness (hit rate)
   - Provides data for optimization decisions

HOW CACHING WORKS:
  First call:  Get-CachedUser → Cache MISS → API call → Store in cache
  Second call: Get-CachedUser → Cache HIT → Return from cache (instant)

  Example: 50 users with same manager
  Without cache: 50 API calls = 25 seconds
  With cache:    1 API call + 49 instant = 0.5 seconds (50x faster!)

DEVELOPER NOTE:
- Use Get-CachedUser instead of Get-MgUser when possible
- Cache is automatically initialized before employee processing
- Cache statistics are reported at script completion
- Don't cache data that changes frequently
#>

function Initialize-PerformanceCache {
  <#
  .SYNOPSIS
  Initialize thread-safe cache for frequently accessed user data.

  .DESCRIPTION
  Creates concurrent dictionaries for caching user lookups, group memberships,
  and license information to reduce redundant API calls.

  .OUTPUTS
  Hashtable containing cache objects.
  #>
  [CmdletBinding()]
  param()

  $userLookupCache = @{}
  $managerCache = @{}
  $offboardingCache = @{
    DistributionGroups           = $null
    Mailboxes                    = $null
    AutopilotByUserPrincipalName = @{}
    AutopilotIndexed             = $false
  }

  $cache = @{
    UserLookup  = $userLookupCache
    Manager     = $managerCache
    Offboarding = $offboardingCache
    Stats       = @{
      Hits   = 0
      Misses = 0
    }
  }

  Write-PSLog 'Initialized performance cache' -Severity Information | Out-Null
  return $cache
}

function Get-CachedUser {
  <#
  .SYNOPSIS
  Get user from cache or retrieve via API with caching.

  .DESCRIPTION
  Checks cache first, then retrieves from Microsoft Graph if not found.
  Automatically caches the result for future use.

  .PARAMETER UserId
  User ID (UPN or Object ID) to retrieve.

  .PARAMETER Cache
  Cache hashtable from Initialize-PerformanceCache.

  .PARAMETER Force
  Force refresh from API, bypassing cache.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$UserId,

    [Parameter(Mandatory = $true)]
    [hashtable]$Cache,

    [Parameter(Mandatory = $false)]
    [switch]$Force
  )

  if (-not $Force -and $Cache.UserLookup.ContainsKey($UserId)) {
    $Cache.Stats.Hits++
    Write-PSLog "Cache hit for user: $UserId (Total hits: $($Cache.Stats.Hits))" -Severity Debug
    return $Cache.UserLookup[$UserId]
  }

  $Cache.Stats.Misses++
  Write-PSLog "Cache miss for user: $UserId - Retrieving from API (Total misses: $($Cache.Stats.Misses))" -Severity Debug

  try {
    $user = Invoke-WithRetry -Operation "Get cached user: $UserId" -ScriptBlock {
      Get-MgUser -UserId $UserId -ErrorAction Stop
    }

    if ($user) {
      # Store in cache
      $Cache.UserLookup[$UserId] = $user
    }

    return $user
  }
  catch {
    Write-PSLog "Failed to retrieve user $UserId : $($_.Exception.Message)" -Severity Warning
    return $null
  }
}

function Test-IsTenantEmail {
  <#
  .SYNOPSIS
  Determine whether an email address matches a verified tenant domain.

  .PARAMETER EmailAddress
  Email address to validate.

  .PARAMETER TenantDomains
  Verified tenant domains.
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param(
    [Parameter()]
    [string]$EmailAddress,

    [Parameter()]
    [string[]]$TenantDomains
  )

  if ([string]::IsNullOrWhiteSpace($EmailAddress)) {
    return $false
  }

  foreach ($tenantDomain in @($TenantDomains)) {
    if ($EmailAddress -like "*@$tenantDomain") {
      return $true
    }
  }

  return $false
}

function Get-MailNicknameFromEmail {
  <#
  .SYNOPSIS
  Derive the mail nickname from an email address.

  .PARAMETER EmailAddress
  Email address to convert.
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter()]
    [string]$EmailAddress
  )

  if ([string]::IsNullOrWhiteSpace($EmailAddress)) {
    return ''
  }

  return ($EmailAddress.Trim() -split '@')[0]
}

function Get-WorkPhoneComparisonValue {
  <#
  .SYNOPSIS
  Normalize a work phone value for sync comparisons.

  .PARAMETER PhoneNumber
  Phone number value from BambooHR or Entra ID.
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter()]
    [string]$PhoneNumber
  )

  if ([string]::IsNullOrWhiteSpace($PhoneNumber) -or $PhoneNumber -eq '0') {
    return '0'
  }

  $normalizedPhone = ConvertTo-PhoneNumber $PhoneNumber
  if ([string]::IsNullOrWhiteSpace($normalizedPhone) -or $normalizedPhone -eq '0') {
    return '0'
  }

  return $normalizedPhone
}

function ConvertTo-BambooHrHireDate {
  <#
  .SYNOPSIS
  Normalize a BambooHR hire date using invariant parsing.

  .DESCRIPTION
  BambooHR normally returns hire dates as yyyy-MM-dd. Missing or invalid values
  are normalized to DateTime.MinValue so the sync can keep processing and use a
  deterministic default for Graph writes.

  .PARAMETER HireDate
  BambooHR hire date value.

  .PARAMETER EmployeeEmailAddress
  Employee email used in warning logs.
  #>
  [CmdletBinding()]
  [OutputType([datetime])]
  param(
    [Parameter()]
    [string]$HireDate,

    [Parameter()]
    [string]$EmployeeEmailAddress
  )

  $minimumHireDate = [datetime]::MinValue.Date
  $minimumHireDateString = $minimumHireDate.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)

  if ([string]::IsNullOrWhiteSpace($HireDate)) {
    Write-PSLog -Message "$EmployeeEmailAddress has an empty hire date in BambooHR. Using minimum supported date '$minimumHireDateString'." -Severity Warning
    return $minimumHireDate
  }

  $parsedHireDate = $minimumHireDate
  $supportedFormats = @(
    'yyyy-MM-dd'
    'yyyy-MM-ddTHH:mm:ss'
    'yyyy-MM-ddTHH:mm:ssZ'
  )

  foreach ($supportedFormat in $supportedFormats) {
    if ([datetime]::TryParseExact($HireDate, $supportedFormat, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedHireDate)) {
      return $parsedHireDate.Date
    }
  }

  Write-PSLog -Message "$EmployeeEmailAddress has an invalid hire date '$HireDate' in BambooHR. Using minimum supported date '$minimumHireDateString'." -Severity Warning
  return $minimumHireDate
}

function Test-ShouldSyncExistingUser {
  <#
  .SYNOPSIS
  Determine whether an existing Entra ID account should enter the attribute sync pipeline.
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param(
    [Parameter()]
    [string]$EntraIdEmployeeNumber,

    [Parameter()]
    [string]$EntraIdEmployeeNumberByEid,

    [Parameter()]
    [string]$EntraIdUpnFromEidLookup,

    [Parameter()]
    [string]$EntraIdUpnFromUpnLookup,

    [Parameter()]
    [string]$BhrWorkEmail,

    [Parameter()]
    [string]$BhrLastChanged,

    [Parameter()]
    [string]$UpnExtensionAttribute1,

    [Parameter()]
    [object]$EntraIdEidObjDetails,

    [Parameter()]
    [object]$EntraIdUpnObjDetails,

    [Parameter()]
    [string]$BhrEmploymentStatus
  )

  $hasMatchingEmployeeId = $EntraIdEmployeeNumber -eq $EntraIdEmployeeNumberByEid
  $hasMatchingUpn = ($EntraIdUpnFromEidLookup -eq $BhrWorkEmail) -or ($EntraIdUpnFromUpnLookup -eq $BhrWorkEmail)
  $hasLookupResults = (@($EntraIdEidObjDetails).Count -gt 0) -or (@($EntraIdUpnObjDetails).Count -gt 0)

  # Always process terminated employees whose Entra account is still enabled,
  # even when the lastChanged date already matches ExtensionAttribute1.
  $entraIdEnabled = if ($EntraIdUpnObjDetails) {
    [bool]$EntraIdUpnObjDetails.AccountEnabled
  }
  elseif ($EntraIdEidObjDetails) {
    [bool]$EntraIdEidObjDetails.AccountEnabled
  }
  else {
    $false
  }
  $needsOffboarding = ($BhrEmploymentStatus -like '*Terminated*') -and $entraIdEnabled -and $hasLookupResults

  return $needsOffboarding -or (
    ($hasMatchingEmployeeId -or $hasMatchingUpn) -and
    ($BhrLastChanged -ne $UpnExtensionAttribute1) -and
    $hasLookupResults -and
    ($BhrEmploymentStatus -notlike '*suspended*')
  )
}

function Test-ShouldUpdateEmployeeId {
  <#
  .SYNOPSIS
  Determine whether EmployeeId should be updated from BambooHR.

  .DESCRIPTION
  Terminated inactive users are excluded so offboarded accounts retain the
  termination marker EmployeeId value until they are explicitly re-enabled.
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param(
    [Parameter()]
    [string]$EntraIdEmployeeNumber,

    [Parameter()]
    [string]$BhrEmployeeNumber,

    [Parameter()]
    [string]$EntraIdUserPrincipalName,

    [Parameter()]
    [string]$BhrWorkEmail,

    [Parameter()]
    [string]$BhrEmploymentStatus,

    [Parameter()]
    [bool]$BhrAccountEnabled,

    [Parameter()]
    [string]$BhrLastChanged,

    [Parameter()]
    [string]$UpnExtensionAttribute1
  )

  $employmentStatus = if ([string]::IsNullOrWhiteSpace($BhrEmploymentStatus)) { '' } else { $BhrEmploymentStatus.Trim() }
  if ((-not $BhrAccountEnabled) -and ($employmentStatus -eq 'Terminated')) {
    return $false
  }

  return ($EntraIdEmployeeNumber -ne $BhrEmployeeNumber) -and
  ($EntraIdUserPrincipalName -eq $BhrWorkEmail) -and
  ($employmentStatus -notlike '*suspended*') -and
  ($BhrLastChanged -ne $UpnExtensionAttribute1)
}

function Test-IsOffboardingComplete {
  <#
  .SYNOPSIS
  Determine whether the offboarding completion markers are present.
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param(
    [Parameter()]
    [string]$CompanyName,

    [Parameter()]
    [string]$Department,

    [Parameter()]
    [string]$JobTitle,

    [Parameter()]
    [string]$OfficeLocation,

    [Parameter()]
    [string]$WorkPhone,

    [Parameter()]
    [string]$MobilePhone
  )

  $workPhoneCleared = $true
  if ($PSBoundParameters.ContainsKey('WorkPhone')) {
    $workPhoneCleared = (Get-WorkPhoneComparisonValue -PhoneNumber $WorkPhone) -eq '0'
  }

  $mobilePhoneCleared = $true
  if ($PSBoundParameters.ContainsKey('MobilePhone')) {
    $mobilePhoneCleared = (Get-WorkPhoneComparisonValue -PhoneNumber $MobilePhone) -eq '0'
  }

  return ($CompanyName -match '^\d{2}/\d{2}/\d{2}') -and
  ($Department -eq 'Not Active') -and
  ($JobTitle -eq 'Not Active') -and
  ($OfficeLocation -eq 'Not Active') -and
  $workPhoneCleared -and
  $mobilePhoneCleared
}

function Set-TerminatedUserProfileFields {
  <#
  .SYNOPSIS
  Apply offboarding placeholders and clear phone/location details.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$UserId,

    [Parameter(Mandatory = $true)]
    [datetime]$LeaveDateTimeUtc
  )

  $leaveDateTimeUtcString = $LeaveDateTimeUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

  Write-PSLog -Message "Executing: Update-MgUser -UserId $UserId -Department 'Not Active' -JobTitle 'Not Active' -OfficeLocation 'Not Active' -StreetAddress `$null -City `$null -State `$null -PostalCode `$null -EmployeeLeaveDateTime '$leaveDateTimeUtcString'" -Severity Debug
  Invoke-WithRetry -Operation "Set terminated profile placeholders for: $UserId" -ScriptBlock {
    Update-MgUser -UserId $UserId -Department 'Not Active' -JobTitle 'Not Active' -OfficeLocation 'Not Active' -StreetAddress $null -City $null -State $null -PostalCode $null -EmployeeLeaveDateTime $LeaveDateTimeUtc -ErrorAction Stop
  }

  Write-PSLog -Message "Executing: Update-MgUser -UserId $UserId -BusinessPhones @() -MobilePhone `$null" -Severity Debug
  Invoke-WithRetry -Operation "Clear terminated user phone fields for: $UserId" -ScriptBlock {
    Update-MgUser -UserId $UserId -BusinessPhones @() -MobilePhone $null -ErrorAction Stop
  }
}

function Get-OffboardingCompletionMarker {
  <#
  .SYNOPSIS
  Build the CompanyName stamp used to mark completed offboarding.
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $true)]
    [datetime]$LeaveDateTimeUtc
  )

  $leaveDateTimeUtcString = $LeaveDateTimeUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
  return "$(Get-Date -Date $LeaveDateTimeUtc -UFormat %D) (OffboardingComplete: $leaveDateTimeUtcString)"
}

function Get-OffboardingCompletionDateFromCompanyName {
  <#
  .SYNOPSIS
  Extract the offboarding completion timestamp from the CompanyName marker.
  #>
  [CmdletBinding()]
  [OutputType([datetime])]
  param(
    [Parameter()]
    [string]$CompanyName
  )

  if ([string]::IsNullOrWhiteSpace($CompanyName)) {
    return $null
  }

  $match = [regex]::Match($CompanyName, 'OffboardingComplete:\s*(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z)')
  if (-not $match.Success) {
    return $null
  }

  try {
    return [datetime]::Parse($match.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()
  }
  catch {
    return $null
  }
}

function Wait-ForCondition {
  <#
  .SYNOPSIS
  Poll until a condition succeeds or a timeout is reached.
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Operation,

    [Parameter(Mandatory = $true)]
    [ScriptBlock]$Condition,

    [Parameter()]
    [int]$TimeoutSeconds = 120,

    [Parameter()]
    [int]$PollIntervalSeconds = 5
  )

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  do {
    try {
      if (& $Condition) {
        Write-PSLog "$Operation completed within $TimeoutSeconds seconds." -Severity Debug
        return $true
      }
    }
    catch {
      Write-PSLog "$Operation still pending: $($_.Exception.Message)" -Severity Debug
    }

    if ((Get-Date) -ge $deadline) {
      break
    }

    Start-Sleep -Seconds $PollIntervalSeconds
  } while ($true)

  Write-PSLog "$Operation did not complete within $TimeoutSeconds seconds." -Severity Warning
  return $false
}

function Get-CachedDistributionGroups {
  <#
  .SYNOPSIS
  Retrieve and cache Exchange distribution groups for offboarding operations.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$Cache
  )

  if ($null -eq $Cache.Offboarding.DistributionGroups) {
    $Cache.Offboarding.DistributionGroups = @(
      Invoke-WithRetry -Operation 'Get distribution groups for offboarding cache' -ScriptBlock {
        Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
      }
    )
  }

  return $Cache.Offboarding.DistributionGroups
}

function Get-CachedMailboxes {
  <#
  .SYNOPSIS
  Retrieve and cache Exchange mailboxes for offboarding operations.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$Cache
  )

  if ($null -eq $Cache.Offboarding.Mailboxes) {
    $Cache.Offboarding.Mailboxes = @(
      Invoke-WithRetry -Operation 'Get mailboxes for offboarding cache' -ScriptBlock {
        Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
      }
    )
  }

  return $Cache.Offboarding.Mailboxes
}

function Get-CachedAutopilotDevicesForUser {
  <#
  .SYNOPSIS
  Retrieve cached Autopilot devices for a specific user principal name.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$Cache,

    [Parameter()]
    [string]$UserPrincipalName
  )

  if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    return @()
  }

  if (-not $Cache.Offboarding.AutopilotIndexed) {
    $autopilotByUserPrincipalName = @{}
    $autopilotDevices = @(
      Invoke-WithRetry -Operation 'Get Autopilot devices for offboarding cache' -ScriptBlock {
        Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All -ErrorAction Stop
      }
    )

    foreach ($autopilotDevice in $autopilotDevices) {
      $candidateUpns = @(
        $autopilotDevice.AssignedUserPrincipalName
        $autopilotDevice.UserPrincipalName
        $autopilotDevice.AdditionalProperties['assignedUserPrincipalName']
        $autopilotDevice.AdditionalProperties['userPrincipalName']
      ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLowerInvariant() } | Select-Object -Unique

      if (@($candidateUpns).Count -eq 0) {
        continue
      }

      $deviceSummary = [PSCustomObject]@{
        Id          = $autopilotDevice.Id
        DisplayName = $autopilotDevice.DisplayName
      }

      foreach ($candidateUpn in $candidateUpns) {
        if (-not $autopilotByUserPrincipalName.ContainsKey($candidateUpn)) {
          $autopilotByUserPrincipalName[$candidateUpn] = @()
        }

        $autopilotByUserPrincipalName[$candidateUpn] += $deviceSummary
      }
    }

    $Cache.Offboarding.AutopilotByUserPrincipalName = $autopilotByUserPrincipalName
    $Cache.Offboarding.AutopilotIndexed = $true
  }

  $cacheKey = $UserPrincipalName.ToLowerInvariant()
  if ($Cache.Offboarding.AutopilotByUserPrincipalName.ContainsKey($cacheKey)) {
    return @($Cache.Offboarding.AutopilotByUserPrincipalName[$cacheKey])
  }

  return @()
}

function Get-PerformanceStatistics {
  <#
  .SYNOPSIS
  Calculate and display performance statistics.

  .DESCRIPTION
  Analyzes processing times, throughput, and cache effectiveness.

  .PARAMETER StartTime
  Script start time.

  .PARAMETER UserCount
  Total number of users processed.

  .PARAMETER Cache
  Cache object with statistics.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [datetime]$StartTime,

    [Parameter(Mandatory = $true)]
    [int]$UserCount,

    [Parameter(Mandatory = $false)]
    [hashtable]$Cache
  )

  $duration = (Get-Date) - $StartTime
  $usersPerMinute = if ($duration.TotalMinutes -gt 0) { [math]::Round($UserCount / $duration.TotalMinutes, 2) } else { 0 }
  $avgSecondsPerUser = if ($UserCount -gt 0) { [math]::Round($duration.TotalSeconds / $UserCount, 2) } else { 0 }

  $stats = [PSCustomObject]@{
    TotalDuration     = $duration.ToString('hh\:mm\:ss')
    UsersProcessed    = $UserCount
    UsersPerMinute    = $usersPerMinute
    AvgSecondsPerUser = $avgSecondsPerUser
    CacheHits         = if ($Cache) { $Cache.Stats.Hits } else { 0 }
    CacheMisses       = if ($Cache) { $Cache.Stats.Misses } else { 0 }
    CacheHitRate      = if ($Cache -and ($Cache.Stats.Hits + $Cache.Stats.Misses) -gt 0) {
      [math]::Round(($Cache.Stats.Hits / ($Cache.Stats.Hits + $Cache.Stats.Misses)) * 100, 2)
    }
    else { 0 }
  }

  Write-PSLog "Total Duration: $($stats.TotalDuration)" -Severity Information
  Write-PSLog "Users Processed: $($stats.UsersProcessed)" -Severity Information
  Write-PSLog "Throughput: $($stats.UsersPerMinute) users/minute" -Severity Information
  Write-PSLog "Average: $($stats.AvgSecondsPerUser) seconds/user" -Severity Information

  if ($Cache) {
    Write-PSLog "Cache Hit Rate: $($stats.CacheHitRate)% ($($stats.CacheHits) hits / $($stats.CacheMisses) misses)" -Severity Information
  }

  return $stats
}

function Get-ErrorSummaryReport {
  <#
  .SYNOPSIS
  Generate comprehensive error summary report.

  .DESCRIPTION
  Analyzes all errors collected during processing and generates a detailed report
  with categorization, user impact, and recommendations.

  .PARAMETER ErrorSummary
  Error summary hashtable collected during processing.

  .PARAMETER SendEmail
  If specified, sends error report via email.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [hashtable]$ErrorSummary,

    [Parameter(Mandatory = $false)]
    [switch]$SendEmail
  )

  Write-PSLog '=== Error Summary ===' -Severity Information

  if ($ErrorSummary.TotalErrors -eq 0) {
    Write-PSLog 'No errors encountered during processing' -Severity Information
    return $null
  }

  Write-PSLog "Total Errors: $($ErrorSummary.TotalErrors)" -Severity Warning

  # Error breakdown by type
  if ($ErrorSummary.ErrorsByType.Count -gt 0) {
    Write-PSLog "`nErrors by Type:" -Severity Information
    foreach ($errorType in $ErrorSummary.ErrorsByType.Keys | Sort-Object) {
      $count = $ErrorSummary.ErrorsByType[$errorType]
      Write-PSLog "  - $errorType : $count" -Severity Warning
    }
  }

  # Users affected
  if ($ErrorSummary.ErrorsByUser.Count -gt 0) {
    Write-PSLog "`nUsers Affected: $($ErrorSummary.ErrorsByUser.Count)" -Severity Warning
    $topErrors = $ErrorSummary.ErrorsByUser.GetEnumerator() | Select-Object -First 10
    foreach ($userError in $topErrors) {
      Write-PSLog "  - $($userError.Key): $($userError.Value)" -Severity Warning
    }
    if ($ErrorSummary.ErrorsByUser.Count -gt 10) {
      Write-PSLog "  ... and $($ErrorSummary.ErrorsByUser.Count - 10) more" -Severity Warning
    }
  }

  # Critical errors
  if ($ErrorSummary.CriticalErrors.Count -gt 0) {
    Write-PSLog "`nCritical Errors: $($ErrorSummary.CriticalErrors.Count)" -Severity Error
    foreach ($criticalError in $ErrorSummary.CriticalErrors) {
      Write-PSLog "  - $criticalError" -Severity Error
    }
  }

  # Generate report object
  $report = [PSCustomObject]@{
    TotalErrors      = $ErrorSummary.TotalErrors
    ErrorsByType     = $ErrorSummary.ErrorsByType
    UsersAffected    = $ErrorSummary.ErrorsByUser.Count
    CriticalErrors   = $ErrorSummary.CriticalErrors.Count
    TopAffectedUsers = $ErrorSummary.ErrorsByUser.GetEnumerator() | Select-Object -First 10
    CorrelationId    = $Script:Config.CorrelationId
    Timestamp        = Get-Date
  }

  # Send email if requested
  if ($SendEmail -and $ErrorSummary.TotalErrors -gt 0) {
    try {
      $errorHtml = '<h2>BambooHR Sync Error Summary</h2>'
      $errorHtml += "<p><strong>Correlation ID:</strong> $($Script:Config.CorrelationId)</p>"
      $errorHtml += "<p><strong>Total Errors:</strong> $($ErrorSummary.TotalErrors)</p>"
      $errorHtml += "<p><strong>Users Affected:</strong> $($ErrorSummary.ErrorsByUser.Count)</p>"

      if ($ErrorSummary.ErrorsByType.Count -gt 0) {
        $errorHtml += '<h3>Errors by Type:</h3><ul>'
        foreach ($errorType in $ErrorSummary.ErrorsByType.Keys | Sort-Object) {
          $errorHtml += "<li><strong>$errorType</strong>: $($ErrorSummary.ErrorsByType[$errorType])</li>"
        }
        $errorHtml += '</ul>'
      }

      $params = @{
        Message         = @{
          Subject      = "BambooHR Sync Error Report - $($ErrorSummary.TotalErrors) Errors"
          Body         = @{
            ContentType = 'html'
            Content     = $errorHtml + "<p>$($Script:Config.Email.EmailSignature)</p>"
          }
          ToRecipients = @(
            @{
              EmailAddress = @{
                Address = $Script:Config.Email.NotificationEmailAddress
              }
            }
          )
        }
        SaveToSentItems = 'True'
      }

      Invoke-WithRetry -Operation 'Send error summary email' -ScriptBlock {
        Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress -Verbose
      }

      Write-PSLog "Error summary email sent to $($Script:Config.Email.NotificationEmailAddress)" -Severity Information
    }
    catch {
      Write-PSLog "Failed to send error summary email: $($_.Exception.Message)" -Severity Warning
    }
  }

  return $report
}

#endregion Performance Helper Functions

function Connect-ExchangeOnlineIfNeeded {
  <#
  .SYNOPSIS
  Connect to Exchange Online only if not already connected, with verbose suppression.

  .PARAMETER TenantId
  The tenant ID to connect to.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId
  )

  if (-not $Script:ExchangeConnected) {
    Write-PSLog "Connecting to Exchange Online for tenant: $TenantId" -Severity Debug
    $VerbosePreference = 'SilentlyContinue'
    # Mailbox permission cmdlets like Add/Remove-MailboxPermission may require an RPS session.
    $connectParams = @{
      ManagedIdentity = $true
      Organization    = $TenantId
      ShowBanner      = $false
    }

    $connectCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop
    if ($connectCommand.Parameters.ContainsKey('UseRPSSession')) {
      $connectParams.UseRPSSession = $true
    }

    Connect-ExchangeOnline @connectParams | Out-Null

    # Sanity check: if these cmdlets are missing, delegation sync will fail later.
    $requiredCmdlets = @('Get-MailboxPermission', 'Add-MailboxPermission', 'Remove-MailboxPermission', 'Add-RecipientPermission', 'Remove-RecipientPermission')
    $missingCmdlets = @($requiredCmdlets | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missingCmdlets.Count -gt 0) {
      Write-PSLog "Connected to Exchange Online but missing cmdlets: $($missingCmdlets -join ', '). Mailbox delegation sync requires an RPS session and the ExchangeOnlineManagement module." -Severity Error
      throw "Exchange Online cmdlets unavailable: $($missingCmdlets -join ', ')"
    }

    $Script:ExchangeConnected = $true
  }
  else {
    Write-PSLog 'Exchange Online connection already established, skipping reconnection' -Severity Debug
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
    $TeamsCardUri = $Script:Config.Features.TeamsCardUri,
    [Parameter()]
    [int]
    $MaximumExtraLicenses = 4,
    [Parameter()]
    [switch]
    $NewUser
  )

  try {
    Write-PSLog "Checking license status for SKU: $LicenseId" -Severity Debug

    $licenses = Invoke-WithRetry -Operation 'Get-MgSubscribedSku' -ScriptBlock {
      Get-MgSubscribedSku -SubscribedSkuId $LicenseId |
        Select-Object SkuPartNumber, SkuId, ConsumedUnits, @{
          Name       = 'EnabledUnits'
          Expression = { $_.PrepaidUnits.Enabled }
        }, @{
          Name       = 'SuspendedUnits'
          Expression = { $_.PrepaidUnits.Suspended }
        }, @{
          Name       = 'WarningUnits'
          Expression = { $_.PrepaidUnits.Warning }
        }
    }

    $licensesConsumed = $licenses.ConsumedUnits
    $licensesEnabled = $licenses.EnabledUnits

    if ($NewUser.IsPresent) {
      $licensesConsumed++
    }

    $licensesAvailable = $licensesEnabled - $licensesConsumed

    # Add AvailableUnits to the return object
    $licenses | Add-Member -MemberType NoteProperty -Name 'AvailableUnits' -Value $licensesAvailable -Force

    if ($licensesAvailable -lt 0 -and $NewUser.IsPresent) {
      Write-PSLog 'There are no licenses available for a newly created user!' -Severity Warning

      $params = @{
        Message         = @{
          Subject      = 'BhrEntraSync: There are no licenses available for a newly created user!'
          Body         = @{
            ContentType = 'html'
            Content     = "No licenses available for a newly created user. <br/> There are $($licensesConsumed) of $($licensesEnabled) assigned. $($Script:Config.Email.EmailSignature)"
          }
          ToRecipients = @(
            @{
              EmailAddress = @{
                Address = $Script:Config.Email.HelpDeskEmailAddress
              }
            }
          )
        }
        SaveToSentItems = 'True'
      }

      Invoke-WithRetry -Operation 'Send license warning email' -ScriptBlock {
        Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress
      }

      if ($TeamsCardUri) {
        New-AdaptiveCard {
          New-AdaptiveTextBlock -Text "There are $($licensesConsumed) of $($licensesEnabled) assigned!" -HorizontalAlignment Center -Weight Bolder -Size ExtraLarge
          New-AdaptiveTextBlock -Text 'The number of licenses should be increased' -Wrap -Weight Bolder -Size ExtraLarge
        } -Uri $TeamsCardUri -Speak 'There are no licenses left for a new user!'
      }
    }
    elseif ($licensesAvailable -le 0) {
      Write-PSLog 'There are no licenses available!' -Severity Error

      if ($TeamsCardUri) {
        New-AdaptiveCard {
          New-AdaptiveTextBlock -Text "There are $($licensesConsumed) of $($licensesEnabled) assigned!" -HorizontalAlignment Center -Weight Bolder -Size ExtraLarge
          New-AdaptiveTextBlock -Text 'The number of licenses should be increased' -Wrap -Weight Bolder -Size ExtraLarge
        } -Uri $TeamsCardUri -Speak 'There are no licenses left for a new user!'
      }
    }
    elseif ($licenses.ConsumedUnits -lt ($licensesEnabled - $MaximumExtraLicenses)) {
      Write-PSLog 'There are too many licenses left!' -Severity Warning

      $params = @{
        Message         = @{
          Subject      = 'BhrEntraSync: Too many extra licenses'
          Body         = @{
            ContentType = 'html'
            Content     = "Too many extra licenses. <br/> There are $($licensesConsumed) of $($licensesEnabled) assigned. $($Script:Config.Email.EmailSignature)"
          }
          ToRecipients = @(
            @{
              EmailAddress = @{
                Address = $Script:Config.Email.HelpDeskEmailAddress
              }
            }
          )
        }
        SaveToSentItems = 'True'
      }

      Invoke-WithRetry -Operation 'Send license excess warning email' -ScriptBlock {
        Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress
      }

      if ($TeamsCardUri) {
        New-AdaptiveCard {
          New-AdaptiveTextBlock -Text "There are $($licensesConsumed) of $($licensesEnabled) assigned!" -HorizontalAlignment Center -Weight Bolder
          New-AdaptiveTextBlock -Text 'Consider reducing the number of licenses' -Wrap -Weight Bolder -Size ExtraLarge
        } -Uri $TeamsCardUri -Speak 'Too many extra licenses!'
      }
    }
    else {
      Write-PSLog "$($licensesConsumed) of $($licensesEnabled) licenses $LicenseId have been assigned." -Severity Information
    }

    return $licenses
  }
  catch {
    Write-PSLog "Failed to check license status: $($_.Exception.Message)" -Severity Error
    throw
  }
}

function Get-NewPassword {
  <#
        .DESCRIPTION
            Generate a random password with the configured number of characters and special characters.
            Does not return characters that are commonly confused like 0 and O and 1 and l. Also removes characters
            that cause issues in PowerShell scripts.
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
  $password = ''

  # punctuation options but doesn't include &,',",`,$,{,},[,],(),),|,;,,
  # and a few others can break PowerShell or are difficult to read.
  $special = 43..46 + 94..95 + 126 + 33 + 35 + 61 + 63
  # Remove 0 and 1 because they can be confused with o,O,I,i,l
  $digits = 50..57
  # Remove O,o,i,I,l as these can be confused with other characters
  $letters = 65..72 + 74..78 + 80..90 + 97..104 + 106..107 + 109..110 + 112..122
  # Pick total minus the number of special chars of random letters and digits
  $chars = Get-Random -Count ($PasswordLength - $SpecialChars) -InputObject ($digits + $letters)
  # Pick the specified number of special characters
  $chars += Get-Random -Count $SpecialChars -InputObject ($special)
  # Mix up the chars so that the special char aren't just at the end and then convert each char number
  # to the char and put in a string
  $password = Get-Random -Count $PasswordLength -InputObject ($chars) |
    ForEach-Object -Begin { $aa = $null } -Process { $aa += [char]$_ } -End { $aa }

  return $password
}

function ConvertTo-StandardName {
  <#
        .SYNOPSIS
        Normalize a name for consistent comparison and display.

        .DESCRIPTION
        Replaces common diacritics, collapses whitespace, trims, and applies Title Case.

        .PARAMETER name
        Name string to normalize.
        #>
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $false)]
    [string]$name
  )

  if ([string]::IsNullOrWhiteSpace($name)) {
    return ''
  }

  $cleaned = $name -creplace 'Ă', 'A' -creplace 'ă', 'a' -creplace 'â', 'a' -creplace 'Â', 'A' -creplace 'Î', 'I' -creplace 'î', 'i' -creplace 'Ș', 'S' -creplace 'ș', 's' -creplace 'Ț', 'T' -creplace 'ț', 't'
  $cleaned = ($cleaned -replace '\s+', ' ').Trim()
  return (Get-Culture).TextInfo.ToTitleCase($cleaned)
}

function ConvertTo-PhoneNumber {
  <#
        .SYNOPSIS
        Normalize phone numbers for comparisons.

        .DESCRIPTION
  Preserves E.164 (+) or 00-prefixed international numbers; otherwise strips non-digits.

        .PARAMETER phoneNumber
        Phone number string to normalize.
        #>
  [CmdletBinding()]
  [OutputType([string])]
  param(
    [Parameter(Mandatory = $false)]
    [string]$phoneNumber
  )

  if ([string]::IsNullOrWhiteSpace($phoneNumber)) {
    return ''
  }

  $trimmed = $phoneNumber.Trim()
  if ($trimmed.StartsWith('+')) {
    $digits = ($trimmed -replace '[^0-9]', '')
    if ($digits.Length -gt 0) {
      return "+$digits"
    }
    return ''
  }

  if ($trimmed.StartsWith('00')) {
    $digits = ($trimmed -replace '[^0-9]', '')
    $digits = $digits.TrimStart('0')
    if ($digits.Length -gt 0) {
      return "+$digits"
    }
    return ''
  }

  $digits = ($trimmed -replace '[^0-9]', '')
  return $digits
}

function Get-ValidManagerUser {
  <#
        .SYNOPSIS
        Validate manager UPN and retrieve the user object.

        .DESCRIPTION
        Attempts to retrieve the manager user by UPN and logs if not found.

        .PARAMETER userPrincipalName
        Manager UPN or email address.

        .PARAMETER cache
        Optional performance cache for manager lookups.

        .PARAMETER targetUser
        User being processed (for logging context).
        #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [string]$userPrincipalName,

    [Parameter(Mandatory = $false)]
    [hashtable]$cache,

    [Parameter(Mandatory = $false)]
    [string]$targetUser
  )

  if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
    Write-PSLog -Message "Manager UPN missing for $targetUser; skipping manager update." -Severity Warning
    return $null
  }

  try {
    $managerUser = if ($cache) {
      Get-CachedUser -UserId $userPrincipalName -Cache $cache
    }
    else {
      Get-MgUser -UserId $userPrincipalName -ErrorAction Stop
    }
  }
  catch {
    Write-PSLog -Message "Manager lookup failed for $userPrincipalName (user: $targetUser). Error: $($_.Exception.Message)" -Severity Warning
    return $null
  }

  if (-not $managerUser) {
    Write-PSLog -Message "Manager UPN $userPrincipalName not found for $targetUser; skipping manager update." -Severity Warning
    return $null
  }

  return $managerUser
}

function Get-MgGroupMemberRecursively {
  <#
        .SYNOPSIS
        Get all members of a group recursively.
        .DESCRIPTION
        This function retrieves all members of a group recursively, including nested groups.
        .PARAMETER GroupId
        The ID of the group to retrieve members from.
        .PARAMETER GroupDisplayName
        The display name of the group to retrieve members from.
        #>
  [cmdletbinding()]

  param([Parameter()][string]$GroupId,
    [Parameter()][string]$GroupDisplayName
  )
  if ([string]::IsNullOrWhiteSpace($GroupId)) {
    if ([string]::IsNullOrWhiteSpace($GroupDisplayName)) {
      Write-PSLog 'Group lookup skipped: GroupId and GroupDisplayName are both empty.' -Severity Warning
      return @()
    }

    $safeGroupDisplayName = $GroupDisplayName.Replace("'", "''")
    $matchingGroups = @(Get-MgGroup -Filter "DisplayName eq '$safeGroupDisplayName'" -ErrorAction SilentlyContinue)

    if ($matchingGroups.Count -eq 0) {
      Write-PSLog "Group lookup returned no matches for display name '$GroupDisplayName'." -Severity Warning
      return @()
    }

    if ($matchingGroups.Count -gt 1) {
      $groupIds = ($matchingGroups | Select-Object -ExpandProperty Id) -join ', '
      Write-PSLog "Group lookup found multiple matches for '$GroupDisplayName'. Using first match. Candidate IDs: $groupIds" -Severity Warning
    }

    $GroupId = $matchingGroups[0].Id
    Write-PSLog "Resolved group '$GroupDisplayName' to group id '$GroupId'." -Severity Debug
  }

  $output = @()
  if ($GroupId) {
    Get-MgGroupMember -GroupId $GroupId -All | ForEach-Object {
      if ($_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
        $output += $_
      }
      if ($_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group') {
        $output += @(Get-MgGroupMemberRecursively -GroupId $_.Id)
      }
    }
  }
  return $output
}

function Sync-GroupMailboxDelegation {
  <#
    .SYNOPSIS
    You can assign a group as a mailbox delegate to allow all users delegate access to the mailbox.
    However, when a group is assigned, Outlook for Windows users will not get these delegate mailboxes automapped.
    The user must manually add the mailbox to their Outlook profile.
    If users are accessing mail using Outlook for web or Mac, automapping is not supported,
    so you can simply assign a group delegated permissions.

    .DESCRIPTION
    This script will add and remove delegates to an Exchange Online mailbox. Specify the group name and
    the mailbox for which to provide access.

    .PARAMETER Group
    The Entra ID (Azure AD) Group or Distribution group members to apply permissions

    .PARAMETER DelegateMailbox
    Mailbox to delegate access to

    .PARAMETER LeaveExistingDelegates
    Do not remove any of the existing delegates

    .PARAMETER Permissions
    Provide list of permissions to delegate. Default includes FullAccess and SendAs

    .PARAMETER DoNotConnect
    Specify when the PowerShell session is already properly authenticated with ExchangeOnline.
    Then it will not be connected again inside the function.
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
    $Permissions = @('FullAccess', 'SendAs'),
    [Parameter()]
    [string]
    $TenantId = $Script:Config.Azure.TenantId,
    [switch]$DoNotConnect
  )

  $fullAccessAdded = 0
  $fullAccessRemoved = 0
  $sendAsAdded = 0
  $sendAsRemoved = 0

  if ($DoNotConnect.IsPresent -eq $false) {
    Connect-ExchangeOnlineIfNeeded -TenantId $TenantId
  }

  # Find the shared mailbox
  if ([string]::IsNullOrWhiteSpace($DelegateMailbox) -eq $false) {
    $mObj = Get-EXOMailbox -Anr $DelegateMailbox
  }

  if ($null -eq $mObj) {
    Write-PSLog " Shared mailbox $DelegateMailbox not found!" -Severity Error
    exit 1
  }

  Write-PSLog "`t$DelegateMailbox matched with $($mObj) $($mObj.Identity) " -Severity Debug
  Connect-MgGraph -Identity -NoWelcome
  $gMembers = Get-MgGroupMemberRecursively -GroupDisplayName $Group | Sort-Object -Property Id -Unique
  Write-PSLog " $group member count: $($gMembers.Count)" -Severity Debug
  $groupMemberIds = @($gMembers | Where-Object { $_ -and $_.Id } | Select-Object -ExpandProperty Id)
  $groupMemberIdSample = ($groupMemberIds | Select-Object -First 10) -join ', '
  Write-PSLog " Group member ID sample (up to 10): $groupMemberIdSample" -Severity Information

  if (($gMembers.Count + 0) -eq 0) {
    Write-PSLog "No user members were resolved for group '$Group'. Mailbox delegation sync for '$DelegateMailbox' will be skipped." -Severity Warning
    return [pscustomobject]@{
      DelegateMailbox   = $DelegateMailbox
      FullAccessAdded   = 0
      FullAccessRemoved = 0
      SendAsAdded       = 0
      SendAsRemoved     = 0
      TotalAdded        = 0
      TotalRemoved      = 0
      TotalChanges      = 0
    }
  }

  if ($Permissions -contains 'FullAccess') {

    $existingFullAccessPermissions = @(Invoke-WithRetry -Operation 'Get mailbox permissions' -ScriptBlock {
        Get-EXOMailboxPermission -Identity $mObj.identity |
          Sort-Object -Property User -Unique | Where-Object { $_.User -notlike '*SELF' } |
          Sort-Object -Unique -Property User | ForEach-Object {
            Invoke-WithRetry -Operation "Get user $($_.User)" -ScriptBlock {
              Get-MgUser -UserId $_.User
            }
          }
      })
    $existingFullAccessPermissions = @($existingFullAccessPermissions | Where-Object { $_ -and $_.Id })
    $existingFullAccessIds = @($existingFullAccessPermissions | Select-Object -ExpandProperty Id)
    $existingFullAccessIdSample = ($existingFullAccessIds | Select-Object -First 10) -join ', '
    Write-PSLog " Existing FullAccess principal count: $($existingFullAccessIds.Count)" -Severity Information
    Write-PSLog " Existing FullAccess ID sample (up to 10): $existingFullAccessIdSample" -Severity Information

    $cPermissions = @()
    if (($gMembers.Count + 0) -gt 0) {
      $cPermissions = @(Compare-Object -ReferenceObject $existingFullAccessPermissions -DifferenceObject $gMembers -Property Id -ErrorAction SilentlyContinue)
    }
    if ($null -eq $cPermissions) {
      $cPermissions = @()
    }
    $missingPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ '=>'
    Write-PSLog " Missing perms: $($missingPermissions.Count + 0)" -Severity Debug
    $extraPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ '<='
    Write-PSLog " Extra perms: $($extraPermissions.Count + 0)" -Severity Debug

    # if need to add FullAccess
    if (($missingPermissions.Count + 0) -gt 0) {
      Write-PSLog "Adding $($missingPermissions.Count) missing permission(s) based on group membership" -Severity Information

      foreach ($missing in $missingPermissions) {
        $u = Get-MgUser -UserId $missing.id
        Write-PSLog "`tAdding Full Access permissions for $($u.DisplayName) to $($mObj.Identity) $DelegateMailbox..." -Severity Debug
        Add-MailboxPermission -Identity $mObj.Identity -User $missing.Id -AccessRights 'FullAccess' -Automapping:$true -InheritanceType All | Out-Null
        $fullAccessAdded++
      }
    }
    else {
      Write-PSLog "No Full Access permissions added to $($mObj.Identity)" -Severity Debug
    }

    if (($LeaveExistingDelegates.IsPresent -eq $false) -and (($extraPermissions.Count + 0) -gt 0)) {

      Write-PSLog "Removing $($extraPermissions.Count) extra permission(s) based on group membership" -Severity Debug
      foreach ($extra in $extraPermissions) {
        $u = Invoke-WithRetry -Operation "Get user $($extra.id)" -ScriptBlock {
          Get-MgUser -UserId $extra.id
        }
        Write-PSLog "`tRemoving Full Access $($u.DisplayName) permissions from $($mObj.Identity) $DelegateMailbox..." -Severity Debug
        Invoke-WithRetry -Operation "Remove mailbox permission for $($u.DisplayName)" -ScriptBlock {
          Remove-MailboxPermission -Identity $mObj.identity -User $extra.Id -Confirm:$false -AccessRights 'FullAccess' | Out-Null
        }
        $fullAccessRemoved++
      }
    }
    else {
      Write-PSLog "No Full Access permissions removed from $($mObj.Identity)." -Severity Debug
    }
  }

  # If need to add SendAs
  if ($Permissions -contains 'SendAs') {

    $existingSendAsPermissions = @(Invoke-WithRetry -Operation 'Get recipient permissions' -ScriptBlock {
        Get-EXORecipientPermission -Identity $mObj.identity |
          Where-Object { $_.Trustee -like '*@*' -and $_.AccessControlType -eq 'Allow' -and $_.AccessRights -contains 'SendAs' } |
          Sort-Object -Property Trustee -Unique | ForEach-Object {
            Invoke-WithRetry -Operation "Get user $($_.Trustee)" -ScriptBlock {
              Get-MgUser -UserId $_.Trustee
            }
          }
      })
    $existingSendAsPermissions = @($existingSendAsPermissions | Where-Object { $_ -and $_.Id })
    $existingSendAsIds = @($existingSendAsPermissions | Select-Object -ExpandProperty Id)
    $existingSendAsIdSample = ($existingSendAsIds | Select-Object -First 10) -join ', '
    Write-PSLog " Existing SendAs principal count: $($existingSendAsIds.Count)" -Severity Information
    Write-PSLog " Existing SendAs ID sample (up to 10): $existingSendAsIdSample" -Severity Information

    $cPermissions = @()
    if (($gMembers.Count + 0) -gt 0) {
      $cPermissions = @(Compare-Object -ReferenceObject $existingSendAsPermissions -DifferenceObject $gMembers -Property Id -ErrorAction SilentlyContinue)
    }
    if ($null -eq $cPermissions) {
      $cPermissions = @()
    }
    $missingPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ '=>'
    $extraPermissions = $cPermissions | Where-Object -Property SideIndicator -EQ '<='
    if (($missingPermissions.Count + 0) -gt 0) {
      Write-PSLog "Adding $($missingPermissions.Count) missing permission(s) based on group membership" -Severity Information

      foreach ($missing in $missingPermissions) {
        $u = Invoke-WithRetry -Operation "Get user $($missing.id)" -ScriptBlock {
          Get-MgUser -UserId $missing.id
        }
        Write-PSLog "`tAdding SendAs permissions for $($u.DisplayName) to $($mObj.Identity) $DelegateMailbox..." -Severity Debug
        Invoke-WithRetry -Operation "Add SendAs permission for $($u.DisplayName)" -ScriptBlock {
          Add-RecipientPermission -Identity $mObj.Id -Trustee $missing.Id -AccessRights 'SendAs' -Confirm:$false | Out-Null
        }
        $sendAsAdded++
      }
    }
    else {
      # Write-PSLog "No Send As permissions added to $DelegateMailbox" -Severity Debug
    }

    if (($LeaveExistingDelegates.IsPresent -eq $false) -and (($extraPermissions.Count + 0) -gt 0)) {

      Write-PSLog "Removing $($extraPermissions.Count) extra permission(s) based on group membership" -Severity Information
      foreach ($extra in $extraPermissions) {
        $u = Invoke-WithRetry -Operation "Get user $($extra.id)" -ScriptBlock {
          Get-MgUser -UserId $extra.id
        }
        Write-PSLog "`tRemoving SendAs permissions for $($u.DisplayName) to $($mObj.Identity) $DelegateMailbox..." -Severity Debug
        Invoke-WithRetry -Operation "Remove SendAs permission for $($u.DisplayName)" -ScriptBlock {
          Remove-RecipientPermission -Identity $mObj.identity -Trustee $extra.Id -Confirm:$false -AccessRights 'SendAs' | Out-Null
        }
        $sendAsRemoved++
      }
    }
    else {
      # Write-PSLog "No Send As permissions removed from $DelegateMailbox." -Severity Debug
    }
  }

  return [pscustomobject]@{
    DelegateMailbox   = $DelegateMailbox
    FullAccessAdded   = $fullAccessAdded
    FullAccessRemoved = $fullAccessRemoved
    SendAsAdded       = $sendAsAdded
    SendAsRemoved     = $sendAsRemoved
    TotalAdded        = ($fullAccessAdded + $sendAsAdded)
    TotalRemoved      = ($fullAccessRemoved + $sendAsRemoved)
    TotalChanges      = ($fullAccessAdded + $sendAsAdded + $fullAccessRemoved + $sendAsRemoved)
  }
}

function Invoke-SharedMailboxDelegationSync {
  <#
    .SYNOPSIS
    Runs shared mailbox delegation synchronization when enabled.

    .PARAMETER Reason
    Context string used for logging where this invocation occurred.
  #>
  [CmdletBinding()]
  param(
    [Parameter()]
    [string]
    $Reason = 'completion'
  )

  $createdCount = if ($Script:SignificantChanges -and $Script:SignificantChanges.Created) { $Script:SignificantChanges.Created.Count } else { 0 }
  $disabledCount = if ($Script:SignificantChanges -and $Script:SignificantChanges.Disabled) { $Script:SignificantChanges.Disabled.Count } else { 0 }
  $nameChangedCount = if ($Script:SignificantChanges -and $Script:SignificantChanges.NameChanged) { $Script:SignificantChanges.NameChanged.Count } else { 0 }
  $upnChangedCount = if ($Script:SignificantChanges -and $Script:SignificantChanges.UpnChanged) { $Script:SignificantChanges.UpnChanged.Count } else { 0 }
  $managerChangedCount = if ($Script:SignificantChanges -and $Script:SignificantChanges.ManagerChanged) { $Script:SignificantChanges.ManagerChanged.Count } else { 0 }
  $updatedMajorCount = if ($Script:SignificantChanges -and $Script:SignificantChanges.UpdatedMajor) { $Script:SignificantChanges.UpdatedMajor.Count } else { 0 }

  $hasUserChanges = ($createdCount + $disabledCount + $nameChangedCount + $upnChangedCount + $managerChangedCount + $updatedMajorCount) -gt 0
  $forceDelegationSync = [bool]$Script:Config.Features.ForceSharedMailboxPermissions

  if (-not ($forceDelegationSync -or $hasUserChanges)) {
    Write-PSLog "Skipping shared mailbox delegation sync ($Reason): ForceSharedMailboxPermissions is false and no user changes were detected." -Severity Information
    return
  }

  $delegationParams = @($Script:Config.Features.MailboxDelegationParams)
  if ($delegationParams.Count -eq 0) {
    Write-PSLog "Skipping shared mailbox delegation sync ($Reason): no mailbox delegation parameters were configured." -Severity Warning
    return
  }

  $triggerReason = if ($forceDelegationSync) { 'ForceSharedMailboxPermissions=true' } else { 'user changes detected' }
  Write-PSLog "Running shared mailbox delegation sync ($Reason) for $($delegationParams.Count) mailbox mapping(s). Trigger: $triggerReason." -Severity Information
  Connect-ExchangeOnlineIfNeeded -TenantId $Script:Config.Azure.TenantId
  $mailboxesChanged = 0
  $totalAdded = 0
  $totalRemoved = 0
  $totalFullAccessAdded = 0
  $totalFullAccessRemoved = 0
  $totalSendAsAdded = 0
  $totalSendAsRemoved = 0
  foreach ($params in $delegationParams) {
    $entry = @{}
    if ($params -is [hashtable]) {
      foreach ($k in $params.Keys) {
        $entry[[string]$k] = $params[$k]
      }
    }
    elseif ($params -ne $null) {
      foreach ($p in $params.PSObject.Properties) {
        $entry[$p.Name] = $p.Value
      }
    }

    $groupName = $null
    foreach ($name in @('Group', 'GroupName', 'EntraGroup', 'DistributionGroup')) {
      if ($entry.ContainsKey($name) -and -not [string]::IsNullOrWhiteSpace([string]$entry[$name])) {
        $groupName = [string]$entry[$name]
        break
      }
    }

    $delegateMailbox = $null
    foreach ($name in @('DelegateMailbox', 'DelegateMailBox', 'Mailbox', 'SharedMailbox', 'Delegate')) {
      if ($entry.ContainsKey($name) -and -not [string]::IsNullOrWhiteSpace([string]$entry[$name])) {
        $delegateMailbox = [string]$entry[$name]
        break
      }
    }

    if ([string]::IsNullOrWhiteSpace($groupName) -or [string]::IsNullOrWhiteSpace($delegateMailbox)) {
      $raw = if ($entry.Count -gt 0) { $entry | ConvertTo-Json -Compress } else { '<empty>' }
      Write-PSLog "Skipping mailbox delegation entry due to missing required values (Group/DelegateMailbox): $raw" -Severity Warning
      continue
    }

    $normalizedParams = @{
      Group           = $groupName
      DelegateMailbox = $delegateMailbox
    }
    if ($entry.ContainsKey('LeaveExistingDelegates')) { $normalizedParams.LeaveExistingDelegates = [bool]$entry['LeaveExistingDelegates'] }
    if ($entry.ContainsKey('Permissions') -and $null -ne $entry['Permissions']) { $normalizedParams.Permissions = @($entry['Permissions']) }
    if ($entry.ContainsKey('TenantId') -and -not [string]::IsNullOrWhiteSpace([string]$entry['TenantId'])) { $normalizedParams.TenantId = [string]$entry['TenantId'] }

    $syncResult = Sync-GroupMailboxDelegation @normalizedParams -DoNotConnect
    if ($syncResult) {
      if ($syncResult.TotalChanges -gt 0) {
        $mailboxesChanged++
      }
      $totalAdded += $syncResult.TotalAdded
      $totalRemoved += $syncResult.TotalRemoved
      $totalFullAccessAdded += $syncResult.FullAccessAdded
      $totalFullAccessRemoved += $syncResult.FullAccessRemoved
      $totalSendAsAdded += $syncResult.SendAsAdded
      $totalSendAsRemoved += $syncResult.SendAsRemoved
    }
  }

  Write-PSLog "Shared mailbox delegation summary ($Reason): mailboxes processed=$($delegationParams.Count), mailboxes changed=$mailboxesChanged, added=$totalAdded (FullAccess=$totalFullAccessAdded, SendAs=$totalSendAsAdded), removed=$totalRemoved (FullAccess=$totalFullAccessRemoved, SendAs=$totalSendAsRemoved)." -Severity Information
}

Write-PSLog "Executing Connect-MgGraph -TenantId $($Script:Config.Azure.TenantId)" -Severity Debug

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process

# Connect to Azure with system-assigned managed identity
if (-not $Script:AzureConnected) {
  $AzureContext = Invoke-WithRetry -Operation 'Connect-AzAccount' -ScriptBlock {
    $VerbosePreference = 'SilentlyContinue'
    (Connect-AzAccount -Identity).context
  }
  $Script:AzureConnected = $true
}
else {
  Write-PSLog 'Azure connection already established, skipping reconnection' -Severity Debug
}

# Set and store context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext

# Connect to Microsoft Graph
if (-not $Script:MgGraphConnected) {
  Invoke-WithRetry -Operation 'Connect-MgGraph' -ScriptBlock {
    $VerbosePreference = 'SilentlyContinue'
    Connect-MgGraph -Identity -NoWelcome
  }
  $Script:MgGraphConnected = $true
}
else {
  Write-PSLog 'Microsoft Graph connection already established, skipping reconnection' -Severity Debug
}

# Validate Graph connection
$testUser = Invoke-WithRetry -Operation 'Test Graph Connection' -ScriptBlock {
  Get-MgUser -UserId $Script:Config.Email.AdminEmailAddress -ErrorAction Stop
}

if ([string]::IsNullOrWhiteSpace($testUser)) {
  Write-PSLog "Unable to obtain user information using Get-MgUser -UserId $($Script:Config.Email.AdminEmailAddress)" -Severity Error
  exit 1
}

Write-PSLog 'Successfully connected to Microsoft Graph and validated access' -Severity Information

# Retrieve all verified tenant domains so we can match BambooHR emails against any
# valid UPN suffix — not just the primary company domain.
$Script:TenantDomains = @(
  Get-MgDomain -All | Where-Object { $_.IsVerified } | Select-Object -ExpandProperty Id
)
Write-PSLog "Tenant has $($Script:TenantDomains.Count) verified domain(s): $($Script:TenantDomains -join ', ')" -Severity Information

# Build URIs using configuration
$bhrRootUri = $Script:Config.BambooHR.RootUri
$bhrReportsUri = $Script:Config.BambooHR.ReportsUri

Write-PSLog "Starting BambooHR to Entra ID synchronization at $(Get-Date)" -Severity Information
Write-PSLog "Configuration: Company=$($Script:Config.Azure.CompanyName), BHR=$($Script:Config.BambooHR.CompanyName), Domain=$($Script:Config.Email.CompanyEmailDomain)" -Severity Information
# Provision users to Entra ID using the employee details from BambooHR
# Getting all users details from BambooHR and passing the extracted info to the variable $employees

Write-PSLog "Retrieving employee data from BambooHR API: $bhrReportsUri" -Severity Information

$headers = @{
  'Content-Type'  = 'application/json'
  'Authorization' = "Basic $([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($Script:Config.BambooHR.ApiKey):x")))"
}

$requestBody = @{
  fields = @(
    'status', 'hireDate', 'department', 'employeeNumber', 'firstName', 'lastName',
    'displayName', 'jobTitle', 'supervisorEmail', 'workEmail', 'lastChanged',
    'employmentHistoryStatus', 'bestEmail', 'location', 'workPhone', 'preferredName',
    'homeEmail', 'mobilePhone'
  )
} | ConvertTo-Json

$error.clear()

try {
  $response = Invoke-WithRetry -Operation 'BambooHR API Call' -ScriptBlock {
    Invoke-RestMethod -Uri $bhrReportsUri -Method POST -Headers $headers -ContentType 'application/json' -Body $requestBody -TimeoutSec $Script:Config.Runtime.OperationTimeoutSeconds
  }

  Write-PSLog "Successfully extracted employee information from BambooHR. Retrieved $($response.employees.Count) employee records." -Severity Information
}
catch {
  # If error returned, the API call to BambooHR failed and no usable employee data has been returned
  $errorMessage = "Error calling BambooHR API for user information. EXCEPTION MESSAGE: $($_.Exception.Message), CATEGORY: $($_.CategoryInfo.Category), SCRIPT STACK TRACE: $($_.ScriptStackTrace)"
  Write-PSLog $errorMessage -Severity Error

  # Send email alert with the generated error
  $params = @{
    Message         = @{
      Subject      = 'BhrEntraSync error: BambooHR connection failed'
      Body         = @{
        ContentType = 'html'
        Content     = "BambooHR connection failed. <br/> EXCEPTION MESSAGE: $($_.Exception.Message) <br/>CATEGORY: $($_.CategoryInfo.Category) <br/> SCRIPT STACK TRACE: $($_.ScriptStackTrace) <br/> $($Script:Config.Email.EmailSignature)"
      }
      ToRecipients = @(
        @{
          EmailAddress = @{
            Address = $Script:Config.Email.AdminEmailAddress
          }
        }
      )
    }
    SaveToSentItems = 'True'
  }

  try {
    Invoke-WithRetry -Operation 'Send BambooHR error notification' -ScriptBlock {
      Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress
    }
  }
  catch {
    Write-PSLog "Failed to send error notification email: $($_.Exception.Message)" -Severity Warning
  }

  if ($Script:Config.Features.TeamsCardUri) {
    try {
      New-AdaptiveCard {
        New-AdaptiveTextBlock -Text 'BambooHR API connection failed!' -Weight Bolder -Wrap -Color Red
        New-AdaptiveTextBlock -Text "Exception Message: $($_.Exception.Message)" -Wrap
        New-AdaptiveTextBlock -Text "Category: $($_.CategoryInfo.Category)" -Wrap
        New-AdaptiveTextBlock -Text "Correlation ID: $($Script:Config.CorrelationId)" -Wrap
      } -Uri $Script:Config.Features.TeamsCardUri -Speak 'BhrEntraSync error: BambooHR connection failed'
    }
    catch {
      Write-PSLog "Failed to send Teams notification: $($_.Exception.Message)" -Severity Warning
    }
  }

  exit 1
}

# Saving only the employee data to $employees variable and eliminate $response variable to save memory
$employees = $response.employees
$response = $null

# Delta sync: filter employees to only those modified within the configured window
if ($Script:Config.Features.FullSync) {
  Write-PSLog "Full sync: processing all $($employees.Count) employees (no date filter applied)" -Severity Information
}
else {
  $deltaCutoff = (Get-Date).AddDays(-$Script:Config.Features.ModifiedWithinDays)
  $totalCount = $employees.Count
  $employees = @($employees | Where-Object {
      if ([string]::IsNullOrWhiteSpace($_.lastChanged)) { return $true }
      try { [datetime]$_.lastChanged -ge $deltaCutoff } catch { $true }
    })
  Write-PSLog "Delta sync: $($employees.Count) of $totalCount employees modified within the last $($Script:Config.Features.ModifiedWithinDays) days" -Severity Information
  Write-PSLog 'Ensure a -FullSync run is scheduled daily or weekly so no changes are missed.' -Severity Warning
}

# Connect to Microsoft Graph using PS Graph Module, authenticating as the configured service principal for this operation,
# with certificate auth
$error.Clear()

if ($?) {
  # If no error returned, write to log file and continue
  Write-PSLog -Message "Successfully connected to Entra ID: $($Script:Config.Azure.TenantId)." -Severity Debug
}
else {

  # If error returned, write to log file and exit script
  Write-PSLog -Message "Connection to Entra ID failed.`n EXCEPTION: $($error.Exception) `n CATEGORY: $($error.CategoryInfo) `n ERROR ID: $($error.FullyQualifiedErrorId) `n SCRIPT STACK TRACE: $($error.ScriptStackTrace)" -Severity Error

  # Send email alert with the generated error
  $params = @{
    Message         = @{
      Subject      = 'BhrEntraSync error: Graph connection failed'
      Body         = @{
        ContentType = 'html'
        Content     = "<br/><br/>Microsoft Graph connection failed.<br/>EXCEPTION: $($error.Exception) <br/> CATEGORY:$($error.CategoryInfo) <br/> ERROR ID: $($error.FullyQualifiedErrorId) <br/>SCRIPT STACK TRACE: $mgErrStack <br/> $($Script:Config.Email.EmailSignature)"
      }
      ToRecipients = @(
        @{
          EmailAddress = @{
            Address = $AdminEmailAddress
          }
        }
      )
    }
    SaveToSentItems = 'True'
  }

  Invoke-WithRetry -Operation 'Send Graph connection error notification' -ScriptBlock {
    Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress -Verbose
  }

  New-AdaptiveCard {

    New-AdaptiveTextBlock -Text 'Entra ID Connection Failed' -Weight Bolder -Wrap
    New-AdaptiveTextBlock -Text "Exception Message $($_.Exception.Message)" -Wrap
    New-AdaptiveTextBlock -Text "Category: $($_.CategoryInfo.Category)" -Wrap
    New-AdaptiveTextBlock -Text "ERROR ID: $($error.FullyQualifiedErrorId)" -Wrap
    New-AdaptiveTextBlock -Text "SCRIPT STACK TRACE: $($_.ScriptStackTrace)" -Wrap
  } -Uri $TeamsCardUri -Speak 'BhrEntraSync error: Graph connection failed'

  Disconnect-MgGraph
  exit
}

Write-PSLog -Message "Looping through $($employees.Count) users." -Severity Debug
Write-PSLog -Message 'Selecting employees whose work email matches a verified tenant domain or who are inactive and must be evaluated by EmployeeID.' -Severity Debug

#region Main Processing Loop Setup
<#
===============================================================================
MAIN EMPLOYEE PROCESSING LOOP SETUP
===============================================================================

This section initializes tracking and optimization before processing employees.

WHAT'S BEING INITIALIZED:

1. PERFORMANCE CACHE:
   - Stores user lookups to avoid redundant API calls
   - Particularly effective for manager lookups (shared across many users)
   - Tracks hit/miss statistics for optimization analysis

2. PROCESSED USER COUNTER:
   - Counts how many employees were processed
   - Used for throughput calculation (users per minute)
   - Displayed in performance statistics at end

3. ERROR SUMMARY:
   - Collects all errors that occur during processing
   - Groups errors by type (UserCreation, AttributeUpdate, etc.)
   - Tracks which users were affected
   - Identifies critical vs. non-critical failures
   - Sends email report to admins at completion

DEVELOPER NOTE:
These initializations happen ONCE before the loop, not for each user.
The actual employee processing happens in the ForEach-Object loop below.
#>

# Initialize performance tracking
[hashtable]$performanceCache = Initialize-PerformanceCache
$processedUserCount = 0

# Initialize error tracking
$errorSummary = @{
  TotalErrors    = 0              # Total count of all errors
  ErrorsByType   = @{}           # Hashtable: ErrorType => Count
  ErrorsByUser   = @{}           # Hashtable: UserEmail => ErrorDescription
  CriticalErrors = @()         # Array of critical failure messages
  Warnings       = @()               # Array of warning messages
}

Write-PSLog "Using sequential processing (PowerShell $($PSVersionTable.PSVersion))" -Severity Information
#endregion Main Processing Loop Setup

#region Employee Processing Pipeline
<#
===============================================================================
EMPLOYEE PROCESSING PIPELINE
===============================================================================

This is the heart of the script - it processes each employee from BambooHR.

PIPELINE STAGES:
1. Filter: Employees with company email domain OR inactive (terminated) status
2. Sort: Process alphabetically by last name
3. ForEach: Process each employee individually

FOR EACH EMPLOYEE, THE SCRIPT:
1. Extracts data from BambooHR (name, email, job title, manager, etc.)
2. Normalizes data (trim whitespace, handle special characters)
3. Looks up existing Entra ID account:
   a. By UPN (work email) — skipped if BambooHR email is empty
   b. By EmployeeID — always attempted as secondary lookup
   c. Fallback: if UPN lookup fails but EID lookup succeeds, promotes the
      EID object to primary and overwrites $bhrWorkEmail with the Entra UPN.
      This ensures terminated users whose email was removed/changed can still
      be matched and disabled.
4. Determines action needed:
   - CREATE: New hire, no Entra ID account exists
   - UPDATE: Account exists, attributes need sync
   - DISABLE: Employee terminated in BambooHR
   - SKIP: No changes needed
5. Applies changes with retry logic
6. Sends notifications (email, Teams cards)
7. Updates cache and tracking counters

DATA FLOW:
  BambooHR → Extract → Normalize → Lookup (UPN → EID fallback) → Compare → Apply

DEVELOPER NOTE:
- Each iteration is independent (no shared state between users)
- Errors in one user don't stop processing of others
- All operations wrapped in retry logic for reliability
- Changes only applied if ShouldProcess confirms (not in -WhatIf mode)
- Inactive employees without a company email are included in the filter so their
  Entra ID accounts can be disabled via EmployeeID matching
#>

# Select employees with an email on a verified tenant domain, or inactive (terminated)
# employees who may need their Entra ID account disabled even if their BambooHR email
# was removed or changed.
$employees | Sort-Object -Property LastName |
  Where-Object {
    $email = $_.workEmail
    (Test-IsTenantEmail -EmailAddress $email -TenantDomains $Script:TenantDomains) -or $_.status -eq 'Inactive'
  } | ForEach-Object {
    $error.Clear()

    <#
    =============================================================================
    EMPLOYEE DATA EXTRACTION FROM BAMBOOHR
    =============================================================================

    For each employee object from BambooHR API, extract all fields into variables.
    Using variables makes the code more readable and easier to debug.

    FIELD NAMING CONVENTION:
    - Prefix "bhr" indicates data from BambooHR
    - Prefix "entra" (used later) indicates data from Entra ID.
    - This makes it clear which system is the source of truth

    COMMON FIELDS:
    - lastChanged: Timestamp of last BambooHR update (for change detection)
    - hireDate: Employee start date (used for pre-hire account creation)
    - employeeNumber: Unique employee ID (synced to EmployeeId in Entra ID)
    - jobTitle: Current job title
    - department: Department name
    - supervisorEmail: Manager's email (used for org hierarchy)
    - workEmail: Primary company email (becomes UPN in Entra ID)
    - status: Active or Inactive (determines account enabled state)
    - location: Office location
    - workPhone/mobilePhone: Contact numbers

    #>

    # Metadata fields
    $bhrlastChanged = "$($_.lastChanged)"           # Last modified timestamp in BambooHR
    $bhrHireDate = "$($_.hireDate)"                 # Employee hire date
    $bhremployeeNumber = "$($_.employeeNumber)"     # Unique employee number
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
    # Translating user "status" from BambooHR to boolean, to match and compare with the Entra ID user account status
    $bhrStatus = "$($_.status)"
    if ($bhrStatus -eq 'Inactive')
    { $bhrAccountEnabled = $False }
    if ($bhrStatus -eq 'Active')
    { $bhrAccountEnabled = $True }
    $bhrOfficeLocation = "$($_.location)"
    $bhrPreferredName = ConvertTo-StandardName "$($_.preferredName)"
    $bhrWorkPhone = ConvertTo-PhoneNumber "$($_.workPhone)"
    $bhrMobilePhone = ConvertTo-PhoneNumber "$($_.mobilePhone)"
    $bhrBestEmail = "$($_.bestEmail)"
    $bhrFirstName = ConvertTo-StandardName $_.firstName
    # First name of employee in Bamboo HR
    $bhrLastName = ConvertTo-StandardName $_.lastName
    # The Display Name of the user in BambooHR
    $bhrDisplayName = ConvertTo-StandardName $_.displayName
    $bhrHomeEmail = "$($homeEmail)"

    if ($bhrPreferredName -ne $bhrFirstName -and [string]::IsNullorWhitespace($bhrPreferredName) -eq $false) {
      Write-PSLog -Message "User preferred first name of $bhrPreferredName instead of legal name $bhrFirstName" -Severity Debug
      $bhrFirstName = $bhrPreferredName
      $bhrDisplayName = ConvertTo-StandardName "$bhrPreferredName $bhrLastName"
    }

    Write-PSLog -Message "BambooHR employee: $bhrFirstName $bhrLastName ($bhrDisplayName) $bhrWorkEmail" -Severity Debug
    Write-PSLog -Message "Department: $bhrDepartment, Title: $bhrJobTitle, Manager: $bhrSupervisorEmail HireDate: $bhrHireDate" -Severity Debug
    Write-PSLog -Message "EmployeeId: $bhrEmployeeNumber, Status: $bhrStatus, Employee Status: $bhrEmploymentStatus" -Severity Debug
    Write-PSLog -Message "Location: $bhrOfficeLocation, PreferredName: $bhrPreferredName, BestEmail: $bhrBestEmail HomeEmail: $bhrHomeEmail, WorkPhone: $bhrWorkPhone" -Severity Debug

    $entraIdUpnObjDetails = $null
    $entraIdEidObjDetails = $null

    <#
            If the user start date is in the past, or in less than -DaysAhead days from current time,
            we can begin processing the user: create Entra ID account or correct the attributes in Entra ID for the employee,
            else, the employee found on BambooHR will not be processed
  #>

    [datetime]$parsedHireDate = ConvertTo-BambooHrHireDate -HireDate $bhrHireDate -EmployeeEmailAddress $bhrWorkEmail
    [string]$normalizedBhrHireDate = $parsedHireDate.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)

    if ($parsedHireDate -le (Get-Date).AddDays($Script:Config.Features.DaysAhead)) {

      $error.clear()

      # Check if the user exists in Entra ID and if there is an account with the EmployeeID of the user checked
      # in the current loop

      # Lookup user by UPN (email address) - capture return value from Invoke-WithRetry
      # Guard: skip UPN lookup if BambooHR work email is empty or doesn't match any
      # verified tenant domain — non-tenant emails will never be a valid Entra UPN.
      $isValidTenantEmail = Test-IsTenantEmail -EmailAddress $bhrWorkEmail -TenantDomains $Script:TenantDomains
      if ($isValidTenantEmail) {
        Write-PSLog -Message "Validating $bhrWorkEmail Entra ID account." -Severity Information
        $entraIdUpnObjDetails = Invoke-WithRetry -Operation "Get user by UPN: $bhrWorkEmail" -ScriptBlock {
          Get-MgUser -UserId $bhrWorkEmail -Property id, userprincipalname, Department, EmployeeId, JobTitle, CompanyName, Surname, GivenName, DisplayName, AccountEnabled, Mail, EmployeeHireDate, OfficeLocation, BusinessPhones, MobilePhone, OnPremisesExtensionAttributes, AdditionalProperties -ExpandProperty manager -ErrorAction SilentlyContinue
        }
      }
      else {
        Write-PSLog -Message "Skipping UPN lookup for $bhrDisplayName (EID $bhrEmployeeNumber) — email '$bhrWorkEmail' is not on a verified tenant domain. Falling back to EmployeeID." -Severity Debug
      }

      # Lookup user by EmployeeID - capture return value from Invoke-WithRetry
      $entraIdEidObjDetails = Invoke-WithRetry -Operation "Get user by EmployeeID: $bhrEmployeeNumber" -ScriptBlock {
        Get-MgUser -Filter "employeeID eq '$bhrEmployeeNumber'" -Property employeeid, userprincipalname, Department, JobTitle, CompanyName, Surname, GivenName, DisplayName, MobilePhone, AccountEnabled, Mail, OfficeLocation, BusinessPhones , EmployeeHireDate, OnPremisesExtensionAttributes, AdditionalProperties -ExpandProperty manager -ErrorAction SilentlyContinue
      }
      $error.clear()

      if ([string]::IsNullOrEmpty($entraIdUpnObjDetails) -eq $false) {
        $UpnExtensionAttribute1 = ($entraIdUpnObjDetails |
            Select-Object @{
              Name       = 'ExtensionAttribute1'
              Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 }
            } -ErrorAction SilentlyContinue).ExtensionAttribute1
        }

        if ([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false) {
          $EIDExtensionAttribute1 = ($entraIdEidObjDetails |
              Select-Object @{
                Name       = 'ExtensionAttribute1'
                Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 }
              } -ErrorAction SilentlyContinue).ExtensionAttribute1
          }

          # Fallback: If UPN lookup failed but EID lookup found the user, use the EID object as the
          # primary reference. This handles terminated employees whose BambooHR email was removed or
          # changed but whose Entra ID account can still be found by EmployeeId.
          if ([string]::IsNullOrEmpty($entraIdUpnObjDetails) -and ([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false)) {
            Write-PSLog -Message "UPN lookup failed for '$bhrWorkEmail' but found Entra account by EmployeeID $bhrEmployeeNumber (UPN: $($entraIdEidObjDetails.UserPrincipalName)). Using EmployeeID match as primary." -Severity Warning
            $entraIdUpnObjDetails = $entraIdEidObjDetails
            $bhrWorkEmail = $entraIdEidObjDetails.UserPrincipalName
            $UpnExtensionAttribute1 = $EIDExtensionAttribute1
          }

          # Saving Entra ID attributes to be compared one by one with the details pulled from BambooHR
          $entraIdWorkEmail = "$($entraIdUpnObjDetails.Mail)"
          $entraIdJobTitle = "$($entraIdUpnObjDetails.JobTitle)"
          $entraIdDepartment = "$($entraIdUpnObjDetails.Department)"
          [bool]$entraIdStatus = [bool]$entraIdUpnObjDetails.AccountEnabled
          $entraIdEmployeeNumber = "$($entraIdUpnObjDetails.EmployeeId)"
          $entraIdEmployeeNumber2 = "$($entraIdEidObjDetails.EmployeeId)"
          $entraIdSupervisorEmail = "$(($entraIdUpnObjDetails |
            Select-Object @{Name = 'Manager'; Expression = { $_.Manager.AdditionalProperties.mail } }).Manager)"
          $entraIdDisplayname = "$($entraIdUpnObjDetails.displayName)"
          $entraIdFirstName = "$($entraIdUpnObjDetails.GivenName)"
          $entraIdLastName = "$($entraIdUpnObjDetails.Surname)"
          $entraIdCompanyName = "$($entraIdUpnObjDetails.CompanyName)"
          $entraIdWorkPhone = "$($entraIdUpnObjDetails.BusinessPhones)"
          $entraIdMobilePhone = "$($entraIdUpnObjDetails.MobilePhone)"
          $entraIdOfficeLocation = "$($entraIdUpnObjDetails.OfficeLocation)"

          # Clean up phone info to make it easier to work with
          [string]$bhrWorkPhone = ConvertTo-PhoneNumber $bhrWorkPhone
          [string]$entraIdWorkPhone = [int64]($entraIdWorkPhone -replace '[^0-9]', '') -replace '^1', ''
          [string]$bhrMobilePhone = [int64]($bhrMobilePhone -replace '[^0-9]', '') -replace '^1', ''
          [string]$entraIdMobilePhone = [int64]($entraIdMobilePhone -replace '[^0-9]', '') -replace '^1', ''

          if ($entraIdUpnObjDetails.EmployeeHireDate) {
            $entraIdHireDate = $entraIdUpnObjDetails.EmployeeHireDate.AddHours(12).ToString('yyyy-MM-dd')
          }

          Write-PSLog -Message "Entra ID Upn Obj Details: '$([string]::IsNullOrEmpty($entraIdUpnObjDetails) -eq $false)' EntraIdEidObjDetails: $([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false) = $(([string]::IsNullOrEmpty($entraIdUpnObjDetails) -eq $false) -or ([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false))" -Severity Debug

          <#
          USER ACCOUNT EXISTS CHECK:
          If we found a user by UPN OR by EmployeeId, then the account exists in Entra ID.
          This section handles UPDATES to existing accounts (attributes, manager, status).
          If neither lookup returned a user, we'll create the account later (see "else" block).
          #>
          if (([string]::IsNullOrEmpty($entraIdUpnObjDetails) -eq $false) -or ([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false)) {
            Write-PSLog -Message "Entra Id user: $entraIdFirstName $entraIdLastName ($entraIdDisplayName) $entraIdWorkEmail" -Severity Debug
            Write-PSLog -Message "Department: $entraIdDepartment, Title: $entraIdJobTitle, Manager: $entraIdSupervisorEmail, HireDate: $entraIdHireDate" -Severity Debug
            Write-PSLog -Message "EmployeeId: $entraIdEmployeeNumber, Enabled: $entraIdStatus OfficeLocation: $entraIdOfficeLocation, WorkPhone: $entraIdWorkPhone" -Severity Debug

            <#
            ATTRIBUTE SYNC CHECK:
            This complex condition determines if we need to update the user's attributes.

            WE UPDATE ATTRIBUTES WHEN:
            1. EmployeeId matches in both systems (same person)
            2. Object IDs match (data consistency check)
            3. UPN matches work email (correct account)
            4. Last changed date differs (BambooHR has newer data)
            5. Account is not suspended
            6. Both lookups returned valid data

            WHY CHECK LAST CHANGED DATE:
            - Avoids unnecessary API calls if nothing changed
            - BambooHR lastChanged timestamp stored in ExtensionAttribute1
            - Only sync when BambooHR has updates

            ATTRIBUTES WE SYNC:
            - DisplayName, GivenName, Surname
            - JobTitle, Department
            - OfficeLocation
            - BusinessPhones, MobilePhone
            - Manager relationship
            - AccountEnabled status
            #>

            Write-PSLog -Message "Entra Id Employee Number: $entraIdEmployeeNumber -eq $entraIdEmployeeNumber2 = $($entraIdEmployeeNumber -eq $entraIdEmployeeNumber2) -and `
            $($entraIdEidObjDetails.UserPrincipalName) -eq $($entraIdUpnObjDetails.UserPrincipalName) -eq $bhrWorkEmail = $($entraIdEidObjDetails.UserPrincipalName -eq $entraIdUpnObjDetails.UserPrincipalName -eq $bhrWorkEmail) -and `
            $($entraIdUpnObjDetails.id) -eq $($entraIdEidObjDetails.id) = $($entraIdUpnObjDetails.id -eq $entraIdEidObjDetails.id) -and `
            $bhrLastChanged -ne $UpnExtensionAttribute1 = $($bhrLastChanged -ne $UpnExtensionAttribute1) -and `
            $($entraIdEidObjDetails.Capacity) -ne 0 -and $($entraIdUpnObjDetails.Capacity) -ne 0 = $($entraIdEidObjDetails.Capacity -ne 0 -and $entraIdUpnObjDetails.Capacity -ne 0) -and `
            $bhrEmploymentStatus -notlike '*suspended*' = $($bhrEmploymentStatus -notlike '*suspended*') " -Severity Debug

            <#
            EMPLOYEEID MISMATCH FIX:
            Edge case: EmployeeId doesn't match but UPN does.
            This can happen if:
            - User was created manually before BambooHR sync
            - EmployeeId was corrected in BambooHR
            - Data migration issues
            In this case, we update the EmployeeId in Entra ID to match BambooHR.
            #>
            if (Test-ShouldUpdateEmployeeId -EntraIdEmployeeNumber $entraIdEmployeeNumber `
                -BhrEmployeeNumber $bhrEmployeeNumber `
                -EntraIdUserPrincipalName $entraIdUpnObjDetails.UserPrincipalName `
                -BhrWorkEmail $bhrWorkEmail `
                -BhrEmploymentStatus $bhrEmploymentStatus `
                -BhrAccountEnabled $bhrAccountEnabled `
                -BhrLastChanged $bhrLastChanged `
                -UpnExtensionAttribute1 $UpnExtensionAttribute1) {
              # Employee number in Entra Id does not match the one in BambooHR, but the UPN matches. Update the employee number in Entra ID.
              Write-PSLog -Message "Entra Id Employee number $entraIdEmployeeNumber does not match BambooHR $bhrEmployeeNumber, but the UPN matches. Update the employee number in Entra ID." -Severity Debug
              $error.clear()
              if ($PSCmdlet.ShouldProcess($bhrWorkEmail, "Update EmployeeId to '$bhremployeeNumber'")) {
                Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber" -Severity Debug
                Invoke-WithRetry -Operation "Update user EmployeeId: $bhrWorkEmail" -ScriptBlock {
                  Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber
                }
                $entraIdEmployeeNumber = $bhrEmployeeNumber
              }
            }

            if (Test-ShouldSyncExistingUser -EntraIdEmployeeNumber $entraIdEmployeeNumber `
                -EntraIdEmployeeNumberByEid $entraIdEmployeeNumber2 `
                -EntraIdUpnFromEidLookup $entraIdEidObjDetails.UserPrincipalName `
                -EntraIdUpnFromUpnLookup $entraIdUpnObjDetails.UserPrincipalName `
                -BhrWorkEmail $bhrWorkEmail `
                -BhrLastChanged $bhrLastChanged `
                -UpnExtensionAttribute1 $UpnExtensionAttribute1 `
                -EntraIdEidObjDetails $entraIdEidObjDetails `
                -EntraIdUpnObjDetails $entraIdUpnObjDetails `
                -BhrEmploymentStatus $bhrEmploymentStatus) {

              Write-PSLog -Message "$bhrWorkEmail is a valid Entra ID Account, with matching EmployeeId and UPN in Entra ID and BambooHR, but different last modified date." -Severity Debug
              $error.clear()

              # Check if user is active in BambooHR, and set the status of the account as it is in BambooHR
              # (active or inactive)
              if ($bhrAccountEnabled -eq $false -and $bhrEmploymentStatus.Trim() -eq 'Terminated' -and $entraIdStatus -eq $true ) {
                Write-PSLog -Message "$bhrWorkEmail is marked 'Inactive' in BHR and 'Active' in Entra ID. Blocking sign-in, revoking sessions, changing pw, removing auth methods"
                # The account is marked "Inactive" in BHR and "Active" in Entra ID, block sign-in, revoke sessions,
                #change pass, remove auth methods
                $error.clear()
                if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Disable Account (Revoke Sessions and Block Sign-In)')) {
                  Write-PSLog -Message "Executing: Revoke-MgUserSignInSession -UserId $bhrWorkEmail" -Severity Debug
                  Revoke-MgUserSignInSession -UserId $bhrWorkEmail
                  Start-Sleep 10
                  Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Debug
                  Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled
                }

                # Change to a random password that is not known to the user.
                $params = @{
                  PasswordProfile = @{
                    ForceChangePasswordNextSignIn = $true
                    Password                      = (Get-NewPassword)
                  }
                }

                Write-PSLog -Message "User $bhrWorkEmail is no longer active in BambooHR, disabling Entra Id account and offboarding user." -Severity Information
                Add-SignificantChange -Category Disabled -User $bhrWorkEmail -Detail 'Terminated in BambooHR'
                Write-PSLog -Message "Disabled user: $bhrWorkEmail" -Severity Information
                Write-PSLog -Message "Logging enterprise app role assignments and user consents for $bhrWorkEmail" -Severity Information
                $userObjectId = $entraIdUpnObjDetails.Id
                $appRoleAssignmentCommand = Get-Command Get-MgUserAppRoleAssignment -ErrorAction SilentlyContinue
                if ($appRoleAssignmentCommand) {
                  try {
                    $appRoleAssignments = Get-MgUserAppRoleAssignment -UserId $userObjectId -All -ErrorAction Stop
                    foreach ($assignment in $appRoleAssignments) {
                      $resourceName = $assignment.ResourceDisplayName
                      if ([string]::IsNullOrWhiteSpace($resourceName)) {
                        $servicePrincipalCommand = Get-Command Get-MgServicePrincipal -ErrorAction SilentlyContinue
                        if ($servicePrincipalCommand) {
                          $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $assignment.ResourceId -ErrorAction SilentlyContinue
                          $resourceName = if ($servicePrincipal) { $servicePrincipal.DisplayName } else { $assignment.ResourceId }
                        }
                        else {
                          $resourceName = $assignment.ResourceId
                        }
                      }
                      Write-PSLog -Message "App role assignment: App='$resourceName' RoleId='$($assignment.AppRoleId)'" -Severity Debug
                    }
                    if (-not $appRoleAssignments) {
                      Write-PSLog -Message "No app role assignments found for $bhrWorkEmail" -Severity Debug
                    }
                  }
                  catch {
                    Write-PSLog -Message "Failed to log app role assignments for $($bhrWorkEmail): $($_.Exception.Message)" -Severity Warning
                  }
                }
                else {
                  Write-PSLog -Message 'Get-MgUserAppRoleAssignment not available; skipping app role assignment logging' -Severity Warning
                }

                $consentCommand = Get-Command Get-MgOauth2PermissionGrant -ErrorAction SilentlyContinue
                if ($consentCommand) {
                  try {
                    $consents = Get-MgOauth2PermissionGrant -Filter "principalId eq '$userObjectId'" -All -ErrorAction Stop
                    foreach ($consent in $consents) {
                      Write-PSLog -Message "User consent: ClientId='$($consent.ClientId)' Scope='$($consent.Scope)'" -Severity Debug
                    }
                    if (-not $consents) {
                      Write-PSLog -Message "No user consents found for $bhrWorkEmail" -Severity Debug
                    }
                  }
                  catch {
                    Write-PSLog -Message "Failed to log user consents for $($bhrWorkEmail): $($_.Exception.Message)" -Severity Warning
                  }
                }
                else {
                  Write-PSLog -Message 'Get-MgOauth2PermissionGrant not available; skipping user consent logging' -Severity Warning
                }
                if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Terminate User Account (Change Password, Update Profile, Convert to Shared Mailbox, Remove Licenses and Groups)')) {
                  Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params" -Severity Debug
                  Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params
                  $leaveDateTimeUtc = (Get-Date).ToUniversalTime()
                  Set-TerminatedUserProfileFields -UserId $bhrWorkEmail -LeaveDateTimeUtc $leaveDateTimeUtc
                  Get-MgUserMemberOf -UserId $bhrWorkEmail

                  # TODO: Does not work for on premises synced accounts. Not a problem with Entra Id native.
                  try {
                    $null = Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrLastChanged } -UserId $bhrWorkEmail -ErrorAction Stop | Out-Null
                    Write-PSLog -Message "$bhrWorkEmail LastChanged attribute set from '$upnExtensionAttribute1' to '$bhrlastChanged'." -Severity Information
                  }
                  catch {
                    $error.Clear()
                  }

                  # Cancel all meetings for the user
                  Write-PSLog -Message "Attempting to cancel meetings for $bhrWorkEmail" -Severity Information
                  try {
                    Write-PSLog -Message "Executing: Get-MgUserEvent -UserId $bhrWorkEmail | ForEach-Object { Remove-MgUserEvent -UserId $bhrWorkEmail -EventId `$_.id }" -Severity Debug
                    $userEvents = Get-MgUserEvent -UserId $bhrWorkEmail -ErrorAction Stop
                    if ($userEvents) {
                      $userEvents | ForEach-Object {
                        Remove-MgUserEvent -UserId $bhrWorkEmail -EventId $_.id -ErrorAction SilentlyContinue
                      } | Out-Null
                      Write-PSLog -Message "Successfully canceled meetings for $bhrWorkEmail" -Severity Information
                    }
                    else {
                      Write-PSLog -Message "No meetings found for $bhrWorkEmail" -Severity Debug
                    }
                  }
                  catch {
                    Write-PSLog -Message "Unable to cancel meetings for $bhrWorkEmail. This requires Calendars.ReadWrite permission. Error: $($_.Exception.Message)" -Severity Warning
                    $errorSummary.Warnings += "Calendar access denied for $bhrWorkEmail - missing Calendars.ReadWrite permission"
                  }

                  # Set the out of office for the user that they are no longer with the company and to contact the manager
                  Connect-ExchangeOnlineIfNeeded -TenantId $Script:Config.Azure.TenantId

                  $internalMessage = "I am no longer with the company. Please contact $($entraIdSupervisorEmail) for assistance."
                  $externalMessage = "My role has changed, please contact $($entraIdSupervisorEmail) for assistance."

                  # set the automatic replies for the user
                  Write-PSLog -Message "Setting automatic replies for $bhrWorkEmail" -Severity Information
                  Write-PSLog -Message "Executing: Set-MailboxAutoReplyConfiguration -Identity $bhrWorkEmail -AutoReplyState Enabled -ExternalAudience All" -Severity Debug
                  Invoke-WithRetry -Operation "Set automatic replies for: $bhrWorkEmail" -ScriptBlock {
                    Set-MailboxAutoReplyConfiguration -Identity $bhrWorkEmail `
                      -AutoReplyState Enabled `
                      -ExternalAudience All `
                      -InternalMessage $internalMessage `
                      -ExternalMessage $externalMessage `
                      -ErrorAction Stop
                  }

                  # Determine if the user was an owner of any groups and assign ownership to the manager
                  Write-PSLog -Message "Checking to see if there are groups owned by $bhrWorkEmail" -Severity Information
                  $groups = Get-MgUserMemberOf -UserId $bhrWorkEmail -ErrorAction SilentlyContinue
                  if ($groups) {
                    $groups | ForEach-Object {
                      $group = $_
                      if ($group.Owners -contains $bhrWorkEmail) {
                        Write-PSLog -Message "User $bhrWorkEmail is an owner of group $($group.DisplayName). Reassigning ownership to $($entraIdSupervisorEmail)." -Severity Information
                      }
                    }
                  }

                  # Convert mailbox to shared
                  Connect-ExchangeOnlineIfNeeded -TenantId $TenantId

                  Write-PSLog -Message "Executing: Set-Mailbox -Identity $bhrWorkEmail -Type Shared" -Severity Debug
                  Write-PSLog -Message "Converting $bhrWorkEmail to a shared mailbox..." -Severity Debug
                  Set-Mailbox -Identity $bhrWorkEmail -Type Shared
                  $null = Wait-ForCondition -Operation "Mailbox conversion to shared for $bhrWorkEmail" -TimeoutSeconds 120 -PollIntervalSeconds 10 -Condition {
                    (Get-EXOMailbox -Anr $bhrWorkEmail -ErrorAction Stop).RecipientTypeDetails -eq 'SharedMailbox'
                  }

                  # Give permissions to converted mailbox to previous manager
                  $mObj = Get-EXOMailbox -Anr $bhrWorkEmail
                  Write-PSLog "`t$($entraIdSupervisorEmail) being given permissions to $bhrWorkEmail now..." -Severity Information
                  Write-PSLog "Executing: Add-MailboxPermission -Identity $($mObj.Id) -User $entraIdSupervisorEmail -AccessRights 'FullAccess' -Automapping:$true -InheritanceType All" -Severity Debug
                  try {
                    Add-MailboxPermission -Identity $mObj.Id -User $entraIdSupervisorEmail -AccessRights 'FullAccess' -Automapping:$true -InheritanceType All -ErrorAction Stop | Out-Null
                  }
                  catch {
                    Write-PSLog -Message "Failed to add mailbox permission for $bhrWorkEmail to $entraIdSupervisorEmail. Error: $($_.Exception.Message)" -Severity Warning
                  }

                  # Grant manager OneDrive access and send the link (30-day access window)
                  if (-not [string]::IsNullOrWhiteSpace($entraIdSupervisorEmail)) {
                    $driveCommand = Get-Command Get-MgUserDrive -ErrorAction SilentlyContinue
                    $inviteCommand = Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue
                    if ($driveCommand -and $inviteCommand) {
                      try {
                        $drive = Invoke-WithRetry -Operation "Get OneDrive for $bhrWorkEmail" -ScriptBlock {
                          Get-MgUserDrive -UserId $bhrWorkEmail -ErrorAction Stop
                        }

                        $driveUrl = $drive.WebUrl
                        $expirationUtc = (Get-Date).ToUniversalTime().AddDays(30)
                        $expirationString = $expirationUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

                        $inviteBody = @{
                          recipients         = @(
                            @{ email = $entraIdSupervisorEmail }
                          )
                          message            = "Access granted to $bhrDisplayName's OneDrive. Access expires in 30 days. Please copy any data you need to retain."
                          requireSignIn      = $true
                          sendInvitation     = $false
                          roles              = @('write')
                          expirationDateTime = $expirationString
                        }

                        Write-PSLog -Message "Granting OneDrive access to $entraIdSupervisorEmail for $bhrWorkEmail" -Severity Information
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($entraIdUpnObjDetails.Id)/drive/root/invite" -Body ($inviteBody | ConvertTo-Json -Depth 6) -ContentType 'application/json' | Out-Null

                        if (-not [string]::IsNullOrWhiteSpace($driveUrl)) {
                          $emailParams = @{
                            Message         = @{
                              Subject      = "OneDrive access for $bhrDisplayName"
                              Body         = @{
                                ContentType = 'html'
                                Content     = "<p>You have been granted access to $bhrDisplayName's OneDrive.</p><p><a href='$driveUrl'>$driveUrl</a></p><p>Access is available for up to 30 days (until $expirationString UTC). After that it will not be available. Please copy all data you would like to save.</p><p>$($Script:Config.Email.EmailSignature)</p>"
                              }
                              ToRecipients = @(
                                @{
                                  EmailAddress = @{
                                    Address = $entraIdSupervisorEmail
                                  }
                                }
                              )
                            }
                            SaveToSentItems = 'True'
                          }

                          Invoke-WithRetry -Operation "Send OneDrive access email to $entraIdSupervisorEmail" -ScriptBlock {
                            Send-MgUserMail -BodyParameter $emailParams -UserId $Script:Config.Email.AdminEmailAddress -Verbose
                          }
                          Write-PSLog -Message "OneDrive access email sent to $entraIdSupervisorEmail" -Severity Information
                        }
                        else {
                          Write-PSLog -Message "OneDrive URL not available for $bhrWorkEmail; access granted but no link to send" -Severity Warning
                        }
                      }
                      catch {
                        Write-PSLog -Message "Failed to grant OneDrive access for $($bhrWorkEmail): $($_.Exception.Message)" -Severity Warning
                      }
                    }
                    else {
                      Write-PSLog -Message 'OneDrive access grant requires Get-MgUserDrive and Invoke-MgGraphRequest; skipping' -Severity Warning
                    }
                  }
                  else {
                    Write-PSLog -Message "Manager email not available for $bhrWorkEmail; skipping OneDrive access grant" -Severity Warning
                  }

                  # Reset/wipe the employees device(s)
                  Write-PSLog -Message "Removing devices for $bhrWorkEmail..." -Severity Information
                  $uDevices = Get-MgUserOwnedDevice -UserId $bhrWorkEmail

                  Write-Output "User's devices"
                  $uDevices | Format-Table

                  $uDevices | ForEach-Object {
                    $deviceObject = $_
                    Write-PSLog -Message "Removing $bhrWorkEmail from $($deviceObject.Id) $($deviceObject.DisplayName)..." -Severity Debug
                    # Invoke-MgDeviceManagementManagedDeviceWindowsAutopilotReset -ManagedDeviceId $_.Id
                    Remove-MgDeviceRegisteredOwnerByRef -DeviceId $deviceObject.Id -DirectoryObjectId $userObjectId

                    $deviceDetails = Get-MgDevice -DeviceId $deviceObject.Id
                    $existingNotes = $deviceDetails.Notes
                    $timestamp = Get-Date -Format 'yyyy-MM-dd'
                    $updatedNotes = "$existingNotes | Owner $bhrWorkEmail removed on $timestamp"
                    Update-MgDevice -DeviceId $deviceObject.Id -BodyParameter @{ Notes = $updatedNotes }
                  }

                  Write-PSLog -Message "Evaluating Intune managed devices and Autopilot registrations for $bhrWorkEmail" -Severity Information
                  $managedDeviceCommand = Get-Command Get-MgUserManagedDevice -ErrorAction SilentlyContinue
                  if ($managedDeviceCommand) {
                    try {
                      $managedDevices = Get-MgUserManagedDevice -UserId $userObjectId -All -ErrorAction Stop
                      foreach ($managedDevice in $managedDevices) {
                        Write-PSLog -Message "Intune managed device should be reset: $($managedDevice.DeviceName) ($($managedDevice.Id))" -Severity Warning
                        $updateManagedDeviceCommand = Get-Command Update-MgDeviceManagementManagedDevice -ErrorAction SilentlyContinue
                        if ($updateManagedDeviceCommand) {
                          try {
                            $updateBody = @{ userId = $null; userPrincipalName = $null }
                            Update-MgDeviceManagementManagedDevice -ManagedDeviceId $managedDevice.Id -BodyParameter $updateBody -ErrorAction Stop
                            Write-PSLog -Message "Cleared Intune primary user for device $($managedDevice.DeviceName)" -Severity Information
                          }
                          catch {
                            Write-PSLog -Message "Failed to clear Intune primary user for $($managedDevice.DeviceName): $($_.Exception.Message)" -Severity Warning
                          }
                        }
                        else {
                          Write-PSLog -Message 'Update-MgDeviceManagementManagedDevice not available; cannot clear Intune primary user' -Severity Warning
                        }
                      }
                    }
                    catch {
                      Write-PSLog -Message "Failed to query Intune managed devices for $($bhrWorkEmail): $($_.Exception.Message)" -Severity Warning
                    }
                  }
                  else {
                    Write-PSLog -Message 'Get-MgUserManagedDevice not available; skipping Intune managed device processing' -Severity Warning
                  }

                  $autopilotCommand = Get-Command Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -ErrorAction SilentlyContinue
                  if ($autopilotCommand) {
                    try {
                      $autopilotDevices = Get-CachedAutopilotDevicesForUser -Cache $performanceCache -UserPrincipalName $bhrWorkEmail
                      foreach ($autopilotDevice in $autopilotDevices) {
                        Write-PSLog -Message "Autopilot device should be reset: $($autopilotDevice.DisplayName) ($($autopilotDevice.Id))" -Severity Warning
                        $updateAutopilotCommand = Get-Command Update-MgDeviceManagementWindowsAutopilotDeviceIdentity -ErrorAction SilentlyContinue
                        if ($updateAutopilotCommand) {
                          try {
                            $autopilotBody = @{ assignedUserPrincipalName = $null; userPrincipalName = $null }
                            Update-MgDeviceManagementWindowsAutopilotDeviceIdentity -WindowsAutopilotDeviceIdentityId $autopilotDevice.Id -BodyParameter $autopilotBody -ErrorAction Stop
                            Write-PSLog -Message "Cleared Autopilot assigned user for $($autopilotDevice.DisplayName)" -Severity Information
                          }
                          catch {
                            Write-PSLog -Message "Failed to clear Autopilot assigned user for $($autopilotDevice.DisplayName): $($_.Exception.Message)" -Severity Warning
                          }
                        }
                        else {
                          Write-PSLog -Message 'Update-MgDeviceManagementWindowsAutopilotDeviceIdentity not available; cannot clear Autopilot assigned user' -Severity Warning
                        }
                      }
                    }
                    catch {
                      Write-PSLog -Message "Failed to query Autopilot devices for $($bhrWorkEmail): $($_.Exception.Message)" -Severity Warning
                    }
                    else {
                      Write-PSLog -Message 'Get-MgDeviceManagementWindowsAutopilotDeviceIdentity not available; skipping Autopilot processing' -Severity Warning
                    }

                    # Remove Licenses
                    Write-PSLog -Message 'Removing licenses...' -Severity Information

                    Write-PSLog -Message "Executing: Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }" -Severity Debug
                    Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }

                    # Transfer group ownership to manager (if any)
                    Write-PSLog -Message "Transferring group ownership for $bhrWorkEmail to $entraIdSupervisorEmail" -Severity Information
                    $managerObj = Get-MgUser -UserId $entraIdSupervisorEmail -ErrorAction SilentlyContinue
                    if ($managerObj) {
                      $ownedGroups = Get-MgUserOwnedObject -UserId $bhrWorkEmail -All | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }
                      foreach ($group in $ownedGroups) {
                        $existingOwners = Get-MgGroupOwner -GroupId $group.Id -All | ForEach-Object { $_.Id }
                        if ($existingOwners -notcontains $managerObj.Id) {
                          Write-PSLog -Message "Adding manager as owner for group $($group.Id)" -Severity Debug
                          New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($managerObj.Id)" }
                        }
                        if ($existingOwners -contains $entraIdUpnObjDetails.Id) {
                          $removeOwnerCommand = Get-Command Remove-MgGroupOwnerByRef -ErrorAction SilentlyContinue
                          if ($removeOwnerCommand) {
                            try {
                              Write-PSLog -Message "Removing $bhrWorkEmail as owner for group $($group.Id)" -Severity Debug
                              Remove-MgGroupOwnerByRef -GroupId $group.Id -DirectoryObjectId $entraIdUpnObjDetails.Id -ErrorAction Stop
                            }
                            catch {
                              Write-PSLog -Message "Failed to remove $bhrWorkEmail as owner for group $($group.Id): $($_.Exception.Message)" -Severity Warning
                            }
                          }
                          else {
                            Write-PSLog -Message "Remove-MgGroupOwnerByRef not available; cannot remove $bhrWorkEmail as owner for group $($group.Id)" -Severity Warning
                          }
                        }
                      }
                    }
                    else {
                      Write-PSLog -Message "Manager $entraIdSupervisorEmail not found; skipping ownership transfer." -Severity Warning
                    }

                    # Remove groups
                    Write-PSLog -Message 'Removing group memberships' -Severity Debug
                    Write-PSLog -Message "Executing: Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object { Remove-MgGroupMemberByRef -GroupId $_.id -DirectoryObjectId $entraIdUpnObjDetails.id } " -Severity Debug

                    Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object {
                      $groupMembership = $_
                      try {
                        Invoke-WithRetry -Operation "Remove group membership $($groupMembership.Id) for: $bhrWorkEmail" -ScriptBlock {
                          Remove-MgGroupMemberByRef -GroupId $groupMembership.id -DirectoryObjectId $entraIdUpnObjDetails.id -ErrorAction Stop
                        }
                      }
                      catch {
                        Write-PSLog -Message "Failed to remove $bhrWorkEmail from Entra group $($groupMembership.Id): $($_.Exception.Message)" -Severity Warning
                      }
                    }

                    Write-PSLog -Message "Removing distribution list memberships for $bhrWorkEmail" -Severity Information
                    Connect-ExchangeOnlineIfNeeded -TenantId $Script:Config.Azure.TenantId
                    try {
                      $distributionGroups = Get-CachedDistributionGroups -Cache $performanceCache
                      foreach ($distributionGroup in $distributionGroups) {
                        try {
                          $members = Get-DistributionGroupMember -Identity $distributionGroup.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue |
                            Where-Object { $_.PrimarySmtpAddress -eq $bhrWorkEmail }
                            if ($members) {
                              Remove-DistributionGroupMember -Identity $distributionGroup.Identity -Member $bhrWorkEmail -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                              Write-PSLog -Message "Removed $bhrWorkEmail from distribution group $($distributionGroup.DisplayName)" -Severity Information
                            }
                          }
                          catch {
                            Write-PSLog -Message "Failed to remove $bhrWorkEmail from distribution group $($distributionGroup.DisplayName): $($_.Exception.Message)" -Severity Warning
                          }
                        }
                      }
                      catch {
                        Write-PSLog -Message "Failed to enumerate distribution groups: $($_.Exception.Message)" -Severity Warning
                      }

                      Write-PSLog -Message "Removing mailbox delegate permissions for $bhrWorkEmail" -Severity Information
                      try {
                        $mailboxes = Get-CachedMailboxes -Cache $performanceCache
                        foreach ($mailbox in $mailboxes) {
                          try {
                            $fullAccessPermissions = Get-EXOMailboxPermission -Identity $mailbox.Identity -User $bhrWorkEmail -ErrorAction SilentlyContinue |
                              Where-Object { $_.AccessRights -contains 'FullAccess' }
                              if ($fullAccessPermissions) {
                                Remove-MailboxPermission -Identity $mailbox.Identity -User $bhrWorkEmail -AccessRights 'FullAccess' -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                                Write-PSLog -Message "Removed FullAccess delegate from mailbox $($mailbox.Identity)" -Severity Information
                              }
                            }
                            catch {
                              Write-PSLog -Message "Failed to remove FullAccess delegate from mailbox $($mailbox.Identity): $($_.Exception.Message)" -Severity Warning
                            }
                          }
                        }
                        catch {
                          Write-PSLog -Message "Failed to enumerate mailboxes for delegate cleanup: $($_.Exception.Message)" -Severity Warning
                        }

                        try {
                          $sendAsPermissions = Get-EXORecipientPermission -Trustee $bhrWorkEmail -ResultSize Unlimited -ErrorAction SilentlyContinue |
                            Where-Object { $_.AccessRights -contains 'SendAs' }
                            foreach ($sendAsPermission in $sendAsPermissions) {
                              try {
                                Remove-RecipientPermission -Identity $sendAsPermission.Identity -Trustee $bhrWorkEmail -AccessRights 'SendAs' -Confirm:$false -ErrorAction Stop | Out-Null
                                Write-PSLog -Message "Removed SendAs delegate from mailbox $($sendAsPermission.Identity)" -Severity Information
                              }
                              catch {
                                Write-PSLog -Message "Failed to remove SendAs delegate from mailbox $($sendAsPermission.Identity): $($_.Exception.Message)" -Severity Warning
                              }
                            }
                          }
                          catch {
                            Write-PSLog -Message "Failed to enumerate SendAs permissions for $($bhrWorkEmail): $($_.Exception.Message)" -Severity Warning
                          }
                          $authMethods = @(
                            Invoke-WithRetry -Operation "Get authentication methods for: $bhrWorkEmail" -ScriptBlock {
                              Get-MgUserAuthenticationMethod -UserId $bhrWorkEmail -ErrorAction Stop
                            }
                          )

                          # Loop through and remove each authentication method
                          $error.Clear()

                          foreach ($authMethod in $authMethods) {
                            $methodData = $authMethod.AdditionalProperties.Values

                            if ($methodData -like '*phoneAuthenticationMethod*') { Remove-MgUserAuthenticationPhoneMethod -UserId $bhrWorkEmail -PhoneAuthenticationMethodId $authMethod.id; Write-PSLog -Message "Removed phone auth method for $bhrWorkEmail." -Severity Warning }
                            if ($methodData -like '*microsoftAuthenticatorAuthenticationMethod*') { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $bhrWorkEmail -MicrosoftAuthenticatorAuthenticationMethodId $authMethod.id; Write-PSLog -Message "Removed auth app method for $bhrWorkEmail." -Severity Warning }
                            if ($methodData -like '*fido2AuthenticationMethod*') { Remove-MgUserAuthenticationFido2Method -UserId $bhrWorkEmail -Fido2AuthenticationMethodId $authMethod.id; Write-PSLog -Message "Removed passkey (FIDO2) for $bhrWorkEmail." -Severity Warning }
                            if ($methodData -like '*windowsHelloForBusinessAuthenticationMethod*') { Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $bhrWorkEmail -WindowsHelloForBusinessAuthenticationMethodId $authMethod.id; Write-PSLog -Message "Removed Windows Hello for Business method for $bhrWorkEmail." -Severity Warning }
                          }

                          # Remove Manager
                          Write-PSLog -Message 'Removing Manager...' -Severity Debug
                          Write-PSLog -Message "Executing: Remove-MgUserManagerByRef -UserId $bhrWorkEmail" -Severity Debug
                          Remove-MgUserManagerByRef -UserId $bhrWorkEmail

                          Write-PSLog -Message "Executing: Update-MgUser -EmployeeId 'LVR' -UserId $bhrWorkEmail" -Severity Debug
                          Invoke-WithRetry -Operation "Mark user as on leave: $bhrWorkEmail" -ScriptBlock {
                            Update-MgUser -EmployeeId 'LVR' -UserId $bhrWorkEmail
                          }
                          Write-PSLog -Message 'Updating shared mailbox settings...' -Severity Information

                          if ($error.Count -ne 0) {
                            $error | ForEach-Object {
                              $err_Exception = $_.Exception
                              $err_Target = $_.TargetObject
                              $errCategory = $_.CategoryInfo
                              Write-PSLog " Could not remove authentication details. `n Exception: $err_Exception `n Target Object: $err_Target `n Error Category: $errCategory " -Severity Error
                            }
                          }
                          else {
                            # CompanyName is updated last — it acts as the completion marker.
                            # The incomplete-offboarding detector checks for this date stamp to
                            # decide whether offboarding finished successfully.
                            Write-PSLog -Message "Setting CompanyName completion stamp for $bhrWorkEmail" -Severity Information
                            Update-MgUser -UserId $bhrWorkEmail -CompanyName (Get-OffboardingCompletionMarker -LeaveDateTimeUtc $leaveDateTimeUtc)

                            Write-PSLog -Message " Account $bhrWorkEmail marked as inactive in BambooHR Entra ID account has been disabled, sessions revoked and removed MFA." -Severity Information
                            $error.Clear()
                          }
                        }
                      }
                      elseif ($bhrAccountEnabled -eq $false -and $bhrEmploymentStatus.Trim() -eq 'Terminated' -and $entraIdStatus -eq $false ) {
                        # Account is already disabled. Check whether the full offboarding process
                        # actually completed by inspecting profile fields that get overwritten during
                        # offboarding. If CompanyName does not start with a date stamp (MM/DD/YY),
                        # Department/JobTitle/OfficeLocation are not set to 'Not Active', or phone
                        # fields still contain values, the offboarding was interrupted or never ran —
                        # re-run it.
                        $offboardingComplete = Test-IsOffboardingComplete -CompanyName $entraIdCompanyName -Department $entraIdDepartment -JobTitle $entraIdJobTitle -OfficeLocation $entraIdOfficeLocation -WorkPhone $entraIdWorkPhone -MobilePhone $entraIdMobilePhone

                        if (-not $offboardingComplete) {
                          Write-PSLog -Message "$bhrWorkEmail is disabled but offboarding appears incomplete (Company='$entraIdCompanyName', Dept='$entraIdDepartment', Title='$entraIdJobTitle', Office='$entraIdOfficeLocation', WorkPhone='$entraIdWorkPhone', MobilePhone='$entraIdMobilePhone'). Re-running offboarding." -Severity Warning
                          # Re-run the same offboarding block that fires when the account transitions
                          # from enabled to disabled. We set $entraIdStatus to $true so the primary
                          # termination condition above will match on the next pass, but since we are
                          # already past that branch we duplicate the key termination actions here.

                          $error.clear()
                          if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Complete missed offboarding (Update Profile, Convert to Shared Mailbox, Remove Licenses and Groups)')) {
                            $leaveDateTimeUtc = (Get-Date).ToUniversalTime()

                            # Overwrite profile fields with termination placeholders (CompanyName set last as completion marker)
                            Write-PSLog -Message "Setting termination profile fields for $bhrWorkEmail" -Severity Information
                            Set-TerminatedUserProfileFields -UserId $bhrWorkEmail -LeaveDateTimeUtc $leaveDateTimeUtc

                            # Remove Licenses
                            Write-PSLog -Message "Removing licenses for $bhrWorkEmail..." -Severity Information
                            Get-MgUserLicenseDetail -UserId $bhrWorkEmail | ForEach-Object { Set-MgUserLicense -UserId $bhrWorkEmail -RemoveLicenses $_.SkuId -AddLicenses @{} }

                            # Remove group memberships
                            Write-PSLog -Message "Removing group memberships for $bhrWorkEmail" -Severity Information
                            Get-MgUserMemberOf -UserId $bhrWorkEmail | ForEach-Object {
                              $groupMembership = $_
                              try {
                                Invoke-WithRetry -Operation "Remove group membership $($groupMembership.Id) for: $bhrWorkEmail" -ScriptBlock {
                                  Remove-MgGroupMemberByRef -GroupId $groupMembership.id -DirectoryObjectId $entraIdUpnObjDetails.id -ErrorAction Stop
                                }
                              }
                              catch {
                                Write-PSLog -Message "Failed to remove $bhrWorkEmail from Entra group $($groupMembership.Id): $($_.Exception.Message)" -Severity Warning
                              }
                            }

                            # Remove distribution list memberships
                            Write-PSLog -Message "Removing distribution list memberships for $bhrWorkEmail" -Severity Information
                            Connect-ExchangeOnlineIfNeeded -TenantId $Script:Config.Azure.TenantId
                            try {
                              $distributionGroups = Get-CachedDistributionGroups -Cache $performanceCache
                              foreach ($distributionGroup in $distributionGroups) {
                                try {
                                  $members = Get-DistributionGroupMember -Identity $distributionGroup.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue |
                                    Where-Object { $_.PrimarySmtpAddress -eq $bhrWorkEmail }
                                    if ($members) {
                                      Remove-DistributionGroupMember -Identity $distributionGroup.Identity -Member $bhrWorkEmail -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                                      Write-PSLog -Message "Removed $bhrWorkEmail from distribution group $($distributionGroup.DisplayName)" -Severity Information
                                    }
                                  }
                                  catch {
                                    Write-PSLog -Message "Failed to remove $bhrWorkEmail from distribution group $($distributionGroup.DisplayName): $($_.Exception.Message)" -Severity Warning
                                  }
                                }
                              }
                              catch {
                                Write-PSLog -Message "Failed to enumerate distribution groups: $($_.Exception.Message)" -Severity Warning
                              }

                              # Convert mailbox to shared if not already
                              Write-PSLog -Message "Ensuring $bhrWorkEmail mailbox is shared" -Severity Information
                              try {
                                Set-Mailbox -Identity $bhrWorkEmail -Type Shared -ErrorAction Stop
                                $null = Wait-ForCondition -Operation "Mailbox conversion to shared for $bhrWorkEmail" -TimeoutSeconds 120 -PollIntervalSeconds 10 -Condition {
                                  (Get-EXOMailbox -Anr $bhrWorkEmail -ErrorAction Stop).RecipientTypeDetails -eq 'SharedMailbox'
                                }
                              }
                              catch {
                                Write-PSLog -Message "Mailbox conversion for $bhrWorkEmail may have already been done or failed: $($_.Exception.Message)" -Severity Warning
                              }

                              # Remove Manager
                              Write-PSLog -Message "Removing manager for $bhrWorkEmail" -Severity Information
                              Remove-MgUserManagerByRef -UserId $bhrWorkEmail -ErrorAction SilentlyContinue

                              # Mark EmployeeId as LVR
                              Invoke-WithRetry -Operation "Mark user as on leave: $bhrWorkEmail" -ScriptBlock {
                                Update-MgUser -EmployeeId 'LVR' -UserId $bhrWorkEmail
                              }

                              # CompanyName is updated last as the completion marker
                              Write-PSLog -Message "Setting CompanyName completion stamp for $bhrWorkEmail" -Severity Information
                              Update-MgUser -UserId $bhrWorkEmail -CompanyName (Get-OffboardingCompletionMarker -LeaveDateTimeUtc $leaveDateTimeUtc)

                              Add-SignificantChange -Category Disabled -User $bhrWorkEmail -Detail 'Offboarding completed (was incomplete)'
                              Write-PSLog -Message "Completed missed offboarding for $bhrWorkEmail" -Severity Information
                            }
                          }
                          else {
                            Write-PSLog -Message "$bhrWorkEmail is disabled and offboarding is complete. Nothing to do." -Severity Debug
                          }
                        }
                        else {
                          Write-PSLog 'User account active, looking for user updates.' -Severity Debug

                          if ($bhrAccountEnabled -eq $true -and $entraIdstatus -eq $false) {
                            # The account is marked "Active" in BHR and "Inactive" in Entra ID, enable the Entra ID account
                            Write-PSLog -Message "$bhrWorkEmail is marked Active in BHR and Inactive in Entra ID" -Severity Debug

                            #Change to a random pass
                            $newPas = (Get-NewPassword)
                            $params = @{
                              PasswordProfile = @{
                                ForceChangePasswordNextSignIn = $true
                                Password                      = $newPas
                              }
                            }

                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Re-enable User Account (Reset Password and Convert from Shared Mailbox)')) {
                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled" -Severity Debug
                              Invoke-WithRetry -Operation "Re-enable user account: $bhrWorkEmail" -ScriptBlock {
                                Update-MgUser -UserId $bhrWorkEmail -AccountEnabled:$bhrAccountEnabled
                              }
                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params" -Severity Debug
                              Invoke-WithRetry -Operation "Update re-enabled user attributes: $bhrWorkEmail" -ScriptBlock {
                                Update-MgUser -UserId $bhrWorkEmail -BodyParameter $params
                              }

                              # Convert mailbox from shared to user mailbox
                              Connect-ExchangeOnlineIfNeeded -TenantId $TenantId

                              Write-PSLog -Message "Executing: Set-Mailbox -Identity $bhrWorkEmail -Type Regular" -Severity Debug
                              Write-PSLog -Message "Converting $bhrWorkEmail to a user mailbox..." -Severity Debug
                              try {
                                Set-Mailbox -Identity $bhrWorkEmail -Type Regular -ErrorAction Stop
                              }
                              catch {
                                Write-PSLog -Message "Failed to convert $bhrWorkEmail to a user mailbox. Error: $($_.Exception.Message)" -Severity Warning
                              }

                              $null = Wait-ForCondition -Operation "Mailbox conversion to regular for $bhrWorkEmail" -TimeoutSeconds 120 -PollIntervalSeconds 10 -Condition {
                                (Get-EXOMailbox -Anr $bhrWorkEmail -ErrorAction Stop).RecipientTypeDetails -eq 'UserMailbox'
                              }

                              # Remove permissions to converted mailbox to previous manager
                              $mObj = Get-EXOMailbox -Anr $bhrWorkEmail
                              Write-PSLog "`tShared permissions being revoked for $bhrWorkEmail..." -Severity Information
                              Write-PSLog "Executing: Remove-MailboxPermission -Identity $($mObj.Id) -ResetDefault" -Severity Debug
                              try {
                                Remove-MailboxPermission -Identity $mObj.Id -ResetDefault -ErrorAction Stop | Out-Null
                              }
                              catch {
                                Write-PSLog -Message "Failed to remove mailbox permissions for $bhrWorkEmail. Error: $($_.Exception.Message)" -Severity Warning
                              }

                              # Remove automatic replies when the account is reactivated
                              Write-PSLog -Message "Removing automatic replies for $bhrWorkEmail" -Severity Information
                              Write-PSLog -Message "Executing: Set-MailboxAutoReplyConfiguration -Identity $bhrWorkEmail -AutoReplyState Disabled" -Severity Debug
                              try {
                                Invoke-WithRetry -Operation "Disable automatic replies for: $bhrWorkEmail" -ScriptBlock {
                                  Set-MailboxAutoReplyConfiguration -Identity $bhrWorkEmail `
                                    -AutoReplyState Disabled `
                                    -InternalMessage '' `
                                    -ExternalMessage '' `
                                    -ErrorAction Stop
                                }
                              }
                              catch {
                                Write-PSLog -Message "Failed to disable automatic replies for $bhrWorkEmail. Error: $($_.Exception.Message)" -Severity Warning
                              }

                              $params = @{
                                Message         = @{
                                  Subject      = "User Account Re-enabled: $bhrdisplayName"
                                  Body         = @{
                                    ContentType = 'html'
                                    Content     = "<br/>One of your direct report's user account has been re-enabled. Please securely share this information with them so that they can login.<br/> User name: $bhrWorkEmail <br/> Temporary Password: $newPas.`n<br/><br/> $($Script:Config.Email.EmailSignature)"
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
                                SaveToSentItems = 'True'
                              }

                              Invoke-WithRetry -Operation 'Send re-enable notification email' -ScriptBlock {
                                Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress -Verbose
                              }

                              New-AdaptiveCard {

                                New-AdaptiveTextBlock -Text "User Account $bhrWorkEmail Re-enabled" -HorizontalAlignment Center -Wrap -Weight Large
                                New-AdaptiveTextBlock -Text "User name: $bhrWorkEmail" -Wrap
                                New-AdaptiveTextBlock -Text "Temporary Password: $newPas" -Wrap
                              } -Uri $TeamsCardUri -Speak "User Account Re-enabled: $bhrdisplayName"


                              Write-PSLog -Message " Account $bhrWorkEmail marked as Active in BambooHR but Inactive in Entra ID. Enabled Entra ID account for sign-in." -Severity Information
                              $error.Clear()
                            }
                          }
                          else {
                            Write-PSLog -Message 'Account is in the correct state: Enabled in both BHR and Entra ID' -Severity Debug
                          }

                          # Checking JobTitle if correctly set, if not, configure the JobTitle as set in BambooHR
                          if ($entraIdJobTitle.Trim() -ne $bhrJobTitle.Trim()) {
                            Write-PSLog -Message "Entra ID Job Title $entraIdJobTitle does not match BHR Job Title $bhrJobTitle. Updating title." -Severity Debug

                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {

                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -JobTitle '$bhrJobTitle'" -Severity Debug
                              try {
                                Invoke-WithRetry -Operation "Update JobTitle for: $bhrWorkEmail" -ScriptBlock {
                                  if ([string]::IsNullOrWhiteSpace($bhrJobTitle) -eq $false) {
                                    Update-MgUser -UserId $bhrWorkEmail -JobTitle $bhrJobTitle
                                  }
                                  else {
                                    Update-MgUser -UserId $bhrWorkEmail -JobTitle $null
                                  }
                                }
                                $error.Clear()
                                Write-PSLog -Message "JobTitle for $bhrWorkEmail in Entra ID set from '$entraIdjobTitle' to '$bhrjobTitle'." -Severity Information
                              }
                              catch {
                                Write-PSLog -Message "Error changing Job Title of $bhrWorkEmail.`nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error

                                # Track error for summary report
                                $errorSummary.TotalErrors++
                                $errorType = 'AttributeUpdate_JobTitle'
                                if (-not $errorSummary.ErrorsByType.ContainsKey($errorType)) {
                                  $errorSummary.ErrorsByType[$errorType] = 0
                                }
                                $errorSummary.ErrorsByType[$errorType]++
                                if (-not $errorSummary.ErrorsByUser.ContainsKey($bhrWorkEmail)) {
                                  $errorSummary.ErrorsByUser[$bhrWorkEmail] = 'JobTitle update failed'
                                }

                                $error.Clear()
                              }
                            }
                          }

                          # Checking department if correctly set, if not, configure the Department as set in BambooHR
                          if ($entraIdDepartment.Trim() -ne $bhrDepartment.Trim()) {
                            Write-PSLog -Message "Entra ID department '$entraIdDepartment' does not match BambooHR department '$($bhrDepartment.Trim())'" -Severity Debug
                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -Department $bhrDepartment" -Severity Debug
                              try {
                                Invoke-WithRetry -Operation "Update Department for: $bhrWorkEmail" -ScriptBlock {
                                  Update-MgUser -UserId $bhrWorkEmail -Department "$bhrDepartment"
                                }
                                $error.Clear()
                                Write-PSLog -Message "Department for $bhrWorkEmail in Entra ID set from '$entraIdDepartment' to '$bhrDepartment'." -Severity Information
                              }
                              catch {
                                Write-PSLog -Message "Error changing Department of $bhrWorkEmail `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error

                                # Track error for summary report
                                $errorSummary.TotalErrors++
                                $errorType = 'AttributeUpdate_Department'
                                if (-not $errorSummary.ErrorsByType.ContainsKey($errorType)) {
                                  $errorSummary.ErrorsByType[$errorType] = 0
                                }
                                $errorSummary.ErrorsByType[$errorType]++
                                if (-not $errorSummary.ErrorsByUser.ContainsKey($bhrWorkEmail)) {
                                  $errorSummary.ErrorsByUser[$bhrWorkEmail] = 'Department update failed'
                                }

                                $error.Clear()
                              }
                            }
                          }
                          else {
                            Write-PSLog "Entra ID and BHR department already matches $entraIdDepartment" -Severity Debug
                          }

                          # Checking the manager if correctly set, and removing it when BambooHR clears the relationship.
                          $shouldRemoveManager = [string]::IsNullOrWhiteSpace($bhrSupervisorEmail) -and -not [string]::IsNullOrWhiteSpace($entraIdSupervisorEmail)
                          $shouldSetManager = -not [string]::IsNullOrWhiteSpace($bhrSupervisorEmail) -and $entraIdSupervisorEmail -ne $bhrSupervisorEmail
                          if ($shouldRemoveManager) {
                            Write-PSLog -Message "BambooHR manager is empty for $bhrWorkEmail; removing stale Entra ID manager '$entraIdSupervisorEmail'." -Severity Debug
                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Remove User Manager')) {
                              try {
                                Invoke-WithRetry -Operation "Remove manager for: $bhrWorkEmail" -ScriptBlock {
                                  Remove-MgUserManagerByRef -UserId $bhrWorkEmail -ErrorAction Stop
                                }
                                $error.Clear()
                                Write-PSLog -Message "Removed manager '$entraIdSupervisorEmail' from $bhrWorkEmail because BambooHR no longer has a supervisor assigned." -Severity Information
                                Add-SignificantChange -Category ManagerChanged -User $bhrWorkEmail -Detail "$entraIdSupervisorEmail -> <none>"
                              }
                              catch {
                                Write-PSLog -Message "Error removing manager of $bhrWorkEmail. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                $errorSummary.TotalErrors++
                                $errorType = 'ManagerRemoval'
                                if (-not $errorSummary.ErrorsByType.ContainsKey($errorType)) {
                                  $errorSummary.ErrorsByType[$errorType] = 0
                                }
                                $errorSummary.ErrorsByType[$errorType]++
                                $errorSummary.ErrorsByUser[$bhrWorkEmail] = "Failed to remove manager: $($_.Exception.Message)"
                                $error.Clear()
                              }
                            }
                          }
                          elseif ($shouldSetManager) {
                            Write-PSLog -Message "Manager in Entra ID '$entraIdSupervisorEmail' does not match BHR manager '$bhrSupervisorEmail'" -Severity Debug

                            $managerUser = Get-ValidManagerUser -UserPrincipalName $bhrSupervisorEmail -Cache $performanceCache -TargetUser $bhrWorkEmail

                            if ($managerUser) {
                              $newManager = @{
                                '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($managerUser.Id)"
                              }

                              if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                                Write-PSLog -Message "Executing: Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $newManager" -Severity Debug
                                try {
                                  Invoke-WithRetry -Operation "Set manager for: $bhrWorkEmail" -ScriptBlock {
                                    Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $newManager
                                  }
                                  $error.Clear()
                                  Write-PSLog -Message "Manager of $bhrWorkEmail in Entra ID '$entraIdsupervisorEmail' and in BambooHR '$bhrsupervisorEmail'. Setting new manager to the Azure User Object." -Severity Information
                                  Add-SignificantChange -Category ManagerChanged -User $bhrWorkEmail -Detail "$entraIdSupervisorEmail -> $bhrSupervisorEmail"
                                }
                                catch {
                                  Write-PSLog -Message "Error changing manager of $bhrWorkEmail. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error

                                  # Track error for summary report
                                  $errorSummary.TotalErrors++
                                  $errorType = 'ManagerAssignment'
                                  if (-not $errorSummary.ErrorsByType.ContainsKey($errorType)) {
                                    $errorSummary.ErrorsByType[$errorType] = 0
                                  }
                                  $errorSummary.ErrorsByType[$errorType]++
                                  $errorSummary.ErrorsByUser[$bhrWorkEmail] = "Failed to set manager: $($_.Exception.Message)"

                                  $error.Clear()
                                }
                              }
                            }
                          }
                          else {
                            Write-PSLog -Message "Supervisor email already correct $entraIdSuperVisorEmail" -Severity Debug
                          }

                          # Check and set the Office Location
                          if ($entraIdOfficeLocation.Trim() -ne $bhrOfficeLocation.Trim()) {
                            Write-PSLog -Message "Entra ID office location '$entraIdOfficeLocation' does not match BHR hire data '$bhrOfficeLocation'" -Severity Debug
                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -OfficeLocation $($bhrOfficeLocation.Trim())" -Severity Debug
                              try {
                                Invoke-WithRetry -Operation "Update OfficeLocation for: $bhrWorkEmail" -ScriptBlock {
                                  Update-MgUser -UserId $bhrWorkEmail -OfficeLocation $bhrOfficeLocation.Trim()
                                }
                                $error.Clear()
                                Write-PSLog -Message "Office location of $bhrWorkEmail in Entra ID changed from '$entraIdOfficeLocation' to '$bhrOfficeLocation'." -Severity Information
                              }
                              catch {
                                Write-PSLog -Message "Error changing employee office location. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                              }
                            }
                          }
                          else {
                            Write-PSLog -Message "Office Location correct $entraIdOfficeLocation" -Severity Debug
                          }

                          # Check and set the Employee Hire Date
                          if ($entraIdHireDate -ne $normalizedBhrHireDate) {
                            Write-PSLog -Message "Entra ID hire date '$entraIdHireDate' does not match BHR hire data '$normalizedBhrHireDate'" -Severity Debug
                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeHireDate $normalizedBhrHireDate" -Severity Debug
                              try {
                                Invoke-WithRetry -Operation "Update EmployeeHireDate for: $bhrWorkEmail" -ScriptBlock {
                                  Update-MgUser -UserId $bhrWorkEmail -EmployeeHireDate $normalizedBhrHireDate
                                }
                                $error.Clear()
                                Write-PSLog -Message "Hire date of $bhrWorkEmail changed from '$entraIdHireDate' in Entra ID and BHR '$normalizedBhrHireDate'." -Severity Information
                              }
                              catch {
                                Write-PSLog -Message "Error changing $bhrWorkEmail hire date. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                              }
                            }
                          }
                          else {
                            Write-PSLog -Message "Hire date already correct $entraIdHireDate" -Severity Debug
                          }

                          # Check and set the work phone ignoring formatting
                          # Treat empty/whitespace and '0' as equivalent (both mean "no phone")
                          $normalizedBhrWork = Get-WorkPhoneComparisonValue -PhoneNumber $bhrWorkPhone
                          $normalizedEntraWork = Get-WorkPhoneComparisonValue -PhoneNumber $entraIdWorkPhone
                          if ($normalizedEntraWork -ne $normalizedBhrWork) {

                            Write-PSLog -Message "Entra ID work phone '$entraIdWorkPhone' does not match BHR '$bhrWorkPhone'" -Severity Debug
                            if ([string]::IsNullOrWhiteSpace($bhrWorkPhone)) {
                              $bhrWorkPhone = '0'
                            }

                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                              $workPhoneUpdated = $false
                              if ([string]::IsNullOrWhiteSpace($bhrWorkPhone)) {
                                $bhrWorkPhone = '0'
                                Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones '$bhrWorkPhone'" -Severity Debug
                                try {
                                  Invoke-WithRetry -Operation "Update BusinessPhones for: $bhrWorkEmail" -ScriptBlock {
                                    Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone -ErrorAction Stop | Out-Null
                                  }
                                  $workPhoneUpdated = $true
                                }
                                catch {
                                  Write-PSLog -Message "Error changing work phone for $bhrWorkEmail. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                  $error.Clear()
                                }
                              }
                              else {
                                [string]$bhrWorkPhone = ConvertTo-PhoneNumber $bhrWorkPhone
                                Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone" -Severity Debug
                                try {
                                  Invoke-WithRetry -Operation "Update BusinessPhones for: $bhrWorkEmail" -ScriptBlock {
                                    Update-MgUser -UserId $bhrWorkEmail -BusinessPhones $bhrWorkPhone -ErrorAction Stop | Out-Null
                                  }
                                  $workPhoneUpdated = $true
                                }
                                catch {
                                  Write-PSLog -Message "Error changing work phone for $bhrWorkEmail. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                  $error.Clear()
                                }
                              }

                              if ($workPhoneUpdated) {
                                $error.Clear()
                                Write-PSLog -Message "Work Phone for '$bhrWorkEmail' changed from '$entraIdWorkPhone' to '$bhrWorkPhone'" -Severity Information
                              }
                            }
                          }
                          else {
                            Write-PSLog -Message "Work phone correct $entraIdWorkEmail $entraIdWorkPhone" -Severity Debug
                          }

                          if ($Script:Config.Features.EnableMobilePhoneSync) {
                            [string]$entraIdMobilePhone = $entraIdMobilePhone -replace '[^0-9]', ''
                            [string]$bhrMobilePhone = $bhrMobilePhone -replace '[^0-9]', ''
                            # Check and set the mobile phone ignoring formatting
                            if ($entraIdMobilePhone -ne $bhrMobilePhone) {

                              Write-PSLog -Message "Entra ID mobile phone '$entraIdMobilePhone' does not match BHR '$bhrMobilePhone'" -Severity Debug

                              if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                                $mobilePhoneUpdated = $false
                                if ([string]::IsNullOrWhiteSpace($bhrMobilePhone)) {
                                  $bhrMobilePhone = '0'
                                  Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -MobilePhone '$bhrMobilePhone'" -Severity Debug
                                  try {
                                    Invoke-WithRetry -Operation "Update MobilePhone for: $bhrWorkEmail" -ScriptBlock {
                                      Update-MgUser -UserId $bhrWorkEmail -MobilePhone $bhrMobilePhone -ErrorAction Stop
                                    }
                                    $mobilePhoneUpdated = $true
                                  }
                                  catch {
                                    Write-PSLog -Message "Error changing $bhrWorkEmail mobile phone. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                  }
                                }
                                else {
                                  $bhrMobilePhone = ($bhrMobilePhone -replace '[^0-9]', '' ) -replace '^1', ''
                                  $bhrMobilePhone = '{0:(###) ###-####}' -f [int64]$bhrMobilePhone
                                  Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -MobilePhone $bhrMobilePhone" -Severity Debug
                                  try {
                                    Invoke-WithRetry -Operation "Update MobilePhone for: $bhrWorkEmail" -ScriptBlock {
                                      Update-MgUser -UserId $bhrWorkEmail -MobilePhone $bhrMobilePhone -ErrorAction Stop
                                    }
                                    $mobilePhoneUpdated = $true
                                  }
                                  catch {
                                    Write-PSLog -Message "Error changing $bhrWorkEmail mobile phone. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                    $error.Clear()
                                  }
                                }

                                if ($mobilePhoneUpdated) {
                                  $error.Clear()
                                  Write-PSLog -Message "Work Mobile Phone for '$bhrWorkEmail' changed from '$entraIdMobilePhone' to '$bhrMobilePhone'" -Severity Debug
                                }
                              }
                            }
                            else {
                              Write-PSLog -Message "Mobile phone correct for $entraIdWorkEmail $entraIdMobilePhone" -Severity Debug
                            }
                          }

                          # Compare user employee id with BambooHR and set it if not correct
                          if ($bhrEmployeeNumber.Trim() -ne $entraIdEmployeeNumber.Trim()) {
                            Write-PSLog -Message " BHR employee number $bhrEmployeeNumber does not match Entra ID employee id $entraIdEmployeeNumber" -Severity Debug
                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber  "
                              # Setting the Employee ID found in BHR to the user in Entra ID
                              try {
                                Invoke-WithRetry -Operation "Update EmployeeId for: $bhrWorkEmail" -ScriptBlock {
                                  Update-MgUser -UserId $bhrWorkEmail -EmployeeId $bhremployeeNumber.Trim() -ErrorAction Stop
                                }
                                Write-PSLog -Message " The ID $bhremployeeNumber has been set to $bhrWorkEmail Entra ID account." -Severity Warning
                                $error.Clear()
                              }
                              catch {
                                Write-PSLog -Message " Error changing EmployeeId. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                              }
                            }
                          }
                          else {
                            Write-PSLog -Message "Employee ID matched $bhrEmployeeNumber and $entraIdEmployeeNumber" -Severity Debug
                          }

                          # Set Company name to $($Script:Config.Azure.CompanyName)"
                          if ($entraIdCompanyName.Trim() -ne $Script:Config.Azure.CompanyName.Trim()) {
                            Write-PSLog -Message "Entra ID company name '$entraIdCompany' does not match '$($Script:Config.Azure.CompanyName)'" -Severity Debug
                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                              # Setting Company Name as $CompanyName to the employee, if not already set
                              Write-PSLog -Message "Executing: Update-MgUser -UserId $bhrWorkEmail -CompanyName $($CompanyName.Trim())" -Severity Debug
                              try {
                                Invoke-WithRetry -Operation "Update CompanyName for: $bhrWorkEmail" -ScriptBlock {
                                  Update-MgUser -UserId $bhrWorkEmail -CompanyName $CompanyName.Trim() -ErrorAction Stop
                                }
                                Write-PSLog -Message " The $bhrWorkEmail employee Company attribute has been set to: $($Script:Config.Azure.CompanyName)." -Severity Information
                              }
                              catch {
                                Write-PSLog -Message " Could not change the Company Name of $bhrWorkEmail. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                                $error.Clear()
                              }
                            }
                          }
                          else {
                            Write-PSLog -Message "Company name already matched in Entra ID and BHR $entraIdCompanyName" -Severity Debug
                          }

                          # Set LastModified from BambooHR to ExtensionAttribute1 in Entra ID

                          if ($upnExtensionAttribute1 -ne $bhrLastChanged) {
                            # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in Entra ID
                            Write-PSLog -Message "Entra ID Extension Attribute '$upnExtensionAttribute1' does not match BHR last changed '$bhrLastChanged'" -Severity Debug
                            Write-PSLog -Message 'Set LastModified from BambooHR to ExtensionAttribute1 in Entra ID' -Severity Debug

                            if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                              Write-PSLog -Message "Executing: $null = Update-MgUser -UserId $bhrWorkEmail -OnPremisesExtensionAttributes @{extensionAttribute1 = '$bhrLastChanged' }" -Severity Debug
                              # TODO: Does not work for on premises synched accounts. Not a problem with Entra ID native.
                              try {
                                $null = Update-MgUser -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrLastChanged } -UserId $bhrWorkEmail -ErrorAction Stop | Out-Null
                                Write-PSLog -Message "$bhrWorkEmail LastChanged attribute set from '$upnExtensionAttribute1' to '$bhrlastChanged'." -Severity Information
                              }
                              catch {
                                $error.Clear()
                              }
                            }

                            $error.clear()
                          }
                          else {
                            Write-PSLog -Message "Attribute already matched last changed of $bhrLastChanged" -Severity Debug
                          }
                        }
                      }
                    }
                    else {
                      Write-PSLog -Message "Entra ID user found for $bhrWorkEmail but no attribute sync needed (last changed dates match or user is suspended)" -Severity Debug

                      # This might not be needed anymore
                      $entraIdWorkEmail = ''
                      $entraIdJobTitle = ''
                      $entraIdDepartment = ''
                      $entraIdStatus = $false
                      $entraIdEmployeeNumber = ''
                      $entraIdSupervisorEmail = ''
                      $entraIdDisplayname = ''
                      $entraIdHireDate = ''
                      $entraIdFirstName = ''
                      $entraIdLastName = ''
                      $entraIdCompanyName = ''
                      $entraIdWorkPhone = ''
                      $entraIdOfficeLocation = ''
                    }

                    # Handle name changes
                    if (($entraIdEmployeeNumber2 -eq $bhremployeeNumber) -and ($historystatus -notlike '*inactive*') -and ($entraIdUpnObjDetails.id -eq $entraIdEidObjDetails.id)) {

                      $entraIdUPN = $entraIdEidObjDetails.UserPrincipalName
                      $entraIdObjectID = $entraIdEidObjDetails.id
                      $entraIdworkemail = $entraIdEidObjDetails.Mail
                      $entraIdemployeeNumber = $entraIdEidObjDetails.EmployeeID
                      $entraIddisplayname = $entraIdEidObjDetails.displayname
                      $entraIdfirstName = $entraIdEidObjDetails.GivenName
                      $entraIdlastName = $entraIdEidObjDetails.Surname

                      Write-PSLog -Message "Evaluating if Entra ID name change is required for $entraIdfirstName $entraIdlastName ($entraIddisplayname) `n`t Work Email: $entraIdWorkEmail UserPrincipalName: $entraIdUpn EmployeeId: $entraIdEmployeeNumber" -Severity Debug

                      $error.Clear()

                      # 3/31/2023 Is this required here or should it be handled after the name change or the next sync after the name change?
                      # Set LastModified from BambooHR to ExtensionAttribute1 in Entra ID
                      if ($EIDExtensionAttribute1 -ne $bhrlastChanged) {
                        if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {

                          # Setting the "lastchanged" attribute from BambooHR to ExtensionAttribute1 in Entra ID
                          Write-PSLog -Message "Executing: Update-MgUser -UserId $entraIdObjectID -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } " -Severity Debug
                          # This does not work for AD on premises synced accounts.
                          $null = Update-MgUser -UserId $entraIdObjectID -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged } -ErrorAction SilentlyContinue | Out-Null
                        }
                      }

                      # Change last name in Entra ID
                      if ($entraIdLastName -ne $bhrLastName) {
                        Write-PSLog -Message " Last name in Entra ID $entraIdLastName does not match in BHR $bhrLastName" -Severity Debug
                        Write-PSLog -Message " Changing the last name of $bhrWorkEmail from $entraIdLastName to $bhrLastName." -Severity Debug
                        if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                          Write-PSLog -Message "Executing: Update-MgUser -UserId $entraIdObjectID -Surname $bhrLastName" -Severity Debug
                          try {
                            Invoke-WithRetry -Operation "Update surname for: $bhrWorkEmail" -ScriptBlock {
                              Update-MgUser -UserId $entraIdObjectID -Surname $bhrLastName -ErrorAction Stop
                            }
                            Write-PSLog -Message " Successfully changed the last name of $bhrWorkEmail from $entraIdLastName to $bhrLastName." -Severity Information
                            Add-SignificantChange -Category NameChanged -User $bhrWorkEmail -Detail "Name: '$entraIdDisplayName' -> '$bhrDisplayName'"
                            Write-PSLog -Message "Name change: $bhrWorkEmail" -Severity Information
                          }
                          catch {
                            Write-PSLog -Message "Error changing Entra ID Last Name.`n`nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                          }
                        }
                      }

                      # Change First Name in Entra ID
                      if ($entraIdfirstName -ne $bhrfirstName) {
                        Write-PSLog "Entra ID first name '$entraIdfirstName' is not equal to BHR first name '$bhrFirstName'" -Severity Debug
                        if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                          Write-PSLog -Message "Executing: Update-MgUser -UserId $entraIdObjectID -GivenName $bhrFirstName" -Severity Debug
                          try {
                            Invoke-WithRetry -Operation "Update given name for: $bhrWorkEmail" -ScriptBlock {
                              Update-MgUser -UserId $entraIdObjectID -GivenName $bhrFirstName -ErrorAction Stop
                            }
                            Write-PSLog -Message " Successfully changed $entraIdObjectID first name from $entraIdFirstName to $bhrFirstName." -Severity Information
                            Add-SignificantChange -Category NameChanged -User $bhrWorkEmail -Detail "Name: '$entraIdDisplayName' -> '$bhrDisplayName'"
                            Write-PSLog -Message "Name change: $bhrWorkEmail" -Severity Information
                          }
                          catch {
                            Write-PSLog -Message "Could not change the First Name of $entraIdObjectID. Error details below. `n`nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                          }
                        }
                      }

                      # Change display name
                      if ($entraIdDisplayname -ne $bhrDisplayName) {
                        Write-PSLog -Message "Entra ID Display Name $entraIdDisplayname is not equal to BHR $bhrDisplayName" -Severity Debug
                        if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                          Write-PSLog -Message "Executing: Update-MgUser -UserId $entraIdObjectID -DisplayName $bhrdisplayName" -Severity Debug
                          try {
                            Invoke-WithRetry -Operation "Update display name for: $bhrWorkEmail" -ScriptBlock {
                              Update-MgUser -UserId $entraIdObjectID -DisplayName $bhrdisplayName -ErrorAction Stop
                            }
                            Write-PSLog " Display name $entraIdDisplayName of $entraIdObjectID changed to $bhrDisplayName." -Severity Information
                            Add-SignificantChange -Category NameChanged -User $bhrWorkEmail -Detail "Name: '$entraIdDisplayName' -> '$bhrDisplayName'"
                            Write-PSLog -Message "Name change: $bhrWorkEmail" -Severity Information
                          }
                          catch {
                            Write-PSLog -Message " Could not change the Display Name. Error details below. `n`nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                          }
                        }
                      }

                      # Change Email Address
                      if ($entraIdWorkEmail -ne $bhrWorkEmail) {
                        Write-PSLog -Message "Entra ID work email $entraIdWorkEmail does not match BHR work email $bhrWorkEmail"
                        if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                          Write-PSLog -Message "Executing: Update-MgUser -UserId $entraIdObjectID -Mail $bhrWorkEmail"
                          try {
                            Invoke-WithRetry -Operation "Update mail for: $bhrWorkEmail" -ScriptBlock {
                              Update-MgUser -UserId $entraIdObjectID -Mail $bhrWorkEmail -ErrorAction Stop
                            }
                            # Change Email Address error logging
                            Write-PSLog "The current Email Address: $entraIdworkemail of $entraIdObjectID has been changed to $bhrWorkEmail." -Severity Warning
                          }
                          catch {
                            Write-PSLog -Message "Error changing Email Address. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                          }
                        }
                      }

                      # Change UserPrincipalName and send the details via email to the User
                      if ($entraIdUpn -ne $bhrWorkEmail) {
                        Write-PSLog -Message "aadUPN $entraIdUpn does not match bhrWorkEmail $bhrWorkEmail" -Severity Debug
                        if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Update User')) {
                          Write-PSLog -Message "Executing: Update-MgUser -UserId $entraIdObjectID -UserPrincipalName $bhrWorkEmail" -Severity Debug
                          try {
                            Invoke-WithRetry -Operation "Update UPN for: $bhrWorkEmail" -ScriptBlock {
                              Update-MgUser -UserId $entraIdObjectID -UserPrincipalName $bhrWorkEmail -ErrorAction Stop
                            }
                            Write-PSLog -Message " Changed the current UPN:$entraIdUPN of $entraIdObjectID to $bhrWorkEmail." -Severity Warning
                            Add-SignificantChange -Category UpnChanged -User $bhrWorkEmail -Detail "$entraIdUPN -> $bhrWorkEmail"
                            Write-PSLog -Message "UPN change: $entraIdUPN -> $bhrWorkEmail" -Severity Information
                            $params = @{
                              Message         = @{
                                Subject       = "Login changed for $bhrdisplayName"
                                Body          = @{
                                  ContentType = 'HTML'
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
                              SaveToSentItems = 'True'
                            }

                            Invoke-WithRetry -Operation 'Send email address change notification' -ScriptBlock {
                              Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress -Verbose
                            }

                            New-AdaptiveCard {

                              New-AdaptiveTextBlock -Text "Login changed for $bhrdisplayName" -HorizontalAlignment Center -Weight Bolder -Wrap
                              New-AdaptiveTextBlock -Text "An email address was changed in the $($Script:Config.Azure.CompanyName) BambooHR. Your user account has been changed accordingly." -Wrap
                              New-AdaptiveTextBlock -Text "The user should use the new user name: $bhrWorkEmail" -Wrap
                              New-AdaptiveTextBlock -Text "The user's password has not been modified." -Wrap
                            } -Uri $TeamsCardUri -Speak "Login changed for $bhrdisplayName"
                          }
                          catch {
                            Write-PSLog -Message " Error changing UPN for $entraIdObjectID. `n Exception: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Error
                            $error.Clear()
                          }
                        }
                      }
                    }

                    # Create new employee account
                    if ((-not $entraIdUpnObjDetails) -and (-not $entraIdEidObjDetails) -and ($bhrAccountEnabled -eq $true)) {
                      Write-PSLog -Message "No Entra ID account exist but employee in bhr is $bhrAccountEnabled" -Severity Debug

                      if ([string]::IsNullOrWhiteSpace($Script:Config.Azure.LicenseId) -eq $false) {

                        Get-LicenseStatus -LicenseId $Script:Config.Azure.LicenseId -NewUser
                      }

                      $PasswordProfile = @{
                        Password = (Get-NewPassword)
                      }

                      $error.clear()

                      if ($PSCmdlet.ShouldProcess($bhrWorkEmail, 'Create New User Account')) {
                        # Create Entra ID account, as it doesn't have one, if user hire date is less than $DaysAhead days in the future, or is in the past
                        Write-PSLog -Message "$bhrWorkEmail does not have an Entra ID account and hire date ($normalizedBhrHireDate) is less than $($Script:Config.Features.DaysAhead) days from now." -Severity Information

                        Write-PSLog -Message "Executing New-MgUser -EmployeeId $bhremployeeNumber -Department $bhrDepartment -CompanyName $($Script:Config.Azure.CompanyName) -Surname $bhrlastName -GivenName $bhrfirstName -DisplayName $bhrdisplayName -AccountEnabled -Mail $bhrWorkEmail -OfficeLocation $bhrOfficeLocation `
                        -EmployeeHireDate $normalizedBhrHireDate -UserPrincipalName $bhrWorkEmail -PasswordProfile $PasswordProfile -JobTitle $bhrjobTitle -MailNickname $(Get-MailNicknameFromEmail -EmailAddress $bhrWorkEmail) -UsageLocation $($Script:Config.Azure.UsageLocation) -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged }" -Severity Debug

                        $user = Invoke-WithRetry -Operation "Create new user: $bhrWorkEmail" -ScriptBlock {
                          New-MgUser -EmployeeId $bhrEmployeeNumber -Department $bhrDepartment -CompanyName $Script:Config.Azure.CompanyName -Surname $bhrlastName -GivenName $bhrfirstName -DisplayName $bhrdisplayName `
                            -AccountEnabled -Mail $bhrWorkEmail -OfficeLocation $bhrOfficeLocation -EmployeeHireDate $normalizedBhrHireDate -UserPrincipalName $bhrWorkEmail -PasswordProfile $PasswordProfile `
                            -JobTitle $bhrjobTitle -MailNickname (Get-MailNicknameFromEmail -EmailAddress $bhrWorkEmail) `
                            -UsageLocation $UsageLocation -OnPremisesExtensionAttributes @{extensionAttribute1 = $bhrlastChanged }
                        }

                        # Did the account get created?
                        if ($null -eq $user) {
                          Write-PSLog -Message "Error creating Entra ID account for $bhrWorkEmail. `nException: $($Error.exception) `nTarget object: $($error.TargetObject) `nDetails: $($error.ErrorDetails) `nStackTrace: $($error.ScriptStackTrace)" -Severity Error
                          Write-PSLog -Message "Account $bhrWorkEmail creation failed. New-Mguser cmdlet returned error. `n $($error | Select-Object *)"

                          # Track critical error for summary report
                          $errorSummary.TotalErrors++
                          $errorSummary.CriticalErrors += "User creation failed: $bhrWorkEmail - $($Error.exception.Message)"
                          $errorType = 'UserCreation'
                          if (-not $errorSummary.ErrorsByType.ContainsKey($errorType)) {
                            $errorSummary.ErrorsByType[$errorType] = 0
                          }
                          $errorSummary.ErrorsByType[$errorType]++
                          $errorSummary.ErrorsByUser[$bhrWorkEmail] = "Failed to create user: $($Error.exception.Message)"

                          $params = @{
                            Message         = @{
                              Subject      = "BhrEntraSync error: User creation automation $bhrdisplayName"
                              Body         = @{
                                ContentType = 'html'
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
                            SaveToSentItems = 'True'
                          }

                          # Send Mail Message parameters definition closure
                          Invoke-WithRetry -Operation 'Send user creation error notification' -ScriptBlock {
                            Send-MgUserMail -BodyParameter $params -UserId $AdminEmailAddress -Verbose
                          }

                          New-AdaptiveCard {

                            New-AdaptiveTextBlock -Text "Account creation for user: $bhrWorkEmail failed." -HorizontalAlignment Center -Weight Bolder -Wrap
                            New-AdaptiveTextBlock -Text "Error Message: $($error.Exception.Message)" -Wrap
                            New-AdaptiveTextBlock -Text "Error Category: $($error.CategoryInfo)" -Wrap
                            New-AdaptiveTextBlock -Text "Error ID: $($error.FullyQualifiedErrorId)" -Wrap
                            New-AdaptiveTextBlock -Text "Stack: $($error.ScriptStackTrace)" -Wrap
                          } -Uri $TeamsCardUri -Speak 'BHR-Sync Account Creation Error'
                        }
                        else {
                          Write-PSLog -Message "Entra ID account for $bhrWorkEmail created." -Severity Information
                          Write-PSLog -Message "Created new user: $bhrWorkEmail" -Severity Information
                          Add-SignificantChange -Category Created -User $bhrWorkEmail -Detail $bhrDisplayName

                          # Since we are setting up a new account lets use the image from the BambooHR profile and
                          # add it to the Entra ID account
                          Write-PSLog -Message 'Retrieving user photo from BambooHR...' -Severity Information
                          $bhrEmployeePhotoUri = "$($bhrRootUri)/employees/$bhrEmployeeId/photo/large"
                          $profilePicPath = Join-Path -Path $env:temp -ChildPath "bhr-$($bhrEmployeeId).jpg"
                          Remove-Item -Path $profilePicPath -ErrorAction SilentlyContinue -Force | Out-Null
                          Write-PSLog -Message "Executing: Invoke-RestMethod -Uri $bhrRep -Method POST -Headers $headers -ContentType 'application/json' -OutFile $profilePicPath" -Severity Debug
                          $null = Invoke-RestMethod -Uri $bhrEmployeePhotoUri -Method GET -Headers $headers -ContentType 'application/json' -OutFile $profilePicPath -ErrorAction SilentlyContinue | Out-Null

                          Write-PSLog 'Updating user account with BambooHR profile picture...' -Severity Information
                          $photoUploadUser = $null
                          if ((Test-Path $profilePicPath -PathType Leaf -ErrorAction SilentlyContinue) -eq
                            $false -and (Test-Path $DefaultProfilePicPath)) {
                            $profilePicPath = $DefaultProfilePicPath
                          }

                          if (Test-Path $profilePicPath -PathType Leaf -ErrorAction SilentlyContinue) {
                            $userAvailableForPhoto = Wait-ForCondition -Operation "Wait for newly created user $bhrWorkEmail before photo upload" -TimeoutSeconds $Script:Config.Runtime.OperationTimeoutSeconds -PollIntervalSeconds $Script:Config.Runtime.RetryDelaySeconds -Condition {
                              $photoUploadUser = Get-MgUser -UserId $bhrWorkEmail -ErrorAction Stop
                              return ($null -ne $photoUploadUser) -and (-not [string]::IsNullOrWhiteSpace($photoUploadUser.Id))
                            }

                            if ($userAvailableForPhoto) {
                              Write-PSLog "Executing: Set-MgUserPhotoContent -UserId $($photoUploadUser.Id) -InFile $profilePicPath" -Severity Debug
                              try {
                                Invoke-WithRetry -Operation "Set user photo for: $bhrWorkEmail" -ScriptBlock {
                                  Set-MgUserPhotoContent -UserId $photoUploadUser.Id -InFile $profilePicPath -ErrorAction Stop
                                }
                              }
                              catch {
                                Write-PSLog -Message "Unable to update profile photo for $bhrWorkEmail. `nException: $($_.Exception) `nTarget object: $($_.TargetObject) `nDetails: $($_.ErrorDetails) `nStackTrace: $($_.ScriptStackTrace)" -Severity Warning
                              }
                            }
                            else {
                              Write-PSLog -Message "Skipping profile photo upload for $bhrWorkEmail because the new user was not available within the configured timeout." -Severity Warning
                            }
                          }

                          if ($profilePicPath -ne $DefaultProfilePicPath) {
                            Write-PSLog "Executing: Remove-Item -Path $profilePicPath -ErrorAction SilentlyContinue | Out-Null" -Severity Debug
                            #Remove-Item -Path $profilePicPath -ErrorAction SilentlyContinue -Force | Out-Null
                          }

                          if ([string]::IsNullOrWhiteSpace($bhrSupervisorEmail) -eq $false) {
                            Write-PSLog -Message "Account $bhrWorkEmail successfully created." -Severity Information

                            $managerUser = Get-ValidManagerUser -UserPrincipalName $bhrSupervisorEmail -Cache $performanceCache -TargetUser $bhrWorkEmail
                            if ($managerUser) {
                              $newManager = @{
                                '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($managerUser.Id)"
                              }
                              Start-Sleep -Seconds 8

                              Write-PSLog -Message "Setting manager for newly created user $bhrWorkEmail." -Severity Debug
                              Write-PSLog -Message "Executing: Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $NewManager" -Severity Debug
                              Invoke-WithRetry -Operation "Set manager for new user: $bhrWorkEmail" -ScriptBlock {
                                Set-MgUserManagerByRef -UserId $bhrWorkEmail -BodyParameter $NewManager
                              }
                              Add-SignificantChange -Category ManagerChanged -User $bhrWorkEmail -Detail "Manager: $bhrSupervisorEmail"
                            }
                            $params = @{
                              Message         = @{
                                Subject       = "User account created for: $bhrdisplayName"
                                Body          = @{
                                  ContentType = 'html'
                                  Content     = "<br/><br/><p>A new user account was created for $bhrDisplayName with hire date of $bhrHireDate. </p><p> $($Script:Config.Email.WelcomeUserText) <ul><li>User name: $bhrWorkEmail</li><li>Password: $($PasswordProfile.Values)</li></ul><br/><p>$($Script:Config.Email.EmailSignature)</p>"
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
                                      Address = $Script:Config.Email.NotificationEmailAddress
                                    }
                                  }
                                )
                              }
                              SaveToSentItems = 'True'
                            }
                            Write-PSLog -Message "Sending $bhrSupervisorEmail new employee information for $bhrDisplayName in email." -Severity Information
                            Invoke-WithRetry -Operation 'Send new employee notification to manager' -ScriptBlock {
                              Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress -Verbose
                            }

                            New-AdaptiveCard {
                              New-AdaptiveTextBlock -Text 'New user account created' -HorizontalAlignment Center -Weight Bolder -Wrap
                              New-AdaptiveTextBlock -Text "User name: $bhrWorkEmail" -Wrap
                              #New-AdaptiveTextBlock -Text "Password: $($PasswordProfile.Values)" -Wrap
                            } -Uri $Script:Config.Features.TeamsCardUri -Speak "New User $bhrDisplayName account created"

                            # Todo input these and an array and loop through only if needed.

                            # Give a little time for the mailbox to be setup so that it can receive the message.
                            Write-Output 'Waiting for mailbox setup before continuing'
                            Start-Sleep -Seconds 180
                            Write-Output 'Evaluating shared mailbox permissions'
                            Invoke-SharedMailboxDelegationSync -Reason 'new user onboarding'

                            $newUserWelcomeEmailParams = @{
                              Message         = @{
                                Subject       = "Welcome, $bhrFirstName!"
                                Body          = @{
                                  ContentType = 'html'
                                  Content     = "<br/><br/><p>Welcome to $CompanyName, $bhrFirstName!</p><br/>`
                              <p> $WelcomeUserText</p><br/>`
                              $($Script:Config.Email.WelcomeLinksHtml)<p>$EmailSignature</p>"
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
                              SaveToSentItems = 'True'
                            }
                            # Send Mail Message parameters definition closure
                            Write-Output "Sending welcome email to $bhrWorkEmail"
                            Invoke-WithRetry -Operation 'Send welcome email to new user' -ScriptBlock {
                              Send-MgUserMail -BodyParameter $newUserWelcomeEmailParams -UserId $Script:Config.Email.AdminEmailAddress -Verbose
                            }
                          }
                          else {
                            $params = @{
                              Message         = @{
                                Subject      = "User creation automation: $bhrdisplayName"
                                Body         = @{
                                  ContentType = 'html'
                                  Content     = "<br/><p>New employee user account created for $bhrDisplayName. No manager account is currently active for this account so this info is being sent to the default location.`
                                        <p> $($Script:Config.Email.WelcomeUserText) <ul><li>User name: $bhrWorkEmail</li><li>Password: $($PasswordProfile.Values)</li></ul></p><p>$($Script:Config.Email.EmailSignature)</p>"
                                }
                                ToRecipients = @(
                                  @{
                                    EmailAddress = @{
                                      Address = $Script:Config.Email.HelpDeskEmailAddress
                                    }
                                  }
                                )
                                CCRecipients = @(
                                  @{
                                    EmailAddress = @{
                                      Address = $Script:Config.Email.NotificationEmailAddress
                                    }
                                  }
                                )
                              }
                              SaveToSentItems = 'True'
                            }
                            Write-PSLog -Message 'Sending new employee information to default notification email because no manager was defined.' -Severity Information
                            Invoke-WithRetry -Operation 'Send welcome email (no manager assigned)' -ScriptBlock {
                              Send-MgUserMail -BodyParameter $params -UserId $Script:Config.Email.AdminEmailAddress -Verbose
                            }

                            New-AdaptiveCard {
                              New-AdaptiveTextBlock -Text 'New user account created without an assigned manager' -HorizontalAlignment Center -Weight Bolder -Wrap
                              New-AdaptiveTextBlock -Text 'No manager account is currently active for this account so this info is being sent to the default location.' -Wrap
                              New-AdaptiveTextBlock -Text "User name: $bhrWorkEmail" -Wrap
                              New-AdaptiveTextBlock -Text "Password: $($PasswordProfile.Values)" -Wrap
                            } -Uri $Script:Config.Features.TeamsCardUri -Speak "New User $bhrDisplayName Account Created"
                          }
                        }
                      }
                    }
                  }
                  else {
                    # If Hire Date is less than $days days in the future or in the past closure
                    # The user account does not need to be created as it does not satisfy the condition of the HireDate being $($Script:Config.Features.DaysAhead) days or less in the future
                    if ($bhrAccountEnabled) {
                      Write-PSLog -Message "$bhrWorkEmail's hire date ($bhrHireDate) is more than $($Script:Config.Features.DaysAhead) days from now." -Severity Information
                    }
                    else {
                      Write-PSLog -Message "$bhrWorkEmail has been terminated, the account will not be created." -Severity Debug
                    }
                  }

                  # Increment processed user counter for performance tracking
                  $processedUserCount++
                } # End ForEach-Object loop
              }

#endregion Employee Processing Pipeline

#region Script Completion and Reporting
<#
===============================================================================
SCRIPT COMPLETION AND REPORTING SECTION
===============================================================================

This section handles post-processing tasks and sends notifications.

WHAT HAPPENS HERE:

1. LICENSE CHECK:
   - Verifies sufficient licenses available for future user creation
   - Only runs if LicenseId is configured
   - Sends alerts if licenses running low

2. ERROR SUMMARY REPORT:
   - Analyzes all errors collected during processing
   - Categorizes by type (UserCreation, AttributeUpdate, etc.)
   - Identifies affected users
   - Sends email to admins with summary
   - Only runs if errors occurred

3. PERFORMANCE STATISTICS:
   - Calculates total runtime
   - Determines users per minute throughput
   - Shows cache effectiveness (hit rate)
   - Provides data for optimization decisions

4. NOTIFICATIONS:
   - Sends Teams adaptive card (if configured)
   - Includes log summary
   - Shows overall sync status

5. MAILBOX DELEGATION:
   - Syncs shared mailbox permissions (if ForceSharedMailboxPermissions)
   - Ensures delegations match BambooHR groups
   - Only runs if specifically requested

DEVELOPER NOTE:
This section only executes if:
- NOT in WhatIf mode (changes were actually applied)
- There is log content to report
In WhatIf mode, changes are previewed but not applied.
#>

$hasLogContent = ([string]::IsNullOrWhiteSpace($Script:logContent) -eq $false)
$changesWereApplied = ((-not $WhatIfPreference) -and $hasLogContent)

if ($changesWereApplied) {

  # Check license availability for future user creation
  $licenseInfo = $null
  if ([string]::IsNullOrWhiteSpace($Script:Config.Azure.LicenseId) -eq $false) {
    try {
      $licenseInfo = Get-LicenseStatus -LicenseId $Script:Config.Azure.LicenseId
    }
    catch {
      Write-PSLog "Failed to retrieve license status for Teams summary: $($_.Exception.Message)" -Severity Warning
    }
  }
  $runtime = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
  Write-PSLog -Message "`n Completed sync at $(Get-Date) and ran for $([math]::Round($runtime.TotalSeconds, 2)) seconds" -Severity Information

  # Generate and send error summary report if any errors occurred
  if ($errorSummary.TotalErrors -gt 0) {
    Write-PSLog "`n" -Severity Information
    Get-ErrorSummaryReport -ErrorSummary $errorSummary -SendEmail | Out-Null
  }
  else {
    Write-PSLog 'No errors encountered during sync - Success!' -Severity Information
  }

  # Display performance statistics
  if ($processedUserCount -gt 0) {
    $perfStats = Get-PerformanceStatistics -StartTime $Script:StartTime -UserCount $processedUserCount -Cache $performanceCache
    Write-Output "Performance: $($perfStats.UsersPerMinute) users/min, Cache hit rate: $($perfStats.CacheHitRate)%"
  }

  # Send Teams notification with changes applied
  if (-not [string]::IsNullOrWhiteSpace($Script:Config.Features.TeamsCardUri)) {
    Write-PSLog 'Teams notification needs to be sent and a URL exists' -Severity Debug
    try {
      $maxExamples = 8
      $createdCount = $Script:SignificantChanges.Created.Count
      $disabledCount = $Script:SignificantChanges.Disabled.Count
      $nameChangedCount = $Script:SignificantChanges.NameChanged.Count
      $upnChangedCount = $Script:SignificantChanges.UpnChanged.Count
      $managerChangedCount = $Script:SignificantChanges.ManagerChanged.Count
      $updatedMajorCount = $Script:SignificantChanges.UpdatedMajor.Count
      $hasSignificantChanges = ($createdCount + $disabledCount + $nameChangedCount + $upnChangedCount + $managerChangedCount + $updatedMajorCount) -gt 0

      New-AdaptiveCard {
        New-AdaptiveTextBlock -Text 'BambooHR to Entra ID Sync - Changes Applied' -Wrap -Weight Bolder -Color Good
        New-AdaptiveTextBlock -Text "Users Processed: $processedUserCount" -Wrap
        New-AdaptiveTextBlock -Text "Duration: $([math]::Round((New-TimeSpan -Start $Script:StartTime -End (Get-Date)).TotalMinutes, 2)) minutes" -Wrap
        if ($licenseInfo) {
          New-AdaptiveTextBlock -Text "Licenses: $($licenseInfo.ConsumedUnits) used / $($licenseInfo.AvailableUnits) available / $($licenseInfo.EnabledUnits) total" -Wrap
        }
        if ($errorSummary.TotalErrors -gt 0) {
          New-AdaptiveTextBlock -Text "⚠ Errors: $($errorSummary.TotalErrors)" -Wrap -Color Warning
        }
        else {
          New-AdaptiveTextBlock -Text '✓ No errors' -Wrap -Color Good
        }
        if ($hasSignificantChanges) {
          New-AdaptiveTextBlock -Text "`nSignificant changes:" -Wrap -Weight Bolder

          if ($createdCount -gt 0) {
            New-AdaptiveTextBlock -Text "Created: $createdCount" -Wrap
            $Script:SignificantChanges.Created.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "+ $($_.Name)" } else { "+ $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($disabledCount -gt 0) {
            New-AdaptiveTextBlock -Text "Disabled: $disabledCount" -Wrap
            $Script:SignificantChanges.Disabled.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "- $($_.Name)" } else { "- $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($nameChangedCount -gt 0) {
            New-AdaptiveTextBlock -Text "Name changes: $nameChangedCount" -Wrap
            $Script:SignificantChanges.NameChanged.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "• $($_.Name)" } else { "• $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($upnChangedCount -gt 0) {
            New-AdaptiveTextBlock -Text "UPN changes: $upnChangedCount" -Wrap
            $Script:SignificantChanges.UpnChanged.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "• $($_.Name)" } else { "• $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($managerChangedCount -gt 0) {
            New-AdaptiveTextBlock -Text "Manager changes: $managerChangedCount" -Wrap
            $Script:SignificantChanges.ManagerChanged.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "• $($_.Name)" } else { "• $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($updatedMajorCount -gt 0) {
            New-AdaptiveTextBlock -Text "Other major updates: $updatedMajorCount" -Wrap
            $Script:SignificantChanges.UpdatedMajor.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "• $($_.Name)" } else { "• $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
        }
        else {
          New-AdaptiveTextBlock -Text "`nNo significant changes detected" -Wrap
        }
        New-AdaptiveTextBlock -Text "`nCompleted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Wrap -Size Small
      } -Uri $Script:Config.Features.TeamsCardUri -Speak 'BambooHR to Entra ID sync completed with changes applied'
      Write-PSLog 'Teams notification sent: Changes applied' -Severity Debug
    }
    catch {
      Write-PSLog "Failed to send Teams notification: $($_.Exception.Message)" -Severity Warning
    }
  }

  Start-Sleep 30
  # Todo input these and an array and loop through only if needed.
}

if (-not $changesWereApplied) {
  $licenseInfo = $null
  if ([string]::IsNullOrWhiteSpace($Script:Config.Azure.LicenseId) -eq $false) {
    try {
      $licenseInfo = Get-LicenseStatus -LicenseId $Script:Config.Azure.LicenseId
    }
    catch {
      Write-PSLog "Failed to retrieve license status for Teams summary: $($_.Exception.Message)" -Severity Warning
    }
  }

  $runtime = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
  Write-PSLog -Message "`n Completed sync at $(Get-Date) and ran for $([math]::Round($runtime.TotalSeconds, 2)) seconds" -Severity Information

  # Send Teams summary card
  if (-not [string]::IsNullOrWhiteSpace($Script:Config.Features.TeamsCardUri)) {
    try {
      $maxExamples = 8
      $createdCount = $Script:SignificantChanges.Created.Count
      $disabledCount = $Script:SignificantChanges.Disabled.Count
      $nameChangedCount = $Script:SignificantChanges.NameChanged.Count
      $upnChangedCount = $Script:SignificantChanges.UpnChanged.Count
      $managerChangedCount = $Script:SignificantChanges.ManagerChanged.Count
      $updatedMajorCount = $Script:SignificantChanges.UpdatedMajor.Count
      $hasSignificantChanges = ($createdCount + $disabledCount + $nameChangedCount + $upnChangedCount + $managerChangedCount + $updatedMajorCount) -gt 0

      if (-not $hasSignificantChanges) {
        # No changes were made (WhatIf mode or no updates needed)
        New-AdaptiveCard {
          New-AdaptiveTextBlock -Text 'BambooHR to Entra ID Sync Completed' -Wrap -Weight Bolder
          New-AdaptiveTextBlock -Text 'Status: No changes required' -Wrap
          New-AdaptiveTextBlock -Text 'Mode: WhatIf Preview' -Wrap -Color Accent
          New-AdaptiveTextBlock -Text "Duration: $([math]::Round((New-TimeSpan -Start $Script:StartTime -End (Get-Date)).TotalMinutes, 2)) minutes" -Wrap
          if ($licenseInfo) {
            New-AdaptiveTextBlock -Text "Licenses: $($licenseInfo.ConsumedUnits) used / $($licenseInfo.AvailableUnits) available / $($licenseInfo.EnabledUnits) total" -Wrap
          }
          New-AdaptiveTextBlock -Text "Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Wrap
        } -Uri $Script:Config.Features.TeamsCardUri -Speak 'BambooHR to Entra ID sync completed with no changes'
        Write-PSLog 'Teams notification sent: No changes made' -Severity Information
      }
      else {

        New-AdaptiveCard {
          New-AdaptiveTextBlock -Text 'BambooHR to Entra ID Sync Completed' -Wrap -Weight Bolder
          New-AdaptiveTextBlock -Text "Duration: $([math]::Round((New-TimeSpan -Start $Script:StartTime -End (Get-Date)).TotalMinutes, 2)) minutes" -Wrap
          if ($licenseInfo) {
            New-AdaptiveTextBlock -Text "Licenses: $($licenseInfo.ConsumedUnits) used / $($licenseInfo.AvailableUnits) available / $($licenseInfo.EnabledUnits) total" -Wrap
          }
          New-AdaptiveTextBlock -Text "`nSignificant changes:" -Wrap -Weight Bolder

          if ($createdCount -gt 0) {
            New-AdaptiveTextBlock -Text "Created: $createdCount" -Wrap
            $Script:SignificantChanges.Created.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "+ $($_.Name)" } else { "+ $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($disabledCount -gt 0) {
            New-AdaptiveTextBlock -Text "Disabled: $disabledCount" -Wrap
            $Script:SignificantChanges.Disabled.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "- $($_.Name)" } else { "- $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($nameChangedCount -gt 0) {
            New-AdaptiveTextBlock -Text "Name changes: $nameChangedCount" -Wrap
            $Script:SignificantChanges.NameChanged.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "- $($_.Name)" } else { "- $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($upnChangedCount -gt 0) {
            New-AdaptiveTextBlock -Text "UPN changes: $upnChangedCount" -Wrap
            $Script:SignificantChanges.UpnChanged.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "- $($_.Name)" } else { "- $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($managerChangedCount -gt 0) {
            New-AdaptiveTextBlock -Text "Manager changes: $managerChangedCount" -Wrap
            $Script:SignificantChanges.ManagerChanged.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "- $($_.Name)" } else { "- $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          if ($updatedMajorCount -gt 0) {
            New-AdaptiveTextBlock -Text "Other major updates: $updatedMajorCount" -Wrap
            $Script:SignificantChanges.UpdatedMajor.GetEnumerator() | Sort-Object Name | Select-Object -First $maxExamples | ForEach-Object {
              $detail = $_.Value
              $line = if ([string]::IsNullOrWhiteSpace($detail)) { "- $($_.Name)" } else { "- $($_.Name) ($detail)" }
              New-AdaptiveTextBlock -Text $line -Wrap
            }
          }
          New-AdaptiveTextBlock -Text "`nCompleted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Wrap -Size Small
        } -Uri $Script:Config.Features.TeamsCardUri -Speak 'BambooHR to Entra ID sync completed successfully'
        Write-PSLog 'Teams notification sent: Sync summary with changes' -Severity Information
      }
    }
    catch {
      Write-PSLog "Failed to send Teams notification: $($_.Exception.Message)" -Severity Warning
    }
  }
}

Invoke-SharedMailboxDelegationSync -Reason 'run completion'

function Write-TerminatedAccountDeletionReminders {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [int]
    $DaysToKeepAccountsAfterTermination
  )

  if ($DaysToKeepAccountsAfterTermination -le 0) {
    return
  }

  $cutoffUtc = (Get-Date).ToUniversalTime().AddDays(-$DaysToKeepAccountsAfterTermination)
  Write-PSLog -Message "Checking for terminated accounts older than $DaysToKeepAccountsAfterTermination days (manual deletion cutoff: $($cutoffUtc.ToString('yyyy-MM-dd')))" -Severity Debug

  $users = $null
  try {
    $users = Invoke-WithRetry -Operation 'Query terminated accounts for deletion reminders' -ScriptBlock {
      Get-MgUser -All -Filter "employeeId eq 'LVR'" -Property 'id,displayName,userPrincipalName,mail,employeeId,accountEnabled,employeeLeaveDateTime,department,companyName'
    }
  }
  catch {
    Write-PSLog -Message "Primary query failed (employeeId filter). Falling back to disabled-user scan: $($_.Exception.Message)" -Severity Warning
    try {
      $users = Invoke-WithRetry -Operation 'Fallback query disabled accounts for deletion reminders' -ScriptBlock {
        Get-MgUser -All -Filter 'accountEnabled eq false' -Property 'id,displayName,userPrincipalName,mail,employeeId,accountEnabled,employeeLeaveDateTime,department,companyName'
      }
      if ($users) {
        $users = $users | Where-Object {
          ($_.EmployeeId -eq 'LVR') -or ($_.Department -eq 'Not Active') -or (-not [string]::IsNullOrWhiteSpace($_.CompanyName) -and $_.CompanyName -match 'OffboardingComplete')
        }
      }
    }
    catch {
      Write-PSLog -Message "Fallback query also failed for deletion reminders: $($_.Exception.Message)" -Severity Warning
      return
    }
  }

  if (-not $users) {
    return
  }

  foreach ($u in $users) {
    if ($u.AccountEnabled -ne $false) {
      continue
    }

    $leaveUtc = $null
    if ($u.EmployeeLeaveDateTime) {
      try { $leaveUtc = ([datetime]$u.EmployeeLeaveDateTime).ToUniversalTime() } catch { $leaveUtc = $null }
    }

    if (-not $leaveUtc -and -not [string]::IsNullOrWhiteSpace($u.CompanyName)) {
      $leaveUtc = Get-OffboardingCompletionDateFromCompanyName -CompanyName $u.CompanyName
    }

    if (-not $leaveUtc) {
      continue
    }

    if ($leaveUtc -le $cutoffUtc) {
      $ageDays = [math]::Floor(((Get-Date).ToUniversalTime() - $leaveUtc).TotalDays)
      $upn = if ([string]::IsNullOrWhiteSpace($u.UserPrincipalName)) { $u.Mail } else { $u.UserPrincipalName }
      Write-PSLog -Message "Manual deletion required: $upn (terminated $ageDays days ago; leaveDateTime=$($leaveUtc.ToString('yyyy-MM-dd')))." -Severity Warning
    }
  }
}

#Script End
Write-TerminatedAccountDeletionReminders -DaysToKeepAccountsAfterTermination $Script:Config.Features.DaysToKeepAccountsAfterTermination
exit 0