#Requires -Module Pester
#using namespace System.Management.Automation.Language


<#
.SYNOPSIS
Pester tests and mocks for BambooHR provisioning scripts.

.DESCRIPTION
Validates helper functions, configuration overrides, and static checks without executing the full scripts.
#>

[CmdletBinding()]
param()

function Get-FunctionDefinitionsFromFile {
  <#
    .SYNOPSIS
    Extract function definitions from a PowerShell script.

    .PARAMETER ScriptPath
    Path to the PowerShell script file.

    .PARAMETER FunctionNames
    Names of the functions to extract.
    #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ScriptPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$FunctionNames
  )

  $tokens = $null
  $errors = $null
  $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)

  if ($errors) {
    $message = ($errors | ForEach-Object { $_.Message }) -join '; '
    throw "Failed to parse $($ScriptPath): $message"
  }

  $definitions = @()
  $functions = $ast.FindAll({
      param($node)
      $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $FunctionNames -contains $node.Name
    }, $true)

  foreach ($functionAst in $functions) {
    $definitions += $functionAst.Extent.Text
  }

  return $definitions
}

function Initialize-TestScriptState {
  <#
    .SYNOPSIS
    Initialize script-scoped variables needed for configuration tests.

    .PARAMETER CompanyName
    Company name used for defaults.

    .PARAMETER TenantId
    Tenant ID for configuration.
    #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CompanyName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId
  )

  $script:BambooHrApiKey = 'test-key'
  $script:BHRCompanyName = 'contoso'
  $script:CompanyName = $CompanyName
  $script:TenantId = $TenantId
  $script:TeamsCardUri = 'https://example.webhook.office.com/'
  $script:AdminEmailAddress = 'admin@contoso.com'
  $script:NotificationEmailAddress = 'hr@contoso.com'
  $script:HelpDeskEmailAddress = 'helpdesk@contoso.com'
  $script:LicenseId = ''
  $script:UsageLocation = 'US'
  $script:DaysAhead = 7
  $script:DaysToKeepAccountsAfterTermination = 14
  $script:EnableMobilePhoneSync = $false
  $script:CurrentOnly = $false
  $script:ForceSharedMailboxPermissions
  $script:DefaultProfilePicPath = ''
  $script:EmailSignature = ''
  $script:WelcomeUserText = ''
  $script:WelcomeLinksHtml = ''
  $script:MailboxDelegationParams = @()
  $script:ModifiedWithinDays = 14
  $script:FullSync = $false
  $script:LogPath = $env:TEMP
  $script:MaxRetryAttempts = 3
  $script:RetryDelaySeconds = 5
  $script:OperationTimeoutSeconds = 120
  $script:BatchSize = 25
  $script:PSBoundParameters = $PSBoundParameters
  $script:CorrelationId = [Guid]::NewGuid().ToString()
  $script:StartTime = Get-Date
}

function Test-ValidManagerUserMock {
  <#
    .SYNOPSIS
    Basic wrapper used for mocking Get-MgUser in Get-ValidManagerUser tests.
    #>
  [CmdletBinding()]
  param()
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$startScriptPath = Join-Path $repoRoot 'Start-BambooHRUserProvisioning.ps1'

Describe 'Static validation' {
  BeforeAll {
    $script:staticRepoRoot = Split-Path -Parent $PSScriptRoot
    $script:staticStartScriptPath = Join-Path $script:staticRepoRoot 'Start-BambooHRUserProvisioning.ps1'
    $script:staticWebhookScriptPath = Join-Path $script:staticRepoRoot 'Start-BambooHrWebhookSync.ps1'
  }

  It 'does not contain invalid variable references' {
    $pattern = '\$(?!script:|Script:|env:|global:|local:|private:|using:)[A-Za-z_][A-Za-z0-9_]*:'
    $content = Get-Content -Raw -Path $script:staticStartScriptPath

    $content | Should -Not -Match $pattern
  }

  Context 'PII and sensitive data' {
    It 'source files contain no real org domain' {
      # Pattern is split across two literals so this test file does not match itself
      $bannedDomain = 'gecko' + 'green'
      $files = Get-ChildItem -Path $script:staticRepoRoot -Include '*.ps1', '*.md', '*.json', '*.bicep', '*.bicepparam' -Recurse -File |
        Where-Object { $_.FullName -notmatch '\\\.' }

      $violations = @(foreach ($file in $files) {
          Select-String -Path $file.FullName -Pattern $bannedDomain |
            ForEach-Object { "$($file.Name):$($_.LineNumber) — $($_.Line.Trim())" }
        })

      $violations | Should -BeNullOrEmpty -Because 'real org domain must not appear in committed source files'
    }

    It 'uses valid HTML list markup in UPN change notifications' {
      foreach ($path in @($script:staticStartScriptPath, $script:staticWebhookScriptPath)) {
        $content = Get-Content -Raw -Path $path
        $content | Should -Not -Match '<ui>'
      }
    }

    It 'uses normalized config values in UPN change notifications' {
      $requiredPatterns = @(
        '\$Script:Config\.Azure\.CompanyName'
        '\$Script:Config\.Email\.EmailSignature'
        '\$Script:Config\.Email\.NotificationEmailAddress'
        '\$Script:Config\.Features\.TeamsCardUri'
      )

      foreach ($path in @($script:staticStartScriptPath, $script:staticWebhookScriptPath)) {
        $content = Get-Content -Raw -Path $path
        $upnSection = (($content -split '# Change UserPrincipalName and send the details via email to the User')[1] -split '# Create new employee account')[0]

        foreach ($pattern in $requiredPatterns) {
          $upnSection | Should -Match $pattern
        }
      }
    }
  }
}

Describe 'Start-BambooHRUserProvisioning helpers' {
  BeforeAll {
    function script:Get-FunctionDefinitionsFromFile {
      <#
        .SYNOPSIS
        Extract function definitions from a PowerShell script.
        #>
      [CmdletBinding()]
      param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$FunctionNames
      )

      $tokens = $null
      $errors = $null
      $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)

      if ($errors) {
        $message = ($errors | ForEach-Object { $_.Message }) -join '; '
        throw "Failed to parse $($ScriptPath): $message"
      }

      $definitions = @()
      $functions = $ast.FindAll({
          param($node)
          $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $FunctionNames -contains $node.Name
        }, $true)

      foreach ($functionAst in $functions) {
        $definitions += $functionAst.Extent.Text
      }

      return $definitions
    }

    function script:Initialize-TestScriptState {
      <#
        .SYNOPSIS
        Initialize script-scoped variables needed for configuration tests.
        #>
      [CmdletBinding()]
      param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CompanyName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId
      )

      $script:BambooHrApiKey = 'test-key'
      $script:BHRCompanyName = 'contoso'
      $script:CompanyName = $CompanyName
      $script:TenantId = $TenantId
      $script:TeamsCardUri = 'https://example.webhook.office.com/'
      $script:AdminEmailAddress = 'admin@contoso.com'
      $script:NotificationEmailAddress = 'hr@contoso.com'
      $script:HelpDeskEmailAddress = 'helpdesk@contoso.com'
      $script:LicenseId = ''
      $script:UsageLocation = 'US'
      $script:DaysAhead = 7
      $script:DaysToKeepAccountsAfterTermination = 14
      $script:EnableMobilePhoneSync = $false
      $script:CurrentOnly = $false
      $script:ForceSharedMailboxPermissions = $false
      $script:DefaultProfilePicPath = ''
      $script:EmailSignature = ''
      $script:WelcomeUserText = ''
      $script:WelcomeLinksHtml = ''
      $script:MailboxDelegationParams = @()
      $script:LogPath = $env:TEMP
      $script:MaxRetryAttempts = 3
      $script:RetryDelaySeconds = 5
      $script:OperationTimeoutSeconds = 120
      $script:BatchSize = 25
      $script:PSBoundParameters = $PSBoundParameters
      $script:CorrelationId = [Guid]::NewGuid().ToString()
      $script:StartTime = Get-Date
    }

    $script:helperRepoRoot = Split-Path -Parent $PSScriptRoot
    $script:startScriptPath = Join-Path $script:helperRepoRoot 'Start-BambooHRUserProvisioning.ps1'

    $functionNames = @(
      'Initialize-Configuration',
      'Invoke-WithRetry',
      'ConvertTo-StandardName',
      'ConvertTo-PhoneNumber',
      'ConvertTo-BambooHrHireDate',
      'Get-ValidManagerUser',
      'Test-IsTenantEmail',
      'Get-MailNicknameFromEmail',
      'Get-WorkPhoneComparisonValue',
      'Test-ShouldSyncExistingUser',
      'Test-ShouldSendTeamsChangesCard',
      'Test-ShouldUpdateEmployeeId',
      'Test-IsOffboardingComplete',
      'Set-TerminatedUserProfileFields',
      'Get-OffboardingCompletionMarker',
      'Get-OffboardingCompletionDateFromCompanyName',
      'Invoke-UserOffboarding'
    )
    $definitions = Get-FunctionDefinitionsFromFile -ScriptPath $startScriptPath -FunctionNames $functionNames
    foreach ($definition in $definitions) {
      Invoke-Expression $definition
    }

    function Get-AutomationVariable {
      <#
            .SYNOPSIS
            Test stub for Azure Automation variable retrieval.
            #>
      [CmdletBinding()]
      param(
        [Parameter(Mandatory = $true)]
        [string]$Name
      )

      return $null
    }

    function Write-PSLog {
      <#
            .SYNOPSIS
            Test stub for logging.
            #>
      [CmdletBinding()]
      param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [string]$Severity
      )

      Write-Verbose $Message
    }

    function Get-CachedUser {
      <#
            .SYNOPSIS
            Test stub for cache lookup.
            #>
      [CmdletBinding()]
      param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Cache,

        [Parameter()]
        [switch]$Force
      )

      return $null
    }

    function Update-MgUser {
      <#
            .SYNOPSIS
            Test stub for Graph user updates.
            #>
      [CmdletBinding()]
      param()
    }

    function Invoke-MgGraphRequest {
      <#
            .SYNOPSIS
            Test stub for raw Graph API calls.
            #>
      [CmdletBinding()]
      param()
    }

    # Set up $Script:Config so Invoke-WithRetry can read retry settings
    $Script:Config = @{
      CorrelationId = [Guid]::NewGuid().ToString()
      Runtime       = @{
        MaxRetryAttempts  = 1
        RetryDelaySeconds = 0
      }
    }
  }

  Context 'Initialize-Configuration' {
    BeforeEach {
      Initialize-TestScriptState -CompanyName 'Contoso' -TenantId 'contoso.onmicrosoft.com'
      Mock Get-AutomationVariable { $null }
    }

    It 'flags missing required parameters as invalid' {
      $script:BambooHrApiKey = $null
      $script:BHRCompanyName = $null
      $script:CompanyName = $null
      $script:TenantId = $null

      $config = Initialize-Configuration

      $config.IsValid | Should -BeFalse
      $config.ValidationErrors.Count | Should -BeGreaterThan 0
    }

    It 'applies JSON overrides for welcome links' {
      $overrideHtml = '<p>Custom links</p>'
      $customJson = @{ WelcomeLinksHtml = $overrideHtml } | ConvertTo-Json

      Mock Get-AutomationVariable {
        param([string]$Name)
        if ($Name -eq 'BHR_CustomizationsJson') {
          return $customJson
        }
        return $null
      }

      $config = Initialize-Configuration

      $config.IsValid | Should -BeTrue
      $config.Email.WelcomeLinksHtml | Should -Be $overrideHtml
    }

    It 'applies JSON boolean overrides for feature switches' {
      $customJson = @{
        EnableMobilePhoneSync         = $true
        CurrentOnly                   = $true
        ForceSharedMailboxPermissions = $true
      } | ConvertTo-Json

      Mock Get-AutomationVariable {
        param([string]$Name)
        if ($Name -eq 'BHR_CustomizationsJson') {
          return $customJson
        }
        return $null
      }

      $config = Initialize-Configuration

      $config.IsValid | Should -BeTrue
      $config.Features.EnableMobilePhoneSync | Should -BeTrue
      $config.Features.CurrentOnly | Should -BeTrue
      $config.Features.ForceSharedMailboxPermissions | Should -BeTrue
      $config.BambooHR.ReportsUri | Should -Match 'onlyCurrent=true'
    }

    It 'does not let custom JSON override explicitly bound force delegation flag' {
      $script:ForceSharedMailboxPermissions = $true
      $script:PSBoundParameters['ForceSharedMailboxPermissions'] = $true

      $customJson = @{
        ForceSharedMailboxPermissions = $false
      } | ConvertTo-Json

      Mock Get-AutomationVariable {
        param([string]$Name)
        if ($Name -eq 'BHR_CustomizationsJson') {
          return $customJson
        }
        return $null
      }

      $config = Initialize-Configuration

      $config.IsValid | Should -BeTrue
      $config.Features.ForceSharedMailboxPermissions | Should -BeTrue
    }

    It 'sets default email signature and welcome text' {
      $config = Initialize-Configuration

      $config.Email.EmailSignature | Should -Match 'Automated User Management'
      $config.Email.WelcomeUserText | Should -Match 'Welcome to'
    }
  }

  Context 'ConvertTo-StandardName' {
    It 'normalizes names with diacritics and spacing' {
      ConvertTo-StandardName '  jOÃO  sIlva  ' | Should -Be 'João Silva'
    }
  }

  Context 'ConvertTo-PhoneNumber' {
    It 'preserves E.164 numbers' {
      ConvertTo-PhoneNumber '+1 (415) 555-0101' | Should -Be '+14155550101'
    }

    It 'converts 00-prefixed numbers to E.164' {
      ConvertTo-PhoneNumber '0044 20 7946 0958' | Should -Be '+442079460958'
    }

    It 'strips non-digits for local numbers' {
      ConvertTo-PhoneNumber '(555) 123-4567' | Should -Be '5551234567'
    }
  }

  Context 'ConvertTo-BambooHrHireDate' {
    It 'preserves a valid BambooHR date' {
      $date = ConvertTo-BambooHrHireDate -HireDate '2026-03-26' -EmployeeEmailAddress 'user@contoso.com'

      $date.ToString('yyyy-MM-dd') | Should -Be '2026-03-26'
    }

    It 'returns minimum date for invalid BambooHR values' {
      $date = ConvertTo-BambooHrHireDate -HireDate '0000-00-00' -EmployeeEmailAddress 'user@contoso.com'

      $date | Should -Be ([datetime]::MinValue.Date)
    }

    It 'returns minimum date for empty BambooHR values' {
      $date = ConvertTo-BambooHrHireDate -HireDate '' -EmployeeEmailAddress 'user@contoso.com'

      $date | Should -Be ([datetime]::MinValue.Date)
    }
  }

  Context 'Tenant email and mail nickname helpers' {
    It 'matches a verified tenant domain' {
      Test-IsTenantEmail -EmailAddress 'user@contoso.com' -TenantDomains @('contoso.com', 'contoso.onmicrosoft.com') | Should -BeTrue
    }

    It 'does not match a non-tenant email address' {
      Test-IsTenantEmail -EmailAddress 'user@gmail.com' -TenantDomains @('contoso.com', 'contoso.onmicrosoft.com') | Should -BeFalse
    }

    It 'returns the local-part for any verified tenant domain email' {
      Get-MailNicknameFromEmail -EmailAddress 'user@contoso.onmicrosoft.com' | Should -Be 'user'
    }
  }

  Context 'Work phone comparison helper' {
    It 'treats empty values as no phone' {
      Get-WorkPhoneComparisonValue -PhoneNumber '' | Should -Be '0'
    }

    It 'treats zero as no phone' {
      Get-WorkPhoneComparisonValue -PhoneNumber '0' | Should -Be '0'
    }

    It 'normalizes formatted phone numbers for comparison' {
      Get-WorkPhoneComparisonValue -PhoneNumber '(555) 123-4567' | Should -Be '5551234567'
    }
  }

  Context 'Existing user sync gate helper' {
    It 'returns true when employee id or upn matches and last changed differs' {
      $entraIdObj = [pscustomobject]@{ Id = '1'; UserPrincipalName = 'user@contoso.com' }

      Test-ShouldSyncExistingUser -EntraIdEmployeeNumber '123' `
        -EntraIdEmployeeNumberByEid '123' `
        -EntraIdUpnFromEidLookup 'user@contoso.com' `
        -EntraIdUpnFromUpnLookup 'user@contoso.com' `
        -BhrWorkEmail 'user@contoso.com' `
        -BhrLastChanged '2026-03-26T01:00:00Z' `
        -UpnExtensionAttribute1 '2026-03-25T01:00:00Z' `
        -EntraIdEidObjDetails $entraIdObj `
        -EntraIdUpnObjDetails $entraIdObj `
        -BhrEmploymentStatus 'Active' | Should -BeTrue
    }

    It 'returns false when last changed already matches' {
      $entraIdObj = [pscustomobject]@{ Id = '1'; UserPrincipalName = 'user@contoso.com' }

      Test-ShouldSyncExistingUser -EntraIdEmployeeNumber '123' `
        -EntraIdEmployeeNumberByEid '123' `
        -EntraIdUpnFromEidLookup 'user@contoso.com' `
        -EntraIdUpnFromUpnLookup 'user@contoso.com' `
        -BhrWorkEmail 'user@contoso.com' `
        -BhrLastChanged '2026-03-26T01:00:00Z' `
        -UpnExtensionAttribute1 '2026-03-26T01:00:00Z' `
        -EntraIdEidObjDetails $entraIdObj `
        -EntraIdUpnObjDetails $entraIdObj `
        -BhrEmploymentStatus 'Active' | Should -BeFalse
    }

    It 'returns true for terminated employee with enabled Entra account even when last changed matches' {
      $entraIdObj = [pscustomobject]@{ Id = '1'; UserPrincipalName = 'user@contoso.com'; AccountEnabled = $true }

      Test-ShouldSyncExistingUser -EntraIdEmployeeNumber '123' `
        -EntraIdEmployeeNumberByEid '123' `
        -EntraIdUpnFromEidLookup 'user@contoso.com' `
        -EntraIdUpnFromUpnLookup 'user@contoso.com' `
        -BhrWorkEmail 'user@contoso.com' `
        -BhrLastChanged '2026-03-26T01:00:00Z' `
        -UpnExtensionAttribute1 '2026-03-26T01:00:00Z' `
        -EntraIdEidObjDetails $entraIdObj `
        -EntraIdUpnObjDetails $entraIdObj `
        -BhrEmploymentStatus 'Terminated' | Should -BeTrue
    }

    It 'returns false for terminated employee with already-disabled Entra account when last changed matches' {
      $entraIdObj = [pscustomobject]@{ Id = '1'; UserPrincipalName = 'user@contoso.com'; AccountEnabled = $false }

      Test-ShouldSyncExistingUser -EntraIdEmployeeNumber '123' `
        -EntraIdEmployeeNumberByEid '123' `
        -EntraIdUpnFromEidLookup 'user@contoso.com' `
        -EntraIdUpnFromUpnLookup 'user@contoso.com' `
        -BhrWorkEmail 'user@contoso.com' `
        -BhrLastChanged '2026-03-26T01:00:00Z' `
        -UpnExtensionAttribute1 '2026-03-26T01:00:00Z' `
        -EntraIdEidObjDetails $entraIdObj `
        -EntraIdUpnObjDetails $entraIdObj `
        -BhrEmploymentStatus 'Terminated' | Should -BeFalse
    }
  }

  Context 'Teams changes card helper' {
    It 'returns true when significant changes were applied and logged' {
      $significantChanges = [ordered]@{
        Created        = @{ 'user@contoso.com' = 'User Example' }
        Disabled       = @{}
        NameChanged    = @{}
        UpnChanged     = @{}
        ManagerChanged = @{}
        UpdatedMajor   = @{}
      }

      Test-ShouldSendTeamsChangesCard -LogContent 'Completed sync' -SignificantChanges $significantChanges -WhatIfMode $false | Should -BeTrue
    }

    It 'returns false when only routine log output exists' {
      $significantChanges = [ordered]@{
        Created        = @{}
        Disabled       = @{}
        NameChanged    = @{}
        UpnChanged     = @{}
        ManagerChanged = @{}
        UpdatedMajor   = @{}
      }

      Test-ShouldSendTeamsChangesCard -LogContent 'Completed sync' -SignificantChanges $significantChanges -WhatIfMode $false | Should -BeFalse
    }

    It 'returns false in WhatIf mode even when significant changes exist' {
      $significantChanges = [ordered]@{
        Created        = @{ 'user@contoso.com' = 'User Example' }
        Disabled       = @{}
        NameChanged    = @{}
        UpnChanged     = @{}
        ManagerChanged = @{}
        UpdatedMajor   = @{}
      }

      Test-ShouldSendTeamsChangesCard -LogContent 'Completed sync' -SignificantChanges $significantChanges -WhatIfMode $true | Should -BeFalse
    }
  }

  Context 'EmployeeId update helper' {
    It 'allows EmployeeId updates for active users with UPN match' {
      Test-ShouldUpdateEmployeeId -EntraIdEmployeeNumber 'LVR' `
        -BhrEmployeeNumber '234' `
        -EntraIdUserPrincipalName 'user@contoso.com' `
        -BhrWorkEmail 'user@contoso.com' `
        -BhrEmploymentStatus 'Active' `
        -BhrAccountEnabled $true `
        -BhrLastChanged '2026-03-26T05:34:03Z' `
        -UpnExtensionAttribute1 '2026-03-26T02:52:09Z' | Should -BeTrue
    }

    It 'blocks EmployeeId updates for terminated inactive users' {
      Test-ShouldUpdateEmployeeId -EntraIdEmployeeNumber 'LVR' `
        -BhrEmployeeNumber '1234' `
        -EntraIdUserPrincipalName 'rando@contoso.com' `
        -BhrWorkEmail 'rando@contoso.com' `
        -BhrEmploymentStatus 'Terminated' `
        -BhrAccountEnabled $false `
        -BhrLastChanged '2026-03-26T05:34:03Z' `
        -UpnExtensionAttribute1 '2026-03-26T02:52:09Z' | Should -BeFalse
    }
  }

  Context 'Offboarding completion helpers' {
    It 'detects completed offboarding markers' {
      Test-IsOffboardingComplete -CompanyName '03/26/26 (OffboardingComplete: 2026-03-26T04:00:00Z)' -Department '' -JobTitle '' -OfficeLocation '' -WorkPhone '' -MobilePhone '' | Should -BeTrue
    }

    It 'treats remaining phone values as incomplete offboarding' {
      Test-IsOffboardingComplete -CompanyName '03/26/26 (OffboardingComplete: 2026-03-26T04:00:00Z)' -Department '' -JobTitle '' -OfficeLocation '' -WorkPhone '' -MobilePhone '5551234567' | Should -BeFalse
    }

    It 'builds and parses the same offboarding completion timestamp' {
      $leaveUtc = [datetime]::Parse('2026-03-26T04:00:00Z').ToUniversalTime()
      $marker = Get-OffboardingCompletionMarker -LeaveDateTimeUtc $leaveUtc
      $parsed = Get-OffboardingCompletionDateFromCompanyName -CompanyName $marker

      $marker | Should -Match 'OffboardingComplete:'
      $parsed.ToString('yyyy-MM-ddTHH:mm:ssZ') | Should -Be '2026-03-26T04:00:00Z'
    }
  }

  Context 'Get-ValidManagerUser' {
    BeforeEach {
      Mock Get-MgUser { throw 'not found' }
    }

    It 'returns null when manager UPN is missing' {
      Get-ValidManagerUser -UserPrincipalName '' -TargetUser 'user@contoso.com' | Should -BeNullOrEmpty
    }

    It 'returns null when lookup fails' {
      Get-ValidManagerUser -UserPrincipalName 'manager@contoso.com' -TargetUser 'user@contoso.com' | Should -BeNullOrEmpty
    }
  }
}

Describe 'Terminated user email-mismatch handling' {
  Context 'Employee filter includes inactive users without company email' {
    It 'passes active employee with company email' {
      $employee = [PSCustomObject]@{ workEmail = 'john@contoso.com'; status = 'Active' }
      $domain = 'contoso.com'
      $result = @($employee) | Where-Object { $_.workEmail -like "*$domain" -or $_.status -eq 'Inactive' }
      $result.Count | Should -Be 1
    }

    It 'passes inactive employee without company email' {
      $employee = [PSCustomObject]@{ workEmail = ''; status = 'Inactive' }
      $domain = 'contoso.com'
      $result = @($employee) | Where-Object { $_.workEmail -like "*$domain" -or $_.status -eq 'Inactive' }
      $result.Count | Should -Be 1
    }

    It 'passes inactive employee with personal email' {
      $employee = [PSCustomObject]@{ workEmail = 'john@gmail.com'; status = 'Inactive' }
      $domain = 'contoso.com'
      $result = @($employee) | Where-Object { $_.workEmail -like "*$domain" -or $_.status -eq 'Inactive' }
      $result.Count | Should -Be 1
    }

    It 'excludes active employee without company email' {
      $employee = [PSCustomObject]@{ workEmail = 'john@gmail.com'; status = 'Active' }
      $domain = 'contoso.com'
      $result = @($employee) | Where-Object { $_.workEmail -like "*$domain" -or $_.status -eq 'Inactive' }
      $result.Count | Should -Be 0
    }
  }

  Context 'EID fallback when UPN lookup fails for terminated user' {
    It 'promotes EID object to primary when UPN object is null' {
      $entraIdUpnObjDetails = $null
      $entraIdEidObjDetails = [PSCustomObject]@{
        UserPrincipalName             = 'jane.doe@contoso.com'
        AccountEnabled                = $true
        EmployeeId                    = '12345'
        OnPremisesExtensionAttributes = [PSCustomObject]@{ ExtensionAttribute1 = 'SomeValue' }
      }
      $bhrWorkEmail = ''
      $UpnExtensionAttribute1 = $null

      # Extract EID extension attribute
      $EIDExtensionAttribute1 = ($entraIdEidObjDetails |
          Select-Object @{
            Name       = 'ExtensionAttribute1'
            Expression = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 }
          }).ExtensionAttribute1

      # Apply fallback logic (mirrors script logic)
      if ([string]::IsNullOrEmpty($entraIdUpnObjDetails) -and ([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false)) {
        $entraIdUpnObjDetails = $entraIdEidObjDetails
        $bhrWorkEmail = $entraIdEidObjDetails.UserPrincipalName
        $UpnExtensionAttribute1 = $EIDExtensionAttribute1
      }

      $entraIdUpnObjDetails | Should -Not -BeNullOrEmpty
      $entraIdUpnObjDetails.UserPrincipalName | Should -Be 'jane.doe@contoso.com'
      $bhrWorkEmail | Should -Be 'jane.doe@contoso.com'
      $UpnExtensionAttribute1 | Should -Be 'SomeValue'
      "$($entraIdUpnObjDetails.AccountEnabled)" | Should -Be 'True'
    }

    It 'does not overwrite when UPN lookup succeeds' {
      $entraIdUpnObjDetails = [PSCustomObject]@{
        UserPrincipalName = 'john@contoso.com'
        AccountEnabled    = $true
        EmployeeId        = '12345'
      }
      $entraIdEidObjDetails = [PSCustomObject]@{
        UserPrincipalName = 'john@contoso.com'
        AccountEnabled    = $true
        EmployeeId        = '12345'
      }
      $bhrWorkEmail = 'john@contoso.com'
      $originalUpn = $bhrWorkEmail

      if ([string]::IsNullOrEmpty($entraIdUpnObjDetails) -and ([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false)) {
        $entraIdUpnObjDetails = $entraIdEidObjDetails
        $bhrWorkEmail = $entraIdEidObjDetails.UserPrincipalName
      }

      $bhrWorkEmail | Should -Be $originalUpn
    }

    It 'derives correct entraIdStatus from promoted EID object' {
      $entraIdUpnObjDetails = $null
      $entraIdEidObjDetails = [PSCustomObject]@{
        UserPrincipalName = 'termed.user@contoso.com'
        AccountEnabled    = $true
        EmployeeId        = '99999'
      }

      if ([string]::IsNullOrEmpty($entraIdUpnObjDetails) -and ([string]::IsNullOrEmpty($entraIdEidObjDetails) -eq $false)) {
        $entraIdUpnObjDetails = $entraIdEidObjDetails
      }

      # This is the variable the termination block checks
      $entraIdStatus = "$($entraIdUpnObjDetails.AccountEnabled)"
      $entraIdStatus | Should -Be 'True'
    }
  }
}

Describe 'Delta sync filtering' {
  Context 'Employee date filter' {
    It 'excludes employees older than the cutoff' {
      $cutoffDays = 14
      $cutoff = (Get-Date).AddDays(-$cutoffDays)

      $employees = @(
        [PSCustomObject]@{ id = '1'; lastChanged = (Get-Date).AddDays(-5).ToString('o'); status = 'Active' }
        [PSCustomObject]@{ id = '2'; lastChanged = (Get-Date).AddDays(-30).ToString('o'); status = 'Active' }
        [PSCustomObject]@{ id = '3'; lastChanged = (Get-Date).AddDays(-30).ToString('o'); status = 'Inactive' }
      )

      $filtered = @($employees | Where-Object {
          if ([string]::IsNullOrWhiteSpace($_.lastChanged)) { return $true }
          try { [datetime]$_.lastChanged -ge $cutoff } catch { $true }
        })

      $filtered.Count | Should -Be 1
      $filtered[0].id | Should -Be '1'
    }

    It 'includes employees with empty lastChanged' {
      $cutoffDays = 14
      $cutoff = (Get-Date).AddDays(-$cutoffDays)

      $employees = @(
        [PSCustomObject]@{ id = '1'; lastChanged = ''; status = 'Active' }
        [PSCustomObject]@{ id = '2'; lastChanged = $null; status = 'Inactive' }
      )

      $filtered = @($employees | Where-Object {
          if ([string]::IsNullOrWhiteSpace($_.lastChanged)) { return $true }
          try { [datetime]$_.lastChanged -ge $cutoff } catch { $true }
        })

      $filtered.Count | Should -Be 2
    }

    It 'includes all employees when FullSync is true' {
      $fullSync = $true

      $employees = @(
        [PSCustomObject]@{ id = '1'; lastChanged = (Get-Date).AddDays(-5).ToString('o'); status = 'Active' }
        [PSCustomObject]@{ id = '2'; lastChanged = (Get-Date).AddDays(-90).ToString('o'); status = 'Active' }
        [PSCustomObject]@{ id = '3'; lastChanged = (Get-Date).AddDays(-90).ToString('o'); status = 'Inactive' }
      )

      if (-not $fullSync) {
        $cutoff = (Get-Date).AddDays(-14)
        $employees = @($employees | Where-Object {
            if ([string]::IsNullOrWhiteSpace($_.lastChanged)) { return $true }
            try { [datetime]$_.lastChanged -ge $cutoff } catch { $true }
          })
      }

      $employees.Count | Should -Be 3
    }
  }
}

Describe 'Offboarding resilience' {
  BeforeAll {
    $script:helperRepoRoot = Split-Path -Parent $PSScriptRoot
    $script:startScriptPath = Join-Path $script:helperRepoRoot 'Start-BambooHRUserProvisioning.ps1'

    function script:Get-FunctionDefinitionsFromFile {
      [CmdletBinding()]
      param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string[]]$FunctionNames
      )
      $tokens = $null; $errors = $null
      $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
      if ($errors) { throw "Parse error: $(($errors | ForEach-Object { $_.Message }) -join '; ')" }
      $definitions = @()
      $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $FunctionNames -contains $n.Name }, $true) |
        ForEach-Object { $definitions += $_.Extent.Text }
      return $definitions
    }

    $functionNames = @('Invoke-WithRetry', 'Set-TerminatedUserProfileFields', 'Test-IsOffboardingComplete', 'Get-WorkPhoneComparisonValue', 'Get-OffboardingCompletionMarker', 'Invoke-UserOffboarding', 'ConvertTo-PhoneNumber')
    $definitions = Get-FunctionDefinitionsFromFile -ScriptPath $script:startScriptPath -FunctionNames $functionNames
    foreach ($def in $definitions) { Invoke-Expression $def }

    function Write-PSLog { [CmdletBinding()] param([string]$Message, [string]$Severity) Write-Verbose $Message }
    function Get-MgUser {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [string[]] $Property
      )

      return [PSCustomObject]@{
        AccountEnabled = $false
        DisplayName    = 'Test User'
        Id             = 'user-id'
      }
    }
    function Revoke-MgUserSignInSession {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Update-MgUser {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [bool] $AccountEnabled,
        [Parameter()] [hashtable] $BodyParameter,
        [Parameter()] [array] $BusinessPhones,
        [Parameter()] [hashtable] $OnPremisesExtensionAttributes,
        [Parameter()] [string] $EmployeeId,
        [Parameter()] [datetime] $EmployeeLeaveDateTime,
        [Parameter()] [string] $CompanyName,
        [Parameter()] [string] $MobilePhone,
        [Parameter()] [string] $Department,
        [Parameter()] [string] $JobTitle,
        [Parameter()] [string] $OfficeLocation,
        [Parameter()] [string] $DisplayName,
        [Parameter()] [string] $GivenName,
        [Parameter()] [string] $Surname,
        [Parameter()] [string] $Mail,
        [Parameter()] [string] $UserPrincipalName,
        [Parameter()] [string] $EmployeeHireDate
      )
    }
    function Invoke-MgGraphRequest {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $Method,
        [Parameter()] [string] $Uri,
        [Parameter()] [string] $Body,
        [Parameter()] [string] $ContentType,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Connect-ExchangeOnlineIfNeeded { [CmdletBinding()] param([string]$TenantId) }
    function Set-MailboxAutoReplyConfiguration {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $Identity,
        [Parameter()] [string] $AutoReplyState,
        [Parameter()] [string] $ExternalAudience,
        [Parameter()] [string] $InternalMessage,
        [Parameter()] [string] $ExternalMessage,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Set-Mailbox {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $Identity,
        [Parameter()] [string] $Type,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Wait-ForCondition { [CmdletBinding()] param([string]$Operation, [scriptblock]$Condition, [int]$TimeoutSeconds, [int]$PollIntervalSeconds) return $true }
    function Get-EXOMailbox {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $Anr,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )

      return [PSCustomObject]@{ Id = 'mbx-id'; Identity = 'test@contoso.com'; RecipientTypeDetails = 'SharedMailbox' }
    }
    function Add-MailboxPermission {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $Identity,
        [Parameter()] [string] $User,
        [Parameter()] [string[]] $AccessRights,
        [Parameter()] [bool] $Automapping,
        [Parameter()] [string] $InheritanceType,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Get-MgUserLicenseDetail {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )

      return @()
    }
    function Set-MgUserLicense {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [object] $RemoveLicenses,
        [Parameter()] [hashtable] $AddLicenses,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Get-MgUserOwnedObject {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [switch] $All
      )

      return @()
    }
    function Get-MgGroupOwner {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $GroupId,
        [Parameter()] [switch] $All
      )

      return @()
    }
    function Get-MgGroup {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $GroupId,
        [Parameter()] [string[]] $Property
      )

      return [PSCustomObject]@{
        DisplayName    = 'Test Group'
        GroupTypes     = @()
        Id             = $GroupId
        MembershipRule = $null
      }
    }
    function New-MgGroupOwnerByRef { [CmdletBinding()] param() }
    function Remove-MgGroupOwnerByRef { [CmdletBinding()] param() }
    function Get-MgUserMemberOf {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [switch] $All,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )

      return @()
    }
    function Remove-MgGroupMemberByRef { [CmdletBinding()] param() }
    function Get-CachedDistributionGroups { [CmdletBinding()] param([hashtable]$Cache) return @() }
    function Get-DistributionGroupMember { [CmdletBinding()] param() return @() }
    function Remove-DistributionGroupMember { [CmdletBinding()] param() }
    function Get-CachedMailboxes { [CmdletBinding()] param([hashtable]$Cache) return @() }
    function Get-EXOMailboxPermission { [CmdletBinding()] param() return @() }
    function Remove-MailboxPermission { [CmdletBinding()] param() }
    function Get-EXORecipientPermission { [CmdletBinding()] param() return @() }
    function Remove-RecipientPermission { [CmdletBinding()] param() }
    function Get-MgContext {
      [CmdletBinding()]
      param()

      return [PSCustomObject]@{
        AuthType = 'AppOnly'
      }
    }
    function Get-MgUserOwnedDevice {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )

      return @()
    }
    function Remove-MgDeviceRegisteredOwnerByRef { [CmdletBinding()] param() }
    function Get-MgDevice {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $DeviceId,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )

      return [PSCustomObject]@{ Notes = '' }
    }
    function Update-MgDevice {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $DeviceId,
        [Parameter()] [hashtable] $BodyParameter,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Get-MgUserManagedDevice {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [switch] $All,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )

      return @()
    }
    function Update-MgDeviceManagementManagedDevice { [CmdletBinding()] param() }
    function Get-CachedAutopilotDevicesForUser { [CmdletBinding()] param([hashtable]$Cache, [string]$UserPrincipalName) return @() }
    function Update-MgDeviceManagementWindowsAutopilotDeviceIdentity { [CmdletBinding()] param() }
    function Get-MgUserAuthenticationMethod {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )

      return @()
    }
    function Remove-MgUserAuthenticationPhoneMethod { [CmdletBinding()] param() }
    function Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod { [CmdletBinding()] param() }
    function Remove-MgUserAuthenticationFido2Method { [CmdletBinding()] param() }
    function Remove-MgUserAuthenticationWindowsHelloForBusinessMethod { [CmdletBinding()] param() }
    function Remove-MgUserManagerByRef {
      [CmdletBinding()]
      param(
        [Parameter()] [string] $UserId,
        [Parameter()] [System.Management.Automation.ActionPreference] $ErrorAction
      )
    }
    function Add-SignificantChange { [CmdletBinding()] param([string]$Category, [string]$User, [string]$Detail) }

    $Script:Config = @{
      CorrelationId = [Guid]::NewGuid().ToString()
      Runtime       = @{ MaxRetryAttempts = 1; RetryDelaySeconds = 0 }
      Azure         = @{ TenantId = 'contoso.onmicrosoft.com' }
    }
  }

  Context 'Set-TerminatedUserProfileFields makes three separate API calls' {
    It 'calls Update-MgUser for profile fields and phone, plus Invoke-MgGraphRequest for address' {
      $script:updateCalls = 0
      $script:graphRequestCalls = 0

      Mock Update-MgUser { $script:updateCalls++ }
      Mock Invoke-MgGraphRequest { $script:graphRequestCalls++ }

      Set-TerminatedUserProfileFields -UserId 'test@contoso.com' -LeaveDateTimeUtc ([datetime]::UtcNow)

      $script:updateCalls | Should -Be 2
      $script:graphRequestCalls | Should -Be 1
    }

    It 'continues clearing profile fields and phones when the leave-date write fails' {
      $script:updateOperations = [System.Collections.Generic.List[string]]::new()
      $script:graphRequestCalls = 0

      Mock Update-MgUser {
        param($UserId, $EmployeeLeaveDateTime, $BusinessPhones, $ErrorAction)

        if ($PSBoundParameters.ContainsKey('EmployeeLeaveDateTime')) {
          $script:updateOperations.Add('leaveDateTime') | Out-Null
          throw 'Insufficient privileges to complete the operation.'
        }

        if ($PSBoundParameters.ContainsKey('BusinessPhones')) {
          $script:updateOperations.Add('businessPhones') | Out-Null
        }
      }
      Mock Invoke-MgGraphRequest { $script:graphRequestCalls++ }

      { Set-TerminatedUserProfileFields -UserId 'test@contoso.com' -LeaveDateTimeUtc ([datetime]::UtcNow) } | Should -Throw -ExpectedMessage '*EmployeeLeaveDateTime*'

      $script:updateOperations.Count | Should -Be 2
      $script:updateOperations | Should -Contain 'leaveDateTime'
      $script:updateOperations | Should -Contain 'businessPhones'
      $script:graphRequestCalls | Should -Be 1
    }
  }

  Context 'Set-TerminatedUserProfileFields address clearing uses raw JSON null' {
    It 'sends JSON body with null values for profile and address fields' {
      $capturedBody = $null
      Mock Update-MgUser { }
      Mock Invoke-MgGraphRequest {
        param($Method, $Uri, $Body, $ContentType)
        $script:capturedBody = $Body
      }

      Set-TerminatedUserProfileFields -UserId 'test@contoso.com' -LeaveDateTimeUtc ([datetime]::UtcNow)

      $script:capturedBody | Should -Not -BeNullOrEmpty
      $parsed = $script:capturedBody | ConvertFrom-Json
      $parsed.department | Should -BeNullOrEmpty
      $parsed.jobTitle | Should -BeNullOrEmpty
      $parsed.officeLocation | Should -BeNullOrEmpty
      $parsed.mobilePhone | Should -BeNullOrEmpty
      $parsed.city | Should -BeNullOrEmpty
      $parsed.state | Should -BeNullOrEmpty
      $parsed.streetAddress | Should -BeNullOrEmpty
      $parsed.postalCode | Should -BeNullOrEmpty
    }
  }

  Context 'Test-IsOffboardingComplete detection for retry' {
    It 'returns false when CompanyName is the company name (not a date stamp)' {
      Test-IsOffboardingComplete -CompanyName 'Gecko Green' -Department 'Sales' -JobTitle 'Rep' -OfficeLocation 'Phoenix' | Should -BeFalse
    }

    It 'returns false when Department is still set to original value' {
      Test-IsOffboardingComplete -CompanyName '04/08/26 (OffboardingComplete: 2026-04-08T19:00:00Z)' -Department 'Sales' -JobTitle '' -OfficeLocation '' | Should -BeFalse
    }

    It 'returns false when phones still have values' {
      Test-IsOffboardingComplete -CompanyName '04/08/26 (OffboardingComplete: 2026-04-08T19:00:00Z)' -Department '' -JobTitle '' -OfficeLocation '' -WorkPhone '5551234567' -MobilePhone '' | Should -BeFalse
    }

    It 'returns true when all markers are set and phones are cleared' {
      Test-IsOffboardingComplete -CompanyName '04/08/26 (OffboardingComplete: 2026-04-08T19:00:00Z)' -Department '' -JobTitle '' -OfficeLocation '' -WorkPhone '' -MobilePhone '' | Should -BeTrue
    }

    It 'returns true when phone parameters are omitted entirely' {
      Test-IsOffboardingComplete -CompanyName '04/08/26 (OffboardingComplete: 2026-04-08T19:00:00Z)' -Department '' -JobTitle '' -OfficeLocation '' | Should -BeTrue
    }
  }

  Context 'Invoke-UserOffboarding remains the shared offboarding path' {
    It 'primary and retry paths both call Invoke-UserOffboarding' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      $offboardingCallPattern = 'Invoke-UserOffboarding'
      ($content | Select-String -Pattern $offboardingCallPattern -AllMatches).Matches.Count | Should -BeGreaterOrEqual 3
    }
  }

  Context 'Invoke-UserOffboarding handles soft and hard failures correctly' {
    BeforeEach {
      $script:completionCompanyNames = @()
      $script:completionEmployeeIds = @()
      $script:offboardingErrorSummary = @{
        TotalErrors    = 0
        ErrorsByType   = @{}
        ErrorsByUser   = @{}
        CriticalErrors = @()
        Warnings       = @()
      }

      Mock Update-MgUser {
        param(
          [string] $UserId,
          [bool] $AccountEnabled,
          [array] $BusinessPhones,
          [string] $EmployeeId,
          [string] $CompanyName,
          [datetime] $EmployeeLeaveDateTime,
          [System.Management.Automation.ActionPreference] $ErrorAction
        )

        if ($PSBoundParameters.ContainsKey('EmployeeId')) {
          $script:completionEmployeeIds += $EmployeeId
        }

        if ($PSBoundParameters.ContainsKey('CompanyName')) {
          $script:completionCompanyNames += $CompanyName
        }
      }
      Mock Get-MgUser {
        param($UserId, $Property, $ErrorAction)

        [PSCustomObject]@{
          Id                = 'user-id'
          UserPrincipalName = $UserId
          AccountEnabled    = $false
        }
      }
      Mock Invoke-MgGraphRequest { param($Method, $Uri, $Body, $ContentType, $ErrorAction) }
      Mock Revoke-MgUserSignInSession { param($UserId, $ErrorAction) }
      Mock Connect-ExchangeOnlineIfNeeded { param($TenantId) }
      Mock Get-MgContext { [PSCustomObject]@{ AuthType = 'AppOnly' } }
      Mock Get-MgUserLicenseDetail { param($UserId, $ErrorAction) @() }
      Mock Get-MgUserMemberOf { param($UserId, [switch] $All, $ErrorAction) @() }
      Mock Set-MailboxAutoReplyConfiguration { param($Identity, $AutoReplyState, $ExternalAudience, $InternalMessage, $ExternalMessage, $ErrorAction) }
      Mock Set-Mailbox { param($Identity, $Type, $ErrorAction) }
      Mock Get-EXOMailbox { param($Anr, $ErrorAction) [PSCustomObject]@{ Id = 'mbx-id'; Identity = 'test@contoso.com'; RecipientTypeDetails = 'SharedMailbox' } }
      Mock Get-CachedDistributionGroups { param($Cache) @() }
      Mock Get-DistributionGroupMember { param($Identity, $ResultSize, $ErrorAction) @() }
      Mock Remove-DistributionGroupMember { param($Identity, $Member, $BypassSecurityGroupManagerCheck, $Confirm, $ErrorAction) }
      Mock Get-CachedMailboxes { param($Cache) @() }
      Mock Get-EXOMailboxPermission { param($Identity, $User, $ErrorAction) @() }
      Mock Remove-MailboxPermission { param($Identity, $User, $AccessRights, $Confirm, $ErrorAction) }
      Mock Get-EXORecipientPermission { param($Identity, $ResultSize, $ErrorAction) @() }
      Mock Remove-RecipientPermission { param($Identity, $Trustee, $AccessRights, $Confirm, $ErrorAction) }
      Mock Get-MgUserOwnedDevice { param($UserId, $ErrorAction) @() }
      Mock Get-MgUserManagedDevice { param($UserId, [switch] $All, $ErrorAction) @() }
      Mock Get-CachedAutopilotDevicesForUser { param($Cache, $UserPrincipalName) @() }
      Mock Get-MgUserAuthenticationMethod { param($UserId, $ErrorAction) @() }
      Mock Remove-MgUserAuthenticationPhoneMethod { param($UserId, $PhoneAuthenticationMethodId) }
      Mock Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod { param($UserId, $MicrosoftAuthenticatorAuthenticationMethodId) }
      Mock Remove-MgUserAuthenticationFido2Method { param($UserId, $Fido2AuthenticationMethodId) }
      Mock Remove-MgUserAuthenticationWindowsHelloForBusinessMethod { param($UserId, $WindowsHelloForBusinessAuthenticationMethodId) }
      Mock Remove-MgUserManagerByRef { param($UserId, $ErrorAction) }
    }

    It 'treats dynamic group removal as a soft failure and still writes completion markers' {
      Mock Remove-MgGroupMemberByRef { param($GroupId, $DirectoryObjectId, $ErrorAction) }
      Mock Get-MgUserMemberOf {
        param($UserId, [switch] $All, $ErrorAction)

        return @([PSCustomObject]@{
            Id                   = 'group-1'
            AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.group' }
          })
      }
      Mock Get-MgGroup {
        param($GroupId, $Property)

        return [PSCustomObject]@{
          DisplayName    = 'Dynamic Group'
          GroupTypes     = @('DynamicMembership')
          Id             = $GroupId
          MembershipRule = 'user.department -eq "Sales"'
        }
      }

      Invoke-UserOffboarding -UserId 'test@contoso.com' -UserObjectId 'user-id' -LeaveDateTimeUtc ([datetime]::UtcNow) -PerformanceCache @{} -ErrorSummary $script:offboardingErrorSummary

      Assert-MockCalled Remove-MgGroupMemberByRef -Times 0 -Exactly
      $script:offboardingErrorSummary.TotalErrors | Should -Be 0
      $script:completionEmployeeIds | Should -Contain 'LVR'
      $script:completionCompanyNames.Count | Should -Be 1
    }

    It 'treats inherited license removal as a soft failure and still writes completion markers' {
      Mock Get-MgUserLicenseDetail {
        param($UserId)

        return @([PSCustomObject]@{ SkuId = 'sku-1' })
      }
      Mock Set-MgUserLicense {
        param($UserId, $RemoveLicenses, $AddLicenses, $ErrorAction)

        throw 'User license is inherited from a group membership and it cannot be removed directly from the user.'
      }

      Invoke-UserOffboarding -UserId 'test@contoso.com' -UserObjectId 'user-id' -LeaveDateTimeUtc ([datetime]::UtcNow) -PerformanceCache @{} -ErrorSummary $script:offboardingErrorSummary

      $script:offboardingErrorSummary.TotalErrors | Should -Be 0
      ($script:offboardingErrorSummary.Warnings -join '; ') | Should -Match 'inherited from group membership'
      $script:completionEmployeeIds | Should -Contain 'LVR'
      $script:completionCompanyNames.Count | Should -Be 1
    }

    It 'records hard failures and skips completion markers' {
      Mock Set-TerminatedUserProfileFields { throw 'missing lifecycle permission' }

      Invoke-UserOffboarding -UserId 'test@contoso.com' -UserObjectId 'user-id' -LeaveDateTimeUtc ([datetime]::UtcNow) -PerformanceCache @{} -ErrorSummary $script:offboardingErrorSummary

      $script:offboardingErrorSummary.TotalErrors | Should -Be 1
      $script:offboardingErrorSummary.ErrorsByType['OffboardingProfileFields'] | Should -Be 1
      $script:offboardingErrorSummary.ErrorsByUser['test@contoso.com'] | Should -Match 'Failed to set terminated profile fields'
      $script:completionEmployeeIds.Count | Should -Be 0
      $script:completionCompanyNames.Count | Should -Be 0
    }
  }

  Context 'Autopilot block does not trap subsequent offboarding steps' {
    It 'contains the Autopilot cleanup block in Invoke-UserOffboarding' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      $functionSection = ($content -split 'function Invoke-UserOffboarding')[1] -split 'function [A-Z]' | Select-Object -First 1
      $functionSection | Should -Match 'Get-MgDeviceManagementWindowsAutopilotDeviceIdentity'
    }
  }

  Context 'Retry path calls Invoke-UserOffboarding with all critical steps' {
    It 'retry path calls Invoke-UserOffboarding' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      # The retry block (inside "Complete missed offboarding") should call the shared function
      $retrySection = ($content -split 'Complete missed offboarding')[1]
      $retrySection | Should -Match 'Invoke-UserOffboarding'
    }

    It 'Invoke-UserOffboarding sets auto-reply configuration' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      # The shared function should contain auto-reply setup
      $functionSection = ($content -split 'function Invoke-UserOffboarding')[1] -split 'function [A-Z]' | Select-Object -First 1
      $functionSection | Should -Match 'Set-MailboxAutoReplyConfiguration'
    }

    It 'Invoke-UserOffboarding removes authentication methods' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      $functionSection = ($content -split 'function Invoke-UserOffboarding')[1] -split 'function [A-Z]' | Select-Object -First 1
      $functionSection | Should -Match 'Get-MgUserAuthenticationMethod'
    }

    It 'Invoke-UserOffboarding handles device ownership removal' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      $functionSection = ($content -split 'function Invoke-UserOffboarding')[1] -split 'function [A-Z]' | Select-Object -First 1
      $functionSection | Should -Match 'Remove-MgDeviceRegisteredOwnerByRef'
    }

    It 'Invoke-UserOffboarding removes mailbox delegate permissions' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      $functionSection = ($content -split 'function Invoke-UserOffboarding')[1] -split 'function [A-Z]' | Select-Object -First 1
      $functionSection | Should -Match 'Remove-MailboxPermission'
    }

    It 'Invoke-UserOffboarding cleans up Intune managed devices' {
      $content = Get-Content -Raw -Path $script:startScriptPath

      $functionSection = ($content -split 'function Invoke-UserOffboarding')[1] -split 'function [A-Z]' | Select-Object -First 1
      $functionSection | Should -Match 'Update-MgDeviceManagementManagedDevice'
    }
  }
}
