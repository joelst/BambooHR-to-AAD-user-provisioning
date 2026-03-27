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
$startAltScriptPath = Join-Path $repoRoot '_Start-BambooHRUserProvisioning.ps1'

Describe 'Static validation' {
  BeforeAll {
    $script:staticRepoRoot = Split-Path -Parent $PSScriptRoot
    $script:staticStartScriptPath = Join-Path $script:staticRepoRoot 'Start-BambooHRUserProvisioning.ps1'
    $script:staticStartAltScriptPath = Join-Path $script:staticRepoRoot '_Start-BambooHRUserProvisioning.ps1'
  }

  It 'does not contain invalid variable references' {
    $pattern = '\$(?!script:|Script:|env:|global:|local:|private:|using:)[A-Za-z_][A-Za-z0-9_]*:'
    $content = Get-Content -Raw -Path $script:staticStartScriptPath
    $contentAlt = Get-Content -Raw -Path $script:staticStartAltScriptPath

    $content | Should -Not -Match $pattern
    $contentAlt | Should -Not -Match $pattern
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
      'ConvertTo-StandardName',
      'ConvertTo-PhoneNumber',
      'ConvertTo-BambooHrHireDate',
      'Get-ValidManagerUser',
      'Test-IsTenantEmail',
      'Get-MailNicknameFromEmail',
      'Get-WorkPhoneComparisonValue',
      'Test-ShouldSyncExistingUser',
      'Test-ShouldUpdateEmployeeId',
      'Test-IsOffboardingComplete',
      'Get-OffboardingCompletionMarker',
      'Get-OffboardingCompletionDateFromCompanyName'
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
        -BhrEmployeeNumber '234' `
        -EntraIdUserPrincipalName 'rmirouh@geckogreen.com' `
        -BhrWorkEmail 'rmirouh@geckogreen.com' `
        -BhrEmploymentStatus 'Terminated' `
        -BhrAccountEnabled $false `
        -BhrLastChanged '2026-03-26T05:34:03Z' `
        -UpnExtensionAttribute1 '2026-03-26T02:52:09Z' | Should -BeFalse
    }
  }

  Context 'Offboarding completion helpers' {
    It 'detects completed offboarding markers' {
      Test-IsOffboardingComplete -CompanyName '03/26/26 (OffboardingComplete: 2026-03-26T04:00:00Z)' -Department 'Not Active' -JobTitle 'Not Active' -OfficeLocation 'Not Active' -WorkPhone '' -MobilePhone '' | Should -BeTrue
    }

    It 'treats remaining phone values as incomplete offboarding' {
      Test-IsOffboardingComplete -CompanyName '03/26/26 (OffboardingComplete: 2026-03-26T04:00:00Z)' -Department 'Not Active' -JobTitle 'Not Active' -OfficeLocation 'Not Active' -WorkPhone '' -MobilePhone '5551234567' | Should -BeFalse
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
