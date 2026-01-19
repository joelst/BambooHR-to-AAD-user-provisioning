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
  $script:MaxParallelUsers = 5
  $script:BatchSize = 25
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
  It 'does not contain invalid variable references' {
    $pattern = '\$(?!script:|Script:|env:|global:|local:|private:|using:)[A-Za-z_][A-Za-z0-9_]*:'
    $content = Get-Content -Raw -Path $startScriptPath
    $contentAlt = Get-Content -Raw -Path $startAltScriptPath

    $content | Should -Not -Match $pattern
    $contentAlt | Should -Not -Match $pattern
  }
}

Describe 'Start-BambooHRUserProvisioning helpers' {
  BeforeAll {
    $functionNames = @(
      'Initialize-Configuration',
      'ConvertTo-StandardName',
      'ConvertTo-PhoneNumber',
      'Get-ValidManagerUser'
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
