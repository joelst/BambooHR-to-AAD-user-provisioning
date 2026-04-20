BeforeAll {
  $script:webhookScriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Start-BambooHrWebhookSync.ps1'

  function script:Get-FunctionDefinitionsFromFile {
    [CmdletBinding()]
    param(
      [Parameter(Mandatory = $true)]
      [string]
      $ScriptPath,

      [Parameter(Mandatory = $true)]
      [string[]]
      $FunctionNames
    )

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
    if ($errors) {
      throw "Parse error: $(($errors | ForEach-Object { $_.Message }) -join '; ')"
    }

    $definitions = @()
    $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $FunctionNames -contains $node.Name
      }, $true) | ForEach-Object {
      $definitions += $_.Extent.Text
    }

    return $definitions
  }

  $functionNames = @(
    'Get-TargetEmployeeIdsFromWebhookData',
    'Get-WebhookChangedFieldAnalysis',
    'Test-BambooHrWebhookSignature',
    'Test-WebhookChangedFieldsRelevant'
  )
  $definitions = Get-FunctionDefinitionsFromFile -ScriptPath $script:webhookScriptPath -FunctionNames $functionNames
  foreach ($definition in $definitions) {
    Invoke-Expression $definition
  }

  function Write-PSLog {
    [CmdletBinding()]
    param(
      [string]$Message,
      [string]$Severity
    )
  }
}

Describe 'Start-BambooHrWebhookSync helper functions' {
  It 'merges and deduplicates employee IDs from manual input and webhook JSON' {
    $webhookData = [PSCustomObject]@{
      RequestBody = @'
{
  "employees": [
    { "id": "1001", "action": "Updated" },
    { "id": "1002", "action": "Created" },
    { "id": "1001", "action": "Updated" }
  ]
}
'@
    }

    $result = Get-TargetEmployeeIdsFromWebhookData -InputIds @('1003', '1002') -WebhookData $webhookData

    $result | Should -Be @('1001', '1002', '1003')
  }

  It 'extracts employeeId from the real BambooHR webhook payload format' {
    $webhookData = [PSCustomObject]@{
      RequestBody = @'
{
  "type": "employee.updated",
  "timestamp": "2026-04-10T10:28:35Z",
  "data": {
    "changedFields": ["mobilePhone"],
    "companyId": "73142",
    "employeeId": "2145"
  }
}
'@
    }

    $result = Get-TargetEmployeeIdsFromWebhookData -InputIds @() -WebhookData $webhookData

    $result | Should -Be @('2145')
  }

  It 'merges real BambooHR payload employeeId with manual input IDs' {
    $webhookData = [PSCustomObject]@{
      RequestBody = @'
{
  "type": "employee.updated",
  "timestamp": "2026-04-10T10:28:35Z",
  "data": {
    "changedFields": ["workEmail"],
    "companyId": "73142",
    "employeeId": "2145"
  }
}
'@
    }

    $result = Get-TargetEmployeeIdsFromWebhookData -InputIds @('2145', '3001') -WebhookData $webhookData

    $result | Should -Be @('2145', '3001')
  }

  It 'returns manual IDs when webhook data is absent' {
    $result = Get-TargetEmployeeIdsFromWebhookData -InputIds @('2002', '2001', '2002') -WebhookData $null

    $result | Should -Be @('2001', '2002')
  }

  It 'validates a BambooHR signature when the headers and private key match' {
    $body = '{"employees":[{"id":"3001","action":"Updated"}]}'
    $timestamp = '2026-04-09T04:00:00Z'
    $privateKey = 'super-secret-key'

    $hmac = [System.Security.Cryptography.HMACSHA256]::new([System.Text.Encoding]::UTF8.GetBytes($privateKey))
    try {
      $signature = [System.Convert]::ToHexString(
        $hmac.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($body + $timestamp))
      ).ToLowerInvariant()
    }
    finally {
      $hmac.Dispose()
    }

    $webhookData = [PSCustomObject]@{
      RequestBody   = $body
      RequestHeader = [PSCustomObject]@{
        'X-BambooHR-Timestamp' = $timestamp
        'X-BambooHR-Signature' = $signature
      }
    }

    Test-BambooHrWebhookSignature -WebhookData $webhookData -PrivateKey $privateKey | Should -BeTrue
  }

  It 'returns false when the BambooHR signature does not match' {
    $webhookData = [PSCustomObject]@{
      RequestBody   = '{"employees":[{"id":"4001","action":"Updated"}]}'
      RequestHeader = [PSCustomObject]@{
        'X-BambooHR-Timestamp' = '2026-04-09T04:00:00Z'
        'X-BambooHR-Signature' = 'notavalidsignature'
      }
    }

    Test-BambooHrWebhookSignature -WebhookData $webhookData -PrivateKey 'super-secret-key' | Should -BeFalse
  }

  Context 'Test-WebhookChangedFieldsRelevant' {
    It 'returns true when changedFields contain a synced field like workEmail' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.updated",
  "data": {
    "changedFields": ["workEmail", "shirtSize"],
    "employeeId": "2145"
  }
}
'@
      }

      Test-WebhookChangedFieldsRelevant -WebhookData $webhookData | Should -BeTrue
    }

    It 'returns false when changedFields contain only non-synced fields like shirtSize and ssn' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.updated",
  "data": {
    "changedFields": ["shirtSize", "ssn"],
    "employeeId": "2145"
  }
}
'@
      }

      Test-WebhookChangedFieldsRelevant -WebhookData $webhookData | Should -BeFalse
    }

    It 'returns true for employee.created events regardless of changedFields' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.created",
  "data": {
    "changedFields": [],
    "employeeId": "9999"
  }
}
'@
      }

      Test-WebhookChangedFieldsRelevant -WebhookData $webhookData | Should -BeTrue
    }

    It 'returns true when changedFields is missing from the payload' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.updated",
  "data": {
    "employeeId": "2145"
  }
}
'@
      }

      Test-WebhookChangedFieldsRelevant -WebhookData $webhookData | Should -BeTrue
    }

    It 'returns true when WebhookData is null' {
      Test-WebhookChangedFieldsRelevant -WebhookData $null | Should -BeTrue
    }

    It 'returns true when changedFields includes photoUploaded' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.updated",
  "data": {
    "changedFields": ["photoUploaded"],
    "employeeId": "2145"
  }
}
'@
      }

      Test-WebhookChangedFieldsRelevant -WebhookData $webhookData | Should -BeTrue
    }
  }

  Context 'Get-WebhookChangedFieldAnalysis' {
    It 'classifies unsynced-only updates without stopping targeted processing metadata' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.updated",
  "data": {
    "changedFields": ["shirtSize", "ssn"],
    "employeeId": "2145"
  }
}
'@
      }

      $analysis = Get-WebhookChangedFieldAnalysis -WebhookData $webhookData

      $analysis.EventType | Should -Be 'employee.updated'
      $analysis.EmployeeId | Should -Be '2145'
      $analysis.SyncedFields | Should -BeNullOrEmpty
      $analysis.UnsyncedFields | Should -Be @('shirtSize', 'ssn')
      $analysis.HasOnlyUnsyncedFields | Should -BeTrue
      $analysis.ShouldTreatAsRelevant | Should -BeFalse
    }

    It 'separates synced and unsynced fields for mixed updates' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.updated",
  "data": {
    "changedFields": ["workEmail", "shirtSize"],
    "employeeId": "2145"
  }
}
'@
      }

      $analysis = Get-WebhookChangedFieldAnalysis -WebhookData $webhookData

      $analysis.SyncedFields | Should -Be @('workEmail')
      $analysis.UnsyncedFields | Should -Be @('shirtSize')
      $analysis.HasOnlyUnsyncedFields | Should -BeFalse
      $analysis.ShouldTreatAsRelevant | Should -BeTrue
    }

    It 'treats missing changedFields as processable' {
      $webhookData = [PSCustomObject]@{
        RequestBody = @'
{
  "type": "employee.updated",
  "data": {
    "employeeId": "2145"
  }
}
'@
      }

      $analysis = Get-WebhookChangedFieldAnalysis -WebhookData $webhookData

      $analysis.ChangedFields | Should -BeNullOrEmpty
      $analysis.HasOnlyUnsyncedFields | Should -BeFalse
      $analysis.ShouldTreatAsRelevant | Should -BeTrue
    }
  }
}
