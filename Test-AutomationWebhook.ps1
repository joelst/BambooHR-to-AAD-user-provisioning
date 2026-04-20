<#
.SYNOPSIS
    Triggers an Azure Automation webhook with a BambooHR-style payload for testing the webhook sync runbook.

.DESCRIPTION
    Sends an HTTP POST to an Azure Automation webhook URL with a JSON body matching the
    real BambooHR webhook format expected by Start-BambooHrWebhookSync.ps1.

    BambooHR sends one event per changed employee with the structure:
      { "type": "employee.updated", "timestamp": "...", "data": { "employeeId": "...", ... } }

    When a BambooHR webhook private key is provided, the request includes HMAC-SHA256
    signature headers (X-BambooHR-Signature and X-BambooHR-Timestamp) so the runbook
    can validate authenticity. The timestamp is a Unix epoch integer, matching the real
    BambooHR behavior.

    IMPORTANT: The Azure Automation webhook URL is a bearer-secret. Treat it like a password.
    Never commit it to source control. Pass it as a parameter or store it in a secure vault.

.PARAMETER WebhookUri
    The Azure Automation webhook URL. This is the full URI shown once when the webhook is created.

.PARAMETER EmployeeId
    The BambooHR employee ID to include in the payload.

.PARAMETER ChangedFields
    One or more BambooHR field names that changed. Default is 'workEmail'.

.PARAMETER WebhookPrivateKey
    Optional BambooHR webhook private key. When provided, the request is signed with
    HMAC-SHA256 using the same algorithm that Start-BambooHrWebhookSync.ps1 validates.

.EXAMPLE
    .\Test-AutomationWebhook.ps1 -WebhookUri 'https://abc123.webhook.wus2.azure-automation.net/webhooks?token=secret' -EmployeeId '2145'

    Triggers the webhook for a single employee without HMAC signing.

.EXAMPLE
    .\Test-AutomationWebhook.ps1 -WebhookUri $uri -EmployeeId '2145' -ChangedFields 'mobilePhone', 'workEmail'

    Triggers the webhook indicating multiple fields changed.

.EXAMPLE
    .\Test-AutomationWebhook.ps1 -WebhookUri $uri -EmployeeId '2145' -WebhookPrivateKey 'my-private-key'

    Triggers the webhook with HMAC-SHA256 signature headers for validation.

.NOTES
    Requires PowerShell 7+.
    The webhook URL is a secret — never store it in source control.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter(Mandatory = $true, HelpMessage = 'Azure Automation webhook URL.')]
  [ValidateNotNullOrEmpty()]
  [ValidatePattern('^https://')]
  [string]
  $WebhookUri,

  [Parameter(Mandatory = $true, HelpMessage = 'BambooHR employee ID to include in the payload.')]
  [ValidateNotNullOrEmpty()]
  [string]
  $EmployeeId,

  [Parameter(HelpMessage = 'BambooHR field names that changed.')]
  [ValidateNotNullOrEmpty()]
  [string[]]
  $ChangedFields = @('workEmail'),

  [Parameter(HelpMessage = 'BambooHR webhook private key for HMAC-SHA256 signing.')]
  [string]
  $WebhookPrivateKey
)

# Build the real BambooHR webhook payload format
$payload = [PSCustomObject]@{
  type      = 'employee.updated'
  timestamp = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
  data      = [PSCustomObject]@{
    changedFields = @($ChangedFields)
    employeeId    = $EmployeeId
  }
} | ConvertTo-Json -Depth 5 -Compress

Write-Verbose "Payload: $payload"

# Build request headers
$headers = @{
  'Content-Type' = 'application/json'
}

# Sign the request when a private key is provided (BambooHR uses Unix epoch for the timestamp header)
if (-not [string]::IsNullOrWhiteSpace($WebhookPrivateKey)) {
  $timestamp = [string][long]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
  $hmac = [System.Security.Cryptography.HMACSHA256]::new(
    [System.Text.Encoding]::UTF8.GetBytes($WebhookPrivateKey)
  )
  try {
    $signatureBytes = $hmac.ComputeHash(
      [System.Text.Encoding]::UTF8.GetBytes($payload + $timestamp)
    )
    $signature = [System.Convert]::ToHexString($signatureBytes).ToLowerInvariant()
  }
  finally {
    $hmac.Dispose()
  }

  $headers['x-bamboohr-timestamp'] = $timestamp
  $headers['x-bamboohr-signature'] = $signature

  Write-Verbose "Signed request with timestamp $timestamp"
}

# Send the webhook
if ($PSCmdlet.ShouldProcess($WebhookUri, "POST BambooHR webhook payload for employee ID: $EmployeeId")) {
  $splatParams = @{
    Uri         = $WebhookUri
    Method      = 'POST'
    Headers     = $headers
    Body        = $payload
    ContentType = 'application/json'
  }

  try {
    $response = Invoke-RestMethod @splatParams
    Write-Output 'Webhook triggered successfully.'
    Write-Verbose "Response: $($response | ConvertTo-Json -Depth 5 -ErrorAction SilentlyContinue)"
    $response
  }
  catch {
    Write-Error "Failed to trigger webhook: $($_.Exception.Message)"
  }
}