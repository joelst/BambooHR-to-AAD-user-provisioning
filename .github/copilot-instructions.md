# Project Guidelines

## Overview

BambooHR-to-Entra-ID user provisioning script that runs as an **Azure Automation runbook**. Single-script PowerShell solution (`Start-BambooHRUserProvisioning.ps1`) with Pester tests in `tests/`.

See [README.md](../README.md) for the project overview, deployment entry point, and docs map.
See [docs/development.md](../docs/development.md) for architecture, script flow, testing, and change workflow.
See [docs/security.md](../docs/security.md) for the threat model and hardening expectations.
See [docs/troubleshooting.md](../docs/troubleshooting.md) for common operational failure modes.

## Runtime Environment

- **Execution context**: Azure Automation with system-assigned Managed Identity
- **PowerShell version**: 7+ (Azure Automation sandbox)
- **Authentication**: `Connect-AzAccount -Identity`, `Connect-MgGraph -Identity`, `Connect-ExchangeOnline -ManagedIdentity`
- **No interactive prompts** — the script runs unattended; never add `Read-Host`, interactive `Confirm`, or UI-dependent code
- **`ConfirmImpact = 'None'`** is required to prevent Azure Automation from hanging on implicit confirmation prompts
- **`[switch]` parameters do not work** in Azure Automation runbooks — use `[bool]` with a default value instead (e.g., `[bool] $FullSync = $false`)

## Code Style

- Use approved verbs (`Get-`, `Set-`, `New-`, `Remove-`, `Invoke-`, `Test-`, `ConvertTo-`)
- Parameters in `param()` block: required first, then alphabetical
- Use `[CmdletBinding()]` and `[OutputType()]` on all functions
- Use `[ValidateNotNullOrEmpty()]`, `[ValidatePattern()]`, `[ValidateRange()]`, `[ValidateSet()]` on parameters at system boundaries
- Access configuration via `$Script:Config.Section.Setting` — never use raw `$script:ParameterName` after initialization

## Architecture

### Configuration (`$Script:Config`)

All settings live in `Initialize-Configuration`. Structure:

```
$Script:Config.Runtime     — LogPath, WhatIfPreference, retries, timeouts
$Script:Config.BambooHR    — ApiKey, CompanyName, RootUri, ReportsUri
$Script:Config.Azure       — TenantId, UsageLocation, LicenseId, CompanyName
$Script:Config.Email       — AdminEmailAddress, NotificationEmailAddress, signatures
$Script:Config.Features    — EnableMobilePhoneSync, FullSync, MailboxDelegationParams, etc.
```

To add a new setting: add `param()` entry, add to `Initialize-Configuration`, add `BHR_CustomizationsJson` override support (with `PSBoundParameters` guard), document in header comments.

### Variable Naming

| Prefix      | Source                   |
| ----------- | ------------------------ |
| `$bhr*`     | BambooHR employee data   |
| `$entraId*` | Entra ID user attributes |
| No prefix   | Local/function-scoped    |
| `$Script:*` | Script-wide state        |

### Key Functions

| Function                             | Purpose                                                               |
| ------------------------------------ | --------------------------------------------------------------------- |
| `Write-PSLog`                        | All logging (Debug/Information/Warning/Error/Test)                    |
| `Invoke-WithRetry`                   | Retry with exponential backoff for API calls                          |
| `Get-CachedUser`                     | Cached Graph user lookups (use instead of `Get-MgUser` when possible) |
| `Test-ShouldSyncExistingUser`        | Gate for attribute sync pipeline                                      |
| `Test-IsOffboardingComplete`         | Check offboarding completion markers                                  |
| `Add-SignificantChange`              | Track changes for Teams summary card                                  |
| `Invoke-UserOffboarding`             | 17-step idempotent offboarding (shared by primary + retry paths)      |
| `Set-TerminatedUserProfileFields`    | Clear profile fields and address for terminated users                 |
| `Connect-ExchangeOnlineIfNeeded`     | Lazy Exchange Online connection with RPS session check                |
| `Invoke-SharedMailboxDelegationSync` | Group-based shared mailbox delegation reconciliation                  |

### Processing Pipeline

```
BambooHR API -> filter by tenant domain/inactive -> sort by LastName ->
  ForEach employee:
    lookup by UPN -> lookup by EmployeeID -> fallback merge ->
    CREATE / UPDATE / DISABLE (offboard) / RE-ENABLE / SKIP ->
    notifications
  Incomplete offboarding retry: Test-IsOffboardingComplete -> Invoke-UserOffboarding
```

## Build and Test

```powershell
# Run Pester tests
Invoke-Pester -Path ./tests/ -Output Detailed

# Preview changes without applying (WhatIf mode)
.\Start-BambooHRUserProvisioning.ps1 -WhatIf

# Run PSScriptAnalyzer
Invoke-ScriptAnalyzer -Path ./Start-BambooHRUserProvisioning.ps1 -Severity Warning -ReportSummary
```

### PSScriptAnalyzer

Suppress only with `[SuppressMessageAttribute()]` and a justification. Known suppression: `PSAvoidLongLines` (Graph API call complexity).

## Security Requirements

- **Never hardcode credentials** — API keys, passwords, and secrets come from Azure Automation encrypted variables or Managed Identity
- **Passwords**: Generated via `Get-NewPassword` with high entropy; never logged or stored in plaintext; send only via `Send-MgUserMail` to the manager
- **Validate all inputs** at system boundaries: email patterns, GUID formats, phone numbers, date formats
- **Sanitize BambooHR data** before use in Graph API calls — names go through `ConvertTo-StandardName`, phones through `ConvertTo-PhoneNumber`
- **API keys** are transmitted as Basic auth headers (Base64-encoded) over HTTPS only
- **Managed Identity** — no certificates, client secrets, or stored credentials
- **PII handling** — employee data is processed in-memory only; logs redact sensitive fields; temporary profile photos are cleaned up after upload; `$bhrHomeEmail` must never be sent to external recipients
- **OWASP considerations**: no SQL/command injection vectors (Graph SDK parameterized calls); no user-controlled file paths; secrets never in output streams

## Error Handling Conventions

- Wrap API calls in `Invoke-WithRetry` for transient failure resilience
- Track errors in `$errorSummary` hashtable (TotalErrors, ErrorsByType, ErrorsByUser, CriticalErrors)
- Individual user failures must not stop processing of remaining users
- Use `try/catch` around each discrete operation; log with `Write-PSLog -Severity Error`
- Clear `$error` after handling to prevent false positives in downstream checks
- Critical failures (BambooHR API down, Graph connection failed) → send alert email/Teams card and `exit 1`

### Soft vs Hard Failures

Certain errors are expected due to Entra ID's management model and should be treated as **soft failures** (log a warning, continue processing, still write completion markers):

- **Dynamic group membership removal** — Entra ID manages these automatically; removal attempts return errors that are safe to ignore
- **Group-based license removal** — Licenses assigned via group membership cannot be directly removed; log and continue
- **Already-disabled or already-deleted resources** — Idempotent operations that find the target state already achieved

All other API errors during offboarding or provisioning are **hard failures** — log as errors, skip completion markers, and let the retry path handle them on the next run.

## Conventions

- **Offboarding completion marker**: `CompanyName` field stamped with `MM/DD/YY (OffboardingComplete: yyyy-MM-ddTHH:mm:ssZ)` — set last as the transaction commit signal
- **EmployeeId = 'LVR'** marks fully offboarded accounts
- **ExtensionAttribute1** stores BambooHR `lastChanged` timestamp for change detection
- **Delta sync** (`ModifiedWithinDays`) is the default; `FullSync` processes all employees — schedule both
- **Shared mailbox delegation** runs at new-user onboarding and at script completion if `ForceSharedMailboxPermissions` or user changes detected
- **`ShouldProcess`** gates: all state-changing operations must be wrapped in `$PSCmdlet.ShouldProcess()` for `-WhatIf` support
- **Connection deduplication**: check `$Script:MgGraphConnected` / `$Script:ExchangeConnected` / `$Script:AzureConnected` before reconnecting
- **Date handling**: use invariant culture parsing (`[System.Globalization.CultureInfo]::InvariantCulture`) for all date conversions

## Testing Conventions

Tests use Pester 5+ and live in `tests/`. Run with `Invoke-Pester -Path ./tests/ -Output Detailed`.

### Test Architecture

- **AST-based function extraction** — Functions are extracted from the main script using `Get-FunctionDefinitionsFromFile` (parses AST, evaluates function text) rather than dot-sourcing the full script. This isolates unit tests from the script's connection and initialization logic.
- **`Initialize-TestScriptState`** — Sets up all `$script:` variables with safe test values (fake API keys, `contoso` company name, etc.) before running configuration tests.

### Test Patterns

- **Structure**: `Describe` per feature area → `Context` per function or scenario → `It` per assertion
- **Naming**: `It` blocks should read as plain English behavior descriptions (e.g., `'treats dynamic group removal as a soft failure and still writes completion markers'`)
- **Mock external services** — Always mock Graph, Exchange, and BambooHR API calls. Never make real API calls in tests.
- **Test both happy path and edge cases** — Include empty/null inputs, boundary values, and error conditions
- **Soft vs hard failure coverage** — Offboarding tests must verify that soft failures (dynamic groups, group-based licenses) still write completion markers, while hard failures skip them
- **Static validation** — Include at least one test that validates script syntax and coding conventions (e.g., no invalid variable scope references)
- **No real credentials** — Use obvious fakes like `'test-key'`, `'contoso'`, `'admin@contoso.com'` in test fixtures
