# Development

## Repository layout

| Path | Purpose |
| --- | --- |
| `Start-BambooHRUserProvisioning.ps1` | Scheduled reconciliation runbook |
| `Start-BambooHrWebhookSync.ps1` | Webhook-triggered targeted sync runbook |
| `Update-AzureAutomationRuntimeEnvironmentPSModules.ps1` | Runtime package maintenance runbook |
| `Add-ManagedIdentityPermissions.ps1` | Interactive admin permission bootstrap |
| `tests\` | Pester coverage and static validation |
| `infra\` | Azure Automation deployment assets |
| `.github\workflows\` | PR validation and guardrails |

## Architecture summary

The solution has two operational entry points and one maintenance runbook:

1. **Scheduled reconciliation** reads BambooHR broadly and applies lifecycle changes to Entra.
2. **Webhook sync** treats BambooHR webhook payloads as change signals, then re-reads BambooHR for authoritative data before applying targeted lifecycle updates.
3. **Runtime maintenance** keeps the PowerShell 7.4 runtime environment current.

## Configuration model

The runbooks normalize inputs in `Initialize-Configuration` and then rely on `$Script:Config`.

### Main configuration buckets

| Bucket | Examples |
| --- | --- |
| `$Script:Config.Runtime` | retry counts, timeouts, log path, `WhatIfPreference` |
| `$Script:Config.BambooHR` | API key, company subdomain, BambooHR endpoints |
| `$Script:Config.Azure` | tenant ID, usage location, license ID |
| `$Script:Config.Email` | admin/notification/help desk addresses and signatures |
| `$Script:Config.Features` | sync toggles, mailbox delegation settings, webhook-related behavior |

### Adding a new setting safely

1. Add the parameter to the runbook entry point.
2. Extend `Initialize-Configuration`.
3. Decide whether the setting belongs in a direct Automation variable or `BHR_CustomizationsJson`.
4. Use `$Script:Config` after initialization instead of raw parameter variables.
5. Add or update tests before merging.

## Processing pipeline

```text
BambooHR API -> domain and status filtering -> lookup in Entra -> create/update/disable/re-enable/skip ->
notifications -> summary reporting -> offboarding retry checks
```

The webhook runbook shares the same lifecycle logic but scopes the run to the employee IDs found in the webhook payload.

## Key functions worth knowing

| Function | Why it matters |
| --- | --- |
| `Write-PSLog` | Primary logging surface |
| `Invoke-WithRetry` | Retry wrapper for transient API failures |
| `Get-CachedUser` | Reduces repeated Graph lookups |
| `Test-ShouldSyncExistingUser` | Gates attribute updates |
| `Test-IsOffboardingComplete` | Detects incomplete offboarding |
| `Invoke-UserOffboarding` | Central idempotent offboarding flow |
| `Add-SignificantChange` | Feeds the Teams change summary |
| `Connect-ExchangeOnlineIfNeeded` | Avoids unnecessary Exchange reconnects |

## Coding standards

This project is PowerShell-first and tuned for Azure Automation:

- use approved verbs
- keep `param()` blocks with required parameters first, then optional ones
- prefer `[bool]` over `[switch]` in runbook entry points
- validate external inputs with `Validate*` attributes
- use `ShouldProcess` for state-changing work
- do not add interactive prompts to the runbooks
- prefer single-quoted strings unless interpolation is required

## Testing workflow

Run these before shipping a change:

```powershell
Invoke-Pester -Path .\tests\ -Output Detailed
Invoke-ScriptAnalyzer -Path .\Start-BambooHRUserProvisioning.ps1 -Severity Warning -ReportSummary
Invoke-ScriptAnalyzer -Path .\Start-BambooHrWebhookSync.ps1 -Severity Warning -ReportSummary
Invoke-ScriptAnalyzer -Path .\Deploy-AzureAutomation.ps1 -Severity Warning -ReportSummary
```

### Test design notes

- tests extract functions from the main scripts via AST rather than dot-sourcing the whole runbook
- `Initialize-TestScriptState` seeds safe defaults like `contoso`
- Graph, Exchange, and BambooHR calls must be mocked
- offboarding tests should distinguish soft failures from hard failures

## Change workflow

1. Identify whether the change affects scheduled sync, webhook sync, or both.
2. Update the relevant runbook and shared helper behavior.
3. Add or update Pester coverage.
4. Run ScriptAnalyzer and Pester.
5. If deployment behavior changed, compile `infra\main.bicep` and exercise `Deploy-AzureAutomation.ps1 -WhatIf`.

## Copilot guidance

The repo-specific Copilot instructions live in `.github\copilot-instructions.md`. Keep that file aligned with this document and the current documentation map whenever the structure or architecture changes.
