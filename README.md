# BambooHR to Entra ID User Provisioning

Synchronizes BambooHR employees into Microsoft Entra ID from Azure Automation without requiring SCIM or extra BambooHR-side middleware.

This repository includes:

- `Start-BambooHRUserProvisioning.ps1` for scheduled reconciliation and catch-up syncs
- `Start-BambooHrWebhookSync.ps1` for BambooHR webhook-triggered targeted syncs
- `Deploy-AzureAutomation.ps1` and `infra\main.bicep` for repeatable Azure Automation deployment
- Pester, PSScriptAnalyzer, Bicep, and secret/PII guardrails in pull requests

## Quick start

1. Review the security model in [docs/security.md](docs/security.md).
2. Deploy the Azure Automation account with [docs/deployment.md](docs/deployment.md).
3. Grant managed identity Graph and Exchange permissions with [Add-ManagedIdentityPermissions.ps1](Add-ManagedIdentityPermissions.ps1).
4. Run an initial `-WhatIf` validation and confirm schedules, notifications, and webhook behavior.
5. Use [docs/operations.md](docs/operations.md) for daily operations and [docs/troubleshooting.md](docs/troubleshooting.md) when something fails.

## Documentation map

| Topic | Purpose |
| --- | --- |
| [docs/deployment.md](docs/deployment.md) | Prerequisites, Azure deployment flow, configuration inputs, webhook setup, and post-deploy steps |
| [docs/operations.md](docs/operations.md) | Schedules, monitoring, runtime package maintenance, and incident handling |
| [docs/development.md](docs/development.md) | Architecture, code layout, testing, and safe change workflow |
| [docs/troubleshooting.md](docs/troubleshooting.md) | Common failure modes and targeted diagnostics |
| [docs/security.md](docs/security.md) | Threat model, secret handling, PII boundaries, RBAC, and hardening checklist |

## Deployment model

The recommended deployment path is:

1. `Deploy-AzureAutomation.ps1` provisions the Automation account, PowerShell 7.4 runtime environment, recommended runtime packages, Automation variables, and schedules.
2. The same helper imports the runbooks, links them to the runtime environment, publishes them, and can create the BambooHR webhook.
3. `Add-ManagedIdentityPermissions.ps1` performs the privileged Graph and Exchange consent step that must remain an explicit admin action.

Secrets such as the BambooHR API key, Teams webhook URI, and BambooHR webhook signing key are never stored in the template or sample parameter file. They are written as encrypted Automation variables during deployment.

## Security model

Treat the Automation account as a privileged identity automation asset:

- use a **system-assigned** managed identity
- keep Azure RBAC scoped to a small admin/operator set
- disable public network access unless webhook delivery is required
- treat both the Azure Automation webhook URL and `TeamsCardUri` as bearer secrets
- keep employee data in-memory only and out of committed files

See [docs/security.md](docs/security.md) for the full threat model and hardening checklist.

## Noteworthy changes since December 2025

- **January 2026**: expanded offboarding to remove authentication methods, transfer owned groups, remove mailbox permissions, and improve mailbox-related error reporting.
- **March 2026**: made delta sync the default pattern, replaced runbook-incompatible `[switch]` feature flags with `[bool]`, and significantly expanded the Pester suite.
- **April 2026**: added the dedicated BambooHR webhook runbook, tightened Teams summary behavior to report only tracked significant changes, improved webhook handling for unsynced-only BambooHR field changes, and added PR guardrails for tests, linting, Bicep, and sensitive-data scanning.
- **Current deployment refresh**: added a supported Azure Automation Bicep template, a deployment helper, and a reorganized documentation set for deployment, operations, development, troubleshooting, and security.

## Runbooks in this repo

| Runbook | Purpose |
| --- | --- |
| `Start-BambooHRUserProvisioning.ps1` | Scheduled hourly delta sync and weekly full reconciliation |
| `Start-BambooHrWebhookSync.ps1` | Fast reaction path for BambooHR webhooks |
| `Update-AzureAutomationRuntimeEnvironmentPSModules.ps1` | Runtime package maintenance for the Automation PowerShell 7.4 environment |

## Validation commands

```powershell
Invoke-Pester -Path .\tests\ -Output Detailed
Invoke-ScriptAnalyzer -Path .\Start-BambooHRUserProvisioning.ps1 -Severity Warning -ReportSummary
Invoke-ScriptAnalyzer -Path .\Start-BambooHrWebhookSync.ps1 -Severity Warning -ReportSummary
```

For deployment-specific validation, compile `infra\main.bicep` and run `Deploy-AzureAutomation.ps1 -WhatIf` before making live changes.

## Acknowledgment

This repository started as a fork/adaptation of earlier BambooHR-to-Azure AD provisioning work by PaulTony. The current version focuses on Azure Automation, managed identity authentication, webhook support, and safer operational guardrails.
