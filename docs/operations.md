# Operations

## Runbook responsibilities

| Runbook | Normal trigger | Purpose |
| --- | --- | --- |
| `Start-BambooHRUserProvisioning.ps1` | Hourly schedule + weekly full sync | Reconciliation safety net and primary scheduled lifecycle engine |
| `Start-BambooHrWebhookSync.ps1` | Azure Automation webhook | Fast, employee-targeted reaction path for BambooHR changes |
| `Update-AzureAutomationRuntimeEnvironmentPSModules.ps1` | Weekly schedule | Keeps the PowerShell 7.4 runtime environment packages current |

## Recommended operating cadence

### Hourly delta sync

Use the scheduled runbook for routine reconciliation. This catches:

- BambooHR changes that never triggered a webhook
- webhook delivery failures
- stale manager, title, department, and account state mismatches

### Weekly full sync

Run a weekly `FullSync = $true` schedule as the backstop for anything delta logic missed.

### Weekly runtime package maintenance

Keep the runtime package update runbook on a weekly schedule and review package drift in a change window when possible.

## Day-to-day checks

Review these items regularly:

1. Recent Automation jobs and failure counts.
2. Teams or email summaries for unusual spikes in creates, disables, or errors.
3. Runtime package update results.
4. Entra managed identity role assignments after permission changes.
5. Webhook usage if public access is enabled.

## Logs and signals

Use these signals first when diagnosing a problem:

- Azure Automation job output and job streams
- Teams adaptive card summaries
- notification emails sent by the runbooks
- `errorSummary` details in the job log

The Teams summary now reports tracked significant changes instead of generic log volume, so an empty change card is no longer expected on routine no-op runs.

## Safe change windows

Prefer to make operational changes in this order:

1. Run `Deploy-AzureAutomation.ps1 -WhatIf` if infrastructure or schedule behavior is changing.
2. Run the scheduled sync once with `-WhatIf`.
3. Trigger a targeted webhook test if webhook behavior changed.
4. Re-enable normal schedules after validation.

## Rotating secrets

### BambooHR API key

1. Create a replacement key in BambooHR.
2. Update the encrypted `BambooHrApiKey` Automation variable.
3. Run the scheduled sync once to confirm BambooHR connectivity.

### Teams webhook URI

1. Create the replacement Teams workflow/webhook endpoint.
2. Update the encrypted `TeamsCardUri` Automation variable.
3. Trigger a harmless `-WhatIf` run and verify the notification arrives.

### Azure Automation webhook URL

Treat the URL as a bearer secret:

1. Disable or remove the old webhook if it is exposed.
2. Create a new webhook in Azure Automation.
3. Update BambooHR with the new URL.
4. Rotate `BambooHrWebhookPrivateKey` at the same time if you use HMAC validation.

## When to disable public network access

Disable `publicNetworkAccess` on the Automation account if:

- you only use scheduled sync
- BambooHR webhook delivery is not required

Leave it enabled only when BambooHR must reach the Automation webhook endpoint.

## Incident response outline

### Webhook leak

1. Disable or delete the webhook immediately.
2. Rotate the webhook signing secret.
3. Create a new webhook and update BambooHR.
4. Review recent jobs for unexpected employee targets.

### Teams summary stops arriving

1. Verify `TeamsCardUri` still points to a valid Teams workflow/webhook.
2. Confirm the run completed and actually recorded significant changes.
3. Check [docs/troubleshooting.md](troubleshooting.md#teams-summary-or-change-card-missing).

### Permission regression

1. Re-run `Add-ManagedIdentityPermissions.ps1`.
2. Review the Automation account managed identity app role assignments.
3. Validate Exchange `Exchange.ManageAsApp` access separately from Graph permissions.
