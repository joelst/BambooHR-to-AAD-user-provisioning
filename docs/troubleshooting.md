# Troubleshooting

## Start with the first failing signal

When a run fails, work in this order:

1. Azure Automation job status and stream output
2. first error in the runbook log
3. related Teams/email summary behavior
4. managed identity permissions
5. BambooHR or webhook-side evidence

## Teams summary or change card missing

### Symptoms

- no Teams notification after a run
- only completion/no-change notifications arrive
- webhook URI was updated but notifications still do not appear

### Checks

1. Confirm `TeamsCardUri` is still a valid Teams workflow/webhook endpoint.
2. Confirm the run actually recorded significant changes.
3. Check whether the run failed before the Teams notification block.

### Notes

- The change card is intentionally gated on tracked significant changes now; ordinary log chatter is no longer enough to trigger it.
- If the Teams endpoint was rotated or retired, the encrypted Automation variable must be updated.

## Webhook job never starts

### Checks

1. Confirm the Automation account still allows public network access.
2. Confirm the Azure Automation webhook still exists and is enabled.
3. Confirm BambooHR is posting to the current webhook URL.
4. If a signing secret is configured, confirm `BambooHrWebhookPrivateKey` matches BambooHR.

## Webhook job starts but does not update the user you expected

### Checks

1. Review the webhook payload and extracted employee IDs.
2. Confirm the employee's BambooHR work email matches the expected tenant domain.
3. Review the changed field analysis in the webhook runbook output.

### Notes

- Unsynced-only BambooHR field changes still update the stored `lastChanged` marker and can add a note explaining that BambooHR changed a field the runbook does not write to Entra.
- The webhook runbook re-reads BambooHR after receiving the event; the webhook payload is treated as a signal, not as the source of truth.

## Runbook fails with Graph or Exchange errors

### Checks

1. Re-run `Add-ManagedIdentityPermissions.ps1`.
2. Confirm the Automation account managed identity still exists.
3. Validate Graph app role assignments in Entra.
4. Validate Exchange `Exchange.ManageAsApp` consent separately.

## Runbook fails immediately after import or publish

### Checks

1. Confirm the runbook is linked to the expected PowerShell 7.4 runtime environment.
2. Review runtime packages in the Automation account.
3. Run `Update-AzureAutomationRuntimeEnvironmentPSModules.ps1` if package drift is suspected.

## Schedule exists but no jobs are running

### Checks

1. Confirm the schedule exists in the Automation account.
2. Confirm the schedule is registered to the runbook.
3. Confirm the start time is in the future when initially created.
4. Confirm the Automation account region and subscription are the ones you intended to deploy.

## Pull request validation fails

### Pester or ScriptAnalyzer

- run the commands locally from [docs/development.md](development.md)
- fix the first failure before chasing downstream noise

### PII or secret scan

- do not commit real webhook URLs, Teams workflow URLs, API keys, tenant-specific contact details, or other environment secrets
- replace real identifiers with safe placeholders such as `contoso`

### Bicep validation

- compile `infra\main.bicep`
- check schedule start times, array values, and child-resource property names first
