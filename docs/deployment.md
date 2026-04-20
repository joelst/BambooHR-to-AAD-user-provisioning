# Deployment

## What gets deployed

The supported deployment path provisions an Azure Automation-based runtime for this repository:

- Azure Automation account with a **system-assigned** managed identity
- PowerShell 7.4 Runtime Environment
- recommended runtime packages for Graph and Exchange Online
- non-secret Automation variables
- hourly delta, weekly full sync, and weekly runtime package maintenance schedules
- imported and published runbooks from the local checkout
- optional Azure Automation webhook for BambooHR-triggered targeted sync

## Prerequisites

Before deploying, make sure you have:

1. An Azure subscription and permission to create resource groups and Automation resources.
2. An Entra tenant ID.
3. A BambooHR API key.
4. An operator account that can grant Graph and Exchange managed identity permissions after the Automation account exists.
5. Local PowerShell 7 with `Az.Accounts`, `Az.Resources`, and `Az.Automation` available.

## Recommended deployment flow

The easiest path is the helper script because it keeps secrets out of the template and handles runbook import plus webhook creation.

### 1. Collect secrets without pasting them into repo files

```powershell
$bambooHrApiKey = Read-Host 'BambooHR API key' -AsSecureString
$teamsCardUri = Read-Host 'Teams webhook URI (optional)' -AsSecureString
```

If you plan to use BambooHR webhooks, you can let the helper generate a signing secret automatically or supply one as a secure string.

### 2. Deploy Azure Automation

```powershell
.\Deploy-AzureAutomation.ps1 `
  -SubscriptionId '00000000-0000-0000-0000-000000000000' `
  -TenantId '00000000-0000-0000-0000-000000000000' `
  -Location 'eastus' `
  -ResourceGroupName 'rg-bhr-sync-prod' `
  -AutomationAccountName 'aa-bhr-sync-prod' `
  -CompanyName 'Contoso' `
  -BhrCompanyName 'contoso' `
  -BambooHrApiKey $bambooHrApiKey `
  -TeamsCardUri $teamsCardUri `
  -UsageLocation 'US' `
  -ModifiedWithinDays 2 `
  -CreateWebhook $true
```

The helper will:

1. Ensure the resource group exists.
2. Deploy `infra\main.bicep`.
3. Import the local runbooks and link them to the named PowerShell 7.4 runtime environment.
4. Publish the runbooks.
5. Create encrypted Automation variables for secrets.
6. Register the runbooks to the deployed schedules.
7. Optionally create an Azure Automation webhook and return the one-time URL plus the BambooHR signing secret.

### 3. Grant managed identity permissions

After the Automation account exists, run:

```powershell
.\Add-ManagedIdentityPermissions.ps1
```

That script is intentionally separate because app role assignment and Exchange app consent are privileged admin actions.

### 4. Configure BambooHR webhook delivery

If you enabled webhook creation:

1. Save the returned Azure Automation webhook URL immediately. Azure does not show it again.
2. Save the BambooHR webhook signing secret if the helper generated one for you.
3. In BambooHR, create or update the webhook to post to the returned URL.

Use the webhook signing secret in BambooHR if you want `Start-BambooHrWebhookSync.ps1` to validate the HMAC header.

## Customizing non-secret behavior

Most common non-secret settings can be passed directly to `Deploy-AzureAutomation.ps1` and are stored in `BHR_CustomizationsJson`.

Examples:

- `-UsageLocation`
- `-ModifiedWithinDays`
- `-DaysAhead`
- `-DaysToKeepAccountsAfterTermination`
- `-EnableMobilePhoneSync`
- `-ForceSharedMailboxPermissions`
- `-CurrentOnly`

For advanced settings, provide a JSON file and pass `-CustomizationsJsonFilePath`.

## Pure Bicep usage

If you want to deploy only the infrastructure with ARM/Bicep first:

1. Copy `infra\main.parameters.example.json`.
2. Fill in the non-secret values and future schedule start times.
3. Deploy `infra\main.bicep` with your preferred ARM/Bicep tool.
4. Use `Deploy-AzureAutomation.ps1` afterward to import runbooks, set secrets, and optionally create the webhook.

## Deployment choices that matter

### Public network access

- Set `PublicNetworkAccess` to `$false` if you do **not** need BambooHR webhooks.
- Leave it enabled if BambooHR must reach the Automation webhook URL from the public internet.

### Schedule defaults

If you do not supply explicit start times, the helper chooses:

- next top-of-hour UTC for delta sync
- next configured weekly day at 03:00 UTC for full sync
- next configured weekly day at 04:00 UTC for runtime package maintenance

### Secrets

These values are written as encrypted Automation variables instead of template parameters:

- `BambooHrApiKey`
- `TeamsCardUri`
- `BambooHrWebhookPrivateKey`

## Post-deployment checklist

1. Confirm the Automation account identity exists and record its principal ID.
2. Complete `Add-ManagedIdentityPermissions.ps1`.
3. Run `Start-BambooHRUserProvisioning.ps1` once with `-WhatIf`.
4. Confirm the hourly and weekly schedules are linked to the expected runbooks.
5. If webhook mode is enabled, test BambooHR delivery with a non-production employee event or use `Test-AutomationWebhook.ps1`.
6. Review [docs/operations.md](operations.md) and [docs/security.md](security.md) before turning production schedules loose.
