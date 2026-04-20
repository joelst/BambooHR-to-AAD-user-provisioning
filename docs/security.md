# Security and Threat Model

## Security posture summary

This solution automates high-impact identity lifecycle changes. Treat it as privileged identity infrastructure, not as a convenience script.

## Trust boundaries

| Asset | Why it matters | Required control |
| --- | --- | --- |
| Azure Automation managed identity | Can create, disable, license, and modify users, groups, devices, mailboxes, and app assignments | Restrict Azure RBAC and review Graph/Exchange grants regularly |
| `BambooHrApiKey` | Grants BambooHR employee data access | Store only as an encrypted Automation variable |
| Azure Automation webhook URL | Bearer secret that can trigger targeted sync runs | Do not commit, email broadly, or store in plaintext notes |
| `BambooHrWebhookPrivateKey` | Validates BambooHR webhook authenticity | Rotate with the webhook URL when exposed |
| `TeamsCardUri` | Bearer secret to post notifications into Teams | Treat as a secret, not as a benign URL |
| Employee PII | HR and identity data is sensitive even when not strictly secret | Keep data out of the repo and out of notification fan-out beyond authorized recipients |

## Threat model

### 1. Unauthorized runbook execution

**Risk:** anyone with overly broad Automation permissions can run a job under the managed identity.

**Controls:**

- use `Automation Contributor` only for trusted admins
- use `Automation Operator` only for approved operators
- avoid broad `Owner` or `Contributor` assignments at the resource-group level
- prefer PIM or other just-in-time elevation

### 2. Secret leakage

**Risk:** BambooHR keys, Teams webhook URLs, or Automation webhook URLs leak through commits, logs, chat, or docs.

**Controls:**

- secrets stay out of Bicep and sample parameter files
- deployment helper writes secrets into encrypted Automation variables
- PR validation includes sensitive-data scanning
- webhook URLs are revealed only once at creation time

### 3. Overexposed network surface

**Risk:** public webhook access is left enabled when the environment only needs schedules.

**Controls:**

- disable `publicNetworkAccess` unless BambooHR webhook delivery is required
- if webhook mode is enabled, treat the public endpoint as a secret-bearing interface

### 4. Unreviewed code changes

**Risk:** a runbook change introduces unsafe behavior or accidentally commits tenant-specific data.

**Controls:**

- Pester tests
- PSScriptAnalyzer
- Bicep compilation checks
- PR secret/PII scanning

### 5. Data over-disclosure in notifications

**Risk:** Teams/email summaries expose information to recipients who do not need it.

**Controls:**

- keep notification recipients tightly scoped
- avoid including unnecessary personal fields in summaries
- do not send home email or other off-platform contact data externally

## Hardening checklist

Use this checklist for every environment:

1. System-assigned managed identity only.
2. Minimal Azure RBAC on the Automation account.
3. Managed identity permissions limited to the Graph and Exchange roles this solution actually needs.
4. `disableLocalAuth = true` on the Automation account.
5. `publicNetworkAccess = false` unless webhook mode is required.
6. Encrypted Automation variables for all secrets.
7. Routine review of webhook URLs and notification endpoints.
8. Logging and diagnostic retention appropriate for audit needs.

## Operational security practices

### Rotate on suspicion, not just on schedule

If any webhook URL, Teams workflow URL, or BambooHR credential is exposed, rotate it immediately.

### Review app role assignments

Any time you change Graph or Exchange behavior, re-check the managed identity's grants to confirm you did not add broader access than intended.

### Protect the deployment helper

`Deploy-AzureAutomation.ps1` handles secrets locally. Run it from a trusted workstation, avoid transcript logging that captures secure values, and do not save console output that includes generated webhook material in shared locations.

## Repo hygiene

Do not commit:

- real webhook URLs
- API keys or bearer tokens
- tenant-specific employee or admin contact data
- local settings files containing secrets

The repository uses automated PR checks to reduce this risk, but prevention still starts with how changes are authored.
