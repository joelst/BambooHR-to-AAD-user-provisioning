# BambooHR-to-Entra-ID Azure Preparation Plan

## Status

Ready for Validation

## Objective

Prepare this repository for secure, easy-to-use Azure Automation deployment and improve the operator/developer experience by:

1. Adding a reusable Azure deployment template and deployment helper flow.
2. Reorganizing documentation into clear deployment, operations, development, troubleshooting, and security guidance.
3. Tightening repo guardrails and fixing a few adjacent usability/security gaps.

## Current State

- Bicep infrastructure and the `Deploy-AzureAutomation.ps1` helper are implemented under `infra\` and the repo root.
- Documentation has been reorganized into dedicated deployment, operations, development, troubleshooting, and security guides.
- PR guardrails now validate Bicep and block committed webhook or signature-bearing URLs.
- The UPN-change notification path has been corrected in both runbooks and covered by regression checks.

## Phase 1 Summary

### Workspace mode

Modify existing project

### Deployment recipe

Bicep for core Azure resources, plus PowerShell helper scripts for the parts that are operationally safer or more usable outside ARM:

- create/update secret Automation variables
- import and publish local runbook files
- create the Azure Automation webhook and reveal the one-time URL

### Why this recipe

- This is an existing Azure Automation runbook project, not a web app or azd-style app host.
- Azure Automation resource deployment fits ARM/Bicep well.
- One-time webhook URL handling and local runbook import are better handled by PowerShell than by static template outputs.
- Keeping secrets out of repo files and out of sample parameter files is a hard requirement.

## Architecture Decisions

### Core Azure resources

- Resource-group scoped Bicep under `infra\`
- Azure Automation Account
  - system-assigned managed identity
  - `disableLocalAuth = true`
  - configurable `publicNetworkAccess` so scheduled-only installs can disable public access while webhook-driven installs can leave it enabled
  - Basic SKU
- PowerShell 7.4 Runtime Environment
- Required runtime packages for the runbooks
- Non-secret Automation variables
- Scheduled reconciliation schedules and job schedules

### Runbook deployment model

- Deploy core infrastructure declaratively.
- Import local runbook source from the checked-out repo during the deployment helper step so users can deploy the exact code they have locally.
- Keep webhook creation outside Bicep so the URL is only revealed at deployment time and is not stored as a template output.

### Secret handling

Secrets will not be committed to repo files or sample parameter files.

Secret values to handle through the deployment helper:

- `BambooHrApiKey`
- `TeamsCardUri`
- `BambooHrWebhookPrivateKey`

Non-secret values can be stored in example parameters or passed through the helper:

- `BHRCompanyName`
- `CompanyName`
- `TenantId`
- `UsageLocation`
- `LicenseId`
- `AdminEmailAddress`
- `NotificationEmailAddress`
- `HelpDeskEmailAddress`
- `BHR_CustomizationsJson`

### Permission assignment boundary

Managed identity Graph and Exchange permission grants remain a post-deployment admin step, documented and wired around `Add-ManagedIdentityPermissions.ps1`, because that step is outside ARM resource deployment and requires privileged operator consent.

## Microsoft guidance captured during research

- Azure Automation PowerShell 7.4 must use the Runtime Environment experience.
- PowerShell 7.4 runbooks should be linked to a named runtime environment.
- Azure CLI 2.64.0 and Az 12.3.0 are supported default packages in PowerShell 7.4 runtime environments.
- Exchange Online PowerShell 3.x may require explicit `PowerShellGet` and `PackageManagement` packages in the runtime environment.
- PowerShell 7.4 source-control integration in Azure Automation is not supported, so repo-driven deployment should not rely on native Automation source-control sync.
- Azure Automation webhook URLs are bearer secrets and must be handled as one-time secrets.

## Planned Repository Changes

### 1. Azure deployment assets

Add:

- `infra\main.bicep`
- `infra\main.bicepparam` or example parameter file without secrets
- `infra\modules\...` if modularization improves readability
- `Deploy-AzureAutomation.ps1`
- optional webhook helper logic if kept separate from the main deploy helper

The deployment assets will:

- deploy the Automation Account and runtime environment
- provision required non-secret Automation variables
- provision schedules
- allow customization of names and operational defaults
- avoid embedding secrets

### 2. Documentation restructure

Rework docs into:

- `README.md` as concise onboarding and docs map
- `docs\deployment.md`
- `docs\operations.md`
- `docs\development.md`
- `docs\troubleshooting.md`
- `docs\security.md` or equivalent dedicated threat-model guidance

Documentation will cover:

- prerequisites and first-time setup
- deployment flow and required permissions
- threat model and security boundaries
- runtime/package maintenance
- operational monitoring, failure handling, and rollback guidance
- developer architecture and code-change workflow
- Copilot/developer guidance alignment with `.github\copilot-instructions.md`

### 3. Additional hardening and usability improvements

Plan to include:

- extend PR scanning to block committed webhook/bearer-secret URLs
- add CI validation for new Bicep artifacts
- fix the known UPN-change notification HTML/config bug in `Start-BambooHRUserProvisioning.ps1`
- remove or update outdated documentation sections and TODO placeholders

## Expected Files To Modify

- `README.md`
- `DEVELOPER_GUIDE.md`
- `.github\copilot-instructions.md`
- `.github\workflows\pr-checks.yml`
- `.github\scripts\Invoke-PiiScan.ps1`
- `Start-BambooHRUserProvisioning.ps1`
- `tests\Start-BambooHRUserProvisioning.Tests.ps1`

## Expected Files To Add

- `infra\...`
- `Deploy-AzureAutomation.ps1`
- `docs\...`

## Risks and Handling

- **Webhook secret exposure**: keep webhook creation out of Bicep outputs; reveal only once during helper execution.
- **Module/runtime drift**: pin and document runtime package versions; align with `Update-AzureAutomationRuntimeEnvironmentPSModules.ps1`.
- **Permission confusion**: clearly separate Azure deployment from Entra/Exchange permission grants.
- **Documentation sprawl**: move detailed content out of `README.md` and replace it with a clear navigation structure.

## Next Step

Run validation against the prepared artifacts before any live Azure deployment.
