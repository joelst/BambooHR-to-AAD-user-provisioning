# Developer Guide: BambooHR to Entra ID Sync Script

## Overview

This guide helps PowerShell developers understand the BambooHR to Azure AD user provisioning script. The script automatically synchronizes employee data from BambooHR into Entra ID.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Architecture Overview](#architecture-overview)
3. [Script Flow](#script-flow)
4. [Key Concepts](#key-concepts)
5. [Configuration](#configuration)
6. [How to Read the Code](#how-to-read-the-code)
7. [Common Tasks](#common-tasks)
8. [Troubleshooting](#troubleshooting)
9. [Best Practices](#best-practices)

---

## Prerequisites

### Required Knowledge
- PowerShell basics (variables, functions, loops)
- Understanding of Azure AD/Entra ID concepts
- REST API fundamentals
- Basic understanding of hashtables and arrays

### Required Software
1. **PowerShell**: Version 7+ recommended
2. **Required Modules**:
   ```powershell
   Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser
   Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser
   Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
   Install-Module -Name Microsoft.Graph.Identity.SignIns -Scope CurrentUser
   Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
   Install-Module -Name PSTeams -Scope CurrentUser
   ```

### Azure Requirements
- Azure Automation Account with Managed Identity
- BambooHR API key and access
- Microsoft Graph API permissions:
  - `User.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `Mail.Send`

---

## Architecture Overview

```
┌─────────────┐         ┌──────────────┐         ┌─────────────┐
│  BambooHR   │────────▶│  PowerShell  │────────▶│  Azure AD   │
│     API     │         │    Script    │         │  (Entra)    │
└─────────────┘         └──────────────┘         └─────────────┘
      │                        │                         │
      │                        │                         │
      │                        ▼                         │
      │                 ┌──────────────┐                │
      │                 │ Performance  │                │
      │                 │    Cache     │                │
      │                 └──────────────┘                │
      │                        │                         │
      │                        ▼                         │
      │                 ┌──────────────┐                │
      │                 │ Error Summary│                │
      │                 │   Tracking   │                │
      │                 └──────────────┘                │
      │                        │                         │
      │                        ▼                         │
      └────────────────▶ Email & Teams ◀────────────────┘
                         Notifications
```

### Components

1. **Configuration Management**: Centralized settings in `$Script:Config`
2. **Retry Logic**: Automatic retry for transient failures
3. **Performance Caching**: Reduces API calls by 30-50%
4. **Error Tracking**: Comprehensive error collection and reporting
5. **Logging**: Multi-destination logging (console, file, email)

---

## Script Flow

### High-Level Flow

```
1. Initialize Configuration
   ├─ Load parameters
   ├─ Retrieve Azure Automation variables
   ├─ Validate required settings
   └─ Set up logging

2. Connect to Services
   ├─ Microsoft Graph (Azure AD)
   ├─ Exchange Online
   └─ BambooHR API

3. Retrieve Employee Data
   ├─ Call BambooHR API
   ├─ Filter by company email domain
   └─ Sort by last name

4. Process Each Employee
   ├─ Extract BambooHR data
   ├─ Look up Azure AD account
   ├─ Determine action (Create/Update/Disable)
   ├─ Apply changes (with retry)
   ├─ Update manager relationship
   ├─ Send notifications
   └─ Track errors

5. Generate Reports
   ├─ Performance statistics
   ├─ Error summary
   ├─ Email notifications
   └─ Teams adaptive cards
```

### Detailed Employee Processing Flow

```
For Each Employee:
  │
  ├─ Extract Data from BambooHR
  │   ├─ Name, email, job title
  │   ├─ Department, location
  │   ├─ Manager, hire date
  │   └─ Employment status
  │
  ├─ Lookup in Azure AD
  │   ├─ By UPN (email)
  │   └─ By EmployeeId
  │
  ├─ Account Exists?
  │   │
  │   YES ──▶ ├─ Compare Attributes
  │   │       ├─ Update if changed
  │   │       ├─ Set manager
  │   │       ├─ Check status (enable/disable)
  │   │       └─ Send notification
  │   │
  │   NO ───▶ ├─ Create New Account
  │           ├─ Set all attributes
  │           ├─ Assign manager
  │           ├─ Assign license
  │           ├─ Send welcome email
  │           └─ Track in error summary
  │
  └─ Continue to Next Employee
```

---

## Key Concepts

### 1. Configuration Object (`$Script:Config`)

All settings are stored in a centralized hashtable:

```powershell
$Script:Config = @{
    Runtime = @{
        TestOnly            # Preview mode (no changes)
        MaxRetryAttempts    # Number of retry attempts
        LogFilePath         # Where to save logs
        CorrelationId       # Unique run identifier
    }
    BambooHR = @{
        CompanyName         # BambooHR subdomain
        ApiKey              # Authentication key
        ApiBaseUrl          # API endpoint
    }
    Azure = @{
        TenantId            # Azure AD tenant
        LicenseId           # M365 license SKU
        UsageLocation       # Default usage location
    }
    Email = @{
        AdminEmail          # Administrator address
        CompanyEmailDomain  # @company.com
    }
    Features = @{
        EnableMobilePhone   # Sync mobile numbers
        TeamsCardUri        # Teams webhook
    }
    Performance = @{
        MaxParallelUsers    # Parallel processing limit
        BatchSize           # Bulk operation size
    }
}
```

**Access Settings:**
```powershell
# Good: Use centralized config
$apiKey = $Script:Config.BambooHR.ApiKey

# Bad: Direct variable access (outdated pattern)
$apiKey = $BambooHrApiKey  # Don't do this!
```

### 2. Retry Logic (`Invoke-WithRetry`)

Automatically retries failed API calls:

```powershell
Invoke-WithRetry -Operation "Create user" -ScriptBlock {
    New-MgUser -UserPrincipalName $upn -DisplayName $name
}
```

**How it Works:**
- Attempt 1: Immediate execution
- Attempt 2: Wait 1 second + random jitter (0-500ms)
- Attempt 3: Wait 2 seconds + jitter
- Attempt 4: Wait 4 seconds + jitter
- Continues with exponential backoff until max attempts

**Retryable Errors:**
- Network issues (WebException)
- HTTP errors (HttpRequestException)
- Timeouts (TimeoutException)
- Rate limiting (HTTP 429)
- Service unavailable (HTTP 503, 504)

### 3. Performance Caching

Reduces redundant API calls:

```powershell
# First call: API call (~500ms)
$manager1 = Get-CachedUser -UserId "manager@company.com" -Cache $performanceCache

# Second call: Cache hit (<1ms)
$manager2 = Get-CachedUser -UserId "manager@company.com" -Cache $performanceCache
```

**When to Use:**
- Manager lookups (same manager for many employees)
- Repeated user queries within same run
- Operations where slight data staleness is acceptable

**When NOT to Use:**
- Just modified the user (data may be stale)
- Need to verify current account state
- Use `-Force` parameter to bypass cache if needed

### 4. Error Tracking

Collects errors for reporting:

```powershell
# Track an error
$errorSummary.TotalErrors++
$errorType = "UserCreation"
if (-not $errorSummary.ErrorsByType.ContainsKey($errorType)) {
    $errorSummary.ErrorsByType[$errorType] = 0
}
$errorSummary.ErrorsByType[$errorType]++
$errorSummary.ErrorsByUser[$email] = "Failed to create user: $($Error[0].Exception.Message)"
```

**Error Categories:**
- `UserCreation`: Failed to create new accounts
- `AttributeUpdate`: Failed to update user properties
- `ManagerAssignment`: Failed to set manager relationships
- `MailboxPermissions`: Failed to set delegation

### 5. Logging (`Write-PSLog`)

Centralized logging function:

```powershell
Write-PSLog -Message "User created successfully" -Severity Information
Write-PSLog -Message "Warning: Manager not found" -Severity Warning
Write-PSLog -Message "Error: Failed to update" -Severity Error
```

**Severity Levels:**
- `Debug`: Detailed diagnostic info (verbose only)
- `Information`: Normal operations
- `Warning`: Non-critical issues
- `Error`: Failures requiring attention
- `Test`: TestOnly mode previews

**Where Logs Go:**
1. Console (color-coded by severity)
2. CSV file (for historical analysis)
3. In-memory array (for email/Teams notifications)

---

## Configuration

### Azure Automation Variables

Set these in your Azure Automation account:

| Variable Name | Type | Required | Description |
|--------------|------|----------|-------------|
| `BambooHrApiKey` | Encrypted | Yes | BambooHR API authentication key |
| `BHRCompanyName` | String | Yes | BambooHR company subdomain |
| `CompanyName` | String | Yes | Your company display name |
| `TenantId` | String | Yes | Azure AD tenant ID (GUID) |
| `AdminEmailAddress` | String | Yes | Administrator email for notifications |
| `TeamsCardUri` | String | No | Teams webhook URL for cards |
| `LicenseId` | String | No | M365 license SKU for new users |
| `BHRScript_MaxRetryAttempts` | Integer | No | Retry attempts (default: 3) |
| `BHR_CustomizationsJson` | String | No | JSON overrides for environment-specific customizations |

#### Customizations JSON (recommended)

Use `BHR_CustomizationsJson` to override settings without maintaining a separate script. The JSON is loaded during configuration and can override parameters like `DaysToKeepAccountsAfterTermination`, `MailboxDelegationParams`, `TeamsCardUri`, `AdminEmailAddress`, `NotificationEmailAddress`, `HelpDeskEmailAddress`, `UsageLocation`, `DaysAhead`, `EnableMobilePhoneSync`, `CurrentOnly`, `ForceSharedMailboxPermissions`, `DefaultProfilePicPath`, `EmailSignature`, and `WelcomeUserText`.

Example JSON matching _Start-BambooHRUserProvisioning.ps1:

```json
{
    "DaysToKeepAccountsAfterTermination": 14,
    "MailboxDelegationParams": [
        { "Group": "CG-SharedMailboxDelegatedAccessSchedules", "DelegateMailbox": "Scheduling" },
        { "Group": "CG-SharedMailboxDelegatedAccessSupport", "DelegateMailbox": "CustomerCare" },
        { "Group": "CG-SharedMailboxDelegatedAccessLeads", "DelegateMailbox": "Lead" },

    ]
}
```

### Command-Line Parameters

Override automation variables:

```powershell
# Test mode (preview changes without applying)
.\Start-BambooHRUserProvisioning.ps1 -TestOnly

# Enable mobile phone sync
.\Start-BambooHRUserProvisioning.ps1 -EnableMobilePhoneSync

# Current employees only (no pre-hires)
.\Start-BambooHRUserProvisioning.ps1 -CurrentOnly

# Custom retry settings
.\Start-BambooHRUserProvisioning.ps1 -MaxRetryAttempts 5 -RetryDelaySeconds 2
```

---

## How to Read the Code

### Script Structure

The script is organized into logical regions:

1. **Header Section** (Lines 1-140)
   - Requirements declaration
   - Setup guide documentation
   - Synopsis and examples
   - Parameter definitions

2. **Script-Level Variables** (Lines 260-280)
   - `$Script:CorrelationId`: Unique run identifier
   - `$Script:StartTime`: Script start timestamp
   - `$Script:logContent`: Log message array

3. **Configuration Management** (Lines 280-520)
   - `Initialize-Configuration`: Load and validate settings
   - Configuration initialization
   - Validation checks

4. **Logging Functions** (Lines 540-590)
   - `Write-PSLog`: Centralized logging
   - Multiple output destinations

5. **Retry and Error Handling** (Lines 590-710)
   - `Invoke-WithRetry`: Automatic retry logic
   - Exponential backoff implementation

6. **Performance Helper Functions** (Lines 710-900)
   - `Test-ParallelProcessingSupport`: PS version detection
   - `Initialize-PerformanceCache`: Cache setup
   - `Get-CachedUser`: Cached user lookups
   - `Get-PerformanceStatistics`: Metrics calculation
   - `Get-ErrorSummaryReport`: Error analysis

7. **Authentication** (Lines 900-1000)
   - `Connect-MgGraphFunctions`: Graph API connection
   - `Connect-ExchangeOnline`: Exchange connection
   - Managed Identity authentication

8. **BambooHR Integration** (Lines 1000-1600)
   - `Get-BambooHREmployees`: Employee data retrieval
   - API authentication
   - Data parsing

9. **Main Processing Loop** (Lines 1600-3000)
   - Employee iteration
   - User lookup
   - Create/Update/Disable logic
   - Manager assignment
   - Error tracking

10. **Completion and Reporting** (Lines 3000-3160)
    - Performance statistics
    - Error summary report
    - Email notifications
    - Teams cards

### Code Regions

The script uses `#region` markers for easy navigation:

```powershell
#region Configuration Management
# ... configuration code ...
#endregion Configuration Management

#region Retry and Error Handling
# ... retry logic ...
#endregion Retry and Error Handling
```

**In VS Code:** Use `Ctrl+M, Ctrl+O` to collapse all regions.

### Variable Naming Conventions

```powershell
# BambooHR data (source)
$bhrFirstName = "John"
$bhrJobTitle = "Engineer"
$bhrEmployeeNumber = "12345"

# Azure AD data (target)
$aadFirstName = "John"
$aadJobTitle = "Engineer"
$aadEmployeeNumber = "12345"

# Script-level (accessible everywhere)
$Script:Config = @{ ... }
$Script:CorrelationId = "..."

# Local function variables (no prefix)
$retryCount = 0
$cacheHit = $true
```

---

## Common Tasks

### Add a New Configuration Parameter

1. **Add to param() block** (top of script):
```powershell
[Parameter(HelpMessage = "My new parameter description.")]
[string]
$MyNewParameter
```

2. **Add to Initialize-Configuration**:
```powershell
function Initialize-Configuration {
    # ... existing code ...

    $config.MySection = @{
        MyNewParameter = if ($MyNewParameter) { $MyNewParameter }
                        else { Get-AutomationVariable -Name 'MyNewParameter' -ErrorAction SilentlyContinue }
    }

    # Validate if required
    if ([string]::IsNullOrWhiteSpace($config.MySection.MyNewParameter)) {
        $config.ValidationErrors += "MyNewParameter is required"
        $config.IsValid = $false
    }
}
```

3. **Use throughout script**:
```powershell
$value = $Script:Config.MySection.MyNewParameter
```

### Add a New BambooHR Field to Sync

1. **Update BambooHR API call** to include field:
```powershell
$fields = "firstName,lastName,jobTitle,myNewField"
```

2. **Extract in employee loop**:
```powershell
$bhrMyNewField = "$($_.myNewField)"
```

3. **Get from Azure AD**:
```powershell
$aadMyNewField = $aadUpnObjDetails.MyProperty
```

4. **Compare and update**:
```powershell
if ($bhrMyNewField -ne $aadMyNewField) {
    Write-PSLog "Updating MyNewField from '$aadMyNewField' to '$bhrMyNewField'" -Severity Information
    Invoke-WithRetry -Operation "Update MyNewField" -ScriptBlock {
        Update-MgUser -UserId $bhrWorkEmail -MyProperty $bhrMyNewField
    }
}
```

### Wrap an API Call with Retry

Before:
```powershell
Update-MgUser -UserId $email -JobTitle $title
```

After:
```powershell
Invoke-WithRetry -Operation "Update job title: $email" -ScriptBlock {
    Update-MgUser -UserId $email -JobTitle $title
}
```

### Track a New Error Type

1. **When error occurs**:
```powershell
try {
    # ... operation ...
}
catch {
    Write-PSLog "Error: $($_.Exception.Message)" -Severity Error

    # Track in error summary
    $errorSummary.TotalErrors++
    $errorType = "MyNewErrorType"
    if (-not $errorSummary.ErrorsByType.ContainsKey($errorType)) {
        $errorSummary.ErrorsByType[$errorType] = 0
    }
    $errorSummary.ErrorsByType[$errorType]++
    $errorSummary.ErrorsByUser[$email] = "Description: $($_.Exception.Message)"
}
```

2. **Error will be included** in email report automatically.

---

## Troubleshooting

### Common Issues

#### 1. "Configuration validation failed"

**Cause:** Missing required parameters.

**Solution:**
1. Check Azure Automation variables are set
2. Verify parameter names match exactly
3. Check `Initialize-Configuration` function for validation logic

```powershell
# Check what's missing
$Script:Config.ValidationErrors
```

#### 2. "Failed to connect to Microsoft Graph"

**Cause:** Managed Identity not configured or lacks permissions.

**Solution:**
1. Verify Managed Identity is enabled in Azure Automation account
2. Check Graph API permissions in Enterprise Applications
3. Review `Connect-MgGraphFunctions` for connection logic

#### 3. "User not found" but should exist

**Cause:** Lookup by UPN or EmployeeId failing.

**Solution:**
1. Check email domain matches company domain
2. Verify EmployeeId is synced correctly
3. Use `-Force` on cached lookups to get fresh data

```powershell
# Bypass cache to get fresh data
$user = Get-CachedUser -UserId $email -Cache $performanceCache -Force
```

#### 4. "Too Many Requests" (HTTP 429)

**Cause:** Rate limiting from API.

**Solution:**
- Invoke-WithRetry handles this automatically
- Increase retry delay: `-RetryDelaySeconds 2`
- Check error logs for frequency

#### 5. Cache hit rate very low

**Cause:** Not using cached functions for repeated lookups.

**Solution:**
Replace direct Get-MgUser calls with Get-CachedUser:

```powershell
# Bad: Direct API call
$manager = Get-MgUser -UserId $managerEmail

# Good: Use cache
$manager = Get-CachedUser -UserId $managerEmail -Cache $performanceCache
```

### Debugging Tips

1. **Enable TestOnly mode** to preview without changes:
```powershell
.\Start-BambooHRUserProvisioning.ps1 -TestOnly
```

2. **Check correlation ID** in logs:
```
[abc123de] [2025-10-02 14:30:15] User created successfully
```
All related log entries share same correlation ID.

3. **Review error summary** email for patterns.

4. **Check performance statistics**:
```
Performance Statistics:
  Duration: 125.3 seconds
  Users Processed: 250
  Throughput: 119.6 users/minute
  Cache Hit Rate: 76.3% (234/307 lookups)
```

5. **Inspect Azure AD directly**:
```powershell
# Verify user state
Get-MgUser -UserId $email | Format-List *

# Check manager assignment
Get-MgUserManager -UserId $email
```

---

## Best Practices

### 1. Always Use Centralized Config

✅ **Good:**
```powershell
$apiKey = $Script:Config.BambooHR.ApiKey
```

❌ **Bad:**
```powershell
$apiKey = $BambooHrApiKey  # Direct variable access
```

### 2. Wrap API Calls in Retry Logic

✅ **Good:**
```powershell
Invoke-WithRetry -Operation "Update user" -ScriptBlock {
    Update-MgUser -UserId $email -JobTitle $title
}
```

❌ **Bad:**
```powershell
Update-MgUser -UserId $email -JobTitle $title  # No retry
```

### 3. Use Write-PSLog for All Logging

✅ **Good:**
```powershell
Write-PSLog "User created: $email" -Severity Information
```

❌ **Bad:**
```powershell
Write-Host "User created: $email"  # Not captured for notifications
```

### 4. Use Cached Lookups When Possible

✅ **Good:**
```powershell
$manager = Get-CachedUser -UserId $managerEmail -Cache $performanceCache
```

❌ **Bad:**
```powershell
$manager = Get-MgUser -UserId $managerEmail  # Repeated API calls
```

### 5. Track Errors for Reporting

✅ **Good:**
```powershell
try {
    # operation
}
catch {
    Write-PSLog "Error: $_" -Severity Error
    $errorSummary.TotalErrors++
    # ... track error details ...
}
```

❌ **Bad:**
```powershell
try {
    # operation
}
catch {
    Write-Error $_  # Not tracked for summary
}
```

### 6. Test in TestOnly Mode First

✅ **Good:**
```powershell
# Always test first
.\Start-BambooHRUserProvisioning.ps1 -TestOnly

# Review logs carefully

# Run for real
.\Start-BambooHRUserProvisioning.ps1
```

❌ **Bad:**
```powershell
# Run directly in production
.\Start-BambooHRUserProvisioning.ps1
```

### 7. Use Appropriate Severity Levels

```powershell
# Debug: Detailed diagnostic info
Write-PSLog "Checking if user exists..." -Severity Debug

# Information: Normal operations
Write-PSLog "User created successfully" -Severity Information

# Warning: Non-critical issues
Write-PSLog "Manager not found, skipping assignment" -Severity Warning

# Error: Failures requiring attention
Write-PSLog "Failed to create user: $_" -Severity Error
```

### 8. Handle Null/Empty Values

✅ **Good:**
```powershell
if ([string]::IsNullOrWhiteSpace($bhrManagerEmail) -eq $false) {
    $manager = Get-CachedUser -UserId $bhrManagerEmail -Cache $performanceCache
}
```

❌ **Bad:**
```powershell
$manager = Get-CachedUser -UserId $bhrManagerEmail -Cache $performanceCache  # May fail if null
```

### 9. Use PSScriptAnalyzer Suppressions Appropriately

Only suppress warnings when justified:

```powershell
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '')]
param(
    [string]$ApiKey  # Justified: Retrieved from secure Azure Automation variable
)
```

### 10. Document Complex Logic

✅ **Good:**
```powershell
<#
  USER ACCOUNT EXISTS CHECK:
  If we found a user by UPN OR by EmployeeId, then the account exists.
  This section handles UPDATES to existing accounts.
#>
if (([string]::IsNullOrEmpty($aadUpnObjDetails) -eq $false) -or
    ([string]::IsNullOrEmpty($aadEidObjDetails) -eq $false)) {
    # ... update logic ...
}
```

❌ **Bad:**
```powershell
if (([string]::IsNullOrEmpty($aadUpnObjDetails) -eq $false) -or
    ([string]::IsNullOrEmpty($aadEidObjDetails) -eq $false)) {
    # ... update logic with no explanation ...
}
```

---

## Additional Resources

### Documentation Files

- `README.md`: Project overview and basic setup
- `CONFIGURATION_MIGRATION_SUMMARY.md`: Configuration centralization details
- `RETRY_LOGIC_IMPLEMENTATION_SUMMARY.md`: Retry mechanism documentation
- `PERFORMANCE_OPTIMIZATION_SUMMARY.md`: Caching and performance features
- `CACHING_ERROR_SUMMARY_IMPLEMENTATION.md`: Latest optimizations
- `JUNIOR_DEVELOPER_GUIDE.md`: This file!

### Microsoft Documentation

- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/powershell/microsoftgraph/)
- [Azure Automation](https://learn.microsoft.com/azure/automation/)
- [BambooHR API](https://documentation.bamboohr.com/docs)

### PowerShell Best Practices

- [PSScriptAnalyzer Rules](https://github.com/PowerShell/PSScriptAnalyzer)
- [PowerShell Best Practices](https://learn.microsoft.com/powershell/scripting/developer/cmdlet/cmdlet-development-guidelines)

---

## Getting Help

### Within the Script

1. **Read inline comments**: Key sections have comprehensive explanations
2. **Check region headers**: `#region` blocks explain major sections
3. **Review function help**: Each function has `.SYNOPSIS` and `.DESCRIPTION`

### External Resources

1. **Error Messages**: Search for specific error text in Microsoft Docs
2. **Graph API**: Check [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) for API details
3. **Community**: Stack Overflow, Reddit r/PowerShell

### When Asking for Help

Include:
1. **Correlation ID** from logs
2. **Specific error message**
3. **What you expected** vs. what happened
4. **Azure Automation job output** (if running in automation)
5. **PowerShell version**: `$PSVersionTable.PSVersion`

---

## Conclusion

This script implements enterprise-grade user provisioning with:

- ✅ Centralized configuration management
- ✅ Automatic retry logic for reliability
- ✅ Performance caching (30-50% improvement)
- ✅ Comprehensive error tracking and reporting
- ✅ Multiple logging destinations
- ✅ TestOnly mode for safe testing
- ✅ Detailed inline documentation

Take time to read through the code regions systematically. Start with configuration, then understand the flow through authentication, data retrieval, and user processing.

**Remember:** The script is designed to be read and maintained. All complex logic has explanatory comments. When in doubt, enable TestOnly mode and review the logs!

---

*Last Updated: October 2, 2025*
*Script Version: 2.0 (Performance and Error Handling Enhanced)*
