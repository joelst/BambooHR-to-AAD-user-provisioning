#Requires -Version 7.0
<#
.SYNOPSIS
    Scans git diff output for PII and sensitive data patterns.

.DESCRIPTION
    Checks added/modified lines in a git diff for:
    - Email addresses matching real organizational domains (ERROR — blocks PR)
    - Email addresses with unrecognized domains (WARN — requires review)
    - Windows SIDs (ERROR — blocks PR)
    - Real-looking phone numbers (WARN — requires review)
    - Secret-bearing webhook/workflow URLs with live token or signature values (ERROR — blocks PR)

    Exits 1 if any errors are found, 0 otherwise.
    Designed to run in GitHub Actions on pull_request events and locally.

.PARAMETER BaseRef
    The base git reference to diff against.
    GitHub Actions: pass 'origin/${{ github.base_ref }}'.
    Local usage:    use 'origin/main' or a specific commit SHA.

.EXAMPLE
    # Check changes against main branch (local or CI)
    .\.github\scripts\Invoke-PiiScan.ps1 -BaseRef 'origin/main'
#>
[CmdletBinding()]
param(
    [string] $BaseRef = 'HEAD~1'
)

$ErrorActionPreference = 'Stop'

function Test-IsPlaceholderSecretValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $true
    }

    return $Value -match '(?i)(placeholder|example|sample|redacted|replace|changeme|todo|your-)' -or
    $Value -match '^[<{].+[>}]$'
}

# Real org email domains — ANY email matching these patterns blocks the PR.
# Update this list if the organization's domain ever changes.
$BlockedEmailDomains = @(
    'geckogreen\.com'
)

# Domains safe for test fixtures and documentation.
# Emails not matching these AND not matching a blocked domain trigger a review warning.
$AllowedEmailDomains = @(
    'contoso\.com',
    'example\.com',
    'fabrikam\.com',
    'microsoft\.com',
    'outlook\.com',
    'azure\.com',
    'windows\.net',
    'office\.com',
    'gmail\.com',
    'schemas\.',
    'graph\.microsoft',
    'powerautomate\.',
    'support\.microsoft',
    'learn\.microsoft',
    'go\.microsoft',
    'aka\.ms'
)

$findings = [System.Collections.Generic.List[hashtable]]::new()

$diffOutput = git diff "$BaseRef...HEAD" -- '*.ps1' '*.md' '*.json' '*.yml' '*.yaml' '*.txt' '*.bicep' '*.bicepparam' 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Warning "git diff returned exit code $LASTEXITCODE — falling back to two-dot diff."
    $diffOutput = git diff "$BaseRef" -- '*.ps1' '*.md' '*.json' '*.yml' '*.yaml' '*.txt' '*.bicep' '*.bicepparam' 2>&1
}

if (-not $diffOutput) {
    Write-Output 'ℹ️  No relevant file changes in diff. PII scan skipped.'
    exit 0
}

$currentFile = ''
$diffLineNum = 0

foreach ($line in $diffOutput) {
    if ($line -match '^\+\+\+ b/(.+)') {
        $currentFile = $Matches[1]
        continue
    }

    if ($line -notmatch '^\+' -or $line -match '^\+\+\+') { continue }

    $diffLineNum++
    $content = $line.Substring(1)
    $source = if ($currentFile) { "$currentFile (diff line $diffLineNum)" } else { "diff line $diffLineNum" }

    # --- Email address check ---
    $emailMatches = [regex]::Matches($content, '[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
    foreach ($m in $emailMatches) {
        $email = $m.Value
        $isSafe = $AllowedEmailDomains | Where-Object { $email -imatch $_ }
        if ($isSafe) { continue }

        $isBlocked = $BlockedEmailDomains | Where-Object { $email -imatch $_ }
        if ($isBlocked) {
            $findings.Add(@{ Severity = 'ERROR'; Source = $source; Message = "Real org email address '$email' — replace with a fictional domain (e.g. contoso.com)" })
        }
        else {
            $domain = ($email -split '@', 2)[-1]
            $findings.Add(@{ Severity = 'WARN'; Source = $source; Message = "Unrecognized email domain '@$domain' in '$email' — confirm this is not real employee data" })
        }
    }

    # --- Windows SID check ---
    if ($content -match 'S-1-5-21-\d+-\d+-\d+-\d+') {
        $findings.Add(@{ Severity = 'ERROR'; Source = $source; Message = 'Windows SID detected — machine-specific identifiers must not be committed' })
    }

    # --- Phone number check ---
    $phoneMatches = [regex]::Matches($content, '(?<!\w)(\+?1[\s\-.]?)?\(?\d{3}\)?[\s\-.]?\d{3}[\s\-.]?\d{4}(?!\w)')
    foreach ($m in $phoneMatches) {
        $phone = $m.Value.Trim()
        $digits = [regex]::Replace($phone, '[^\d]', '')
        if ($digits.Length -lt 10) { continue }

        $last10   = $digits.Substring($digits.Length - 10)
        $areaCode = $last10.Substring(0, 3)
        $exchange = $last10.Substring(3, 3)
        $tollFree = '800', '888', '877', '866', '855', '844', '833', '822'

        # 555 area code or 555 exchange = reserved/fictional; toll-free = safe in docs
        if ($areaCode -eq '555' -or $exchange -eq '555' -or $areaCode -in $tollFree) { continue }

        $findings.Add(@{ Severity = 'WARN'; Source = $source; Message = "Real-looking phone number '$phone' — confirm this is fictional test data" })
    }

    # --- Secret-bearing URL check ---
    $urlMatches = [regex]::Matches($content, 'https?://[^\s''"<>]+')
    foreach ($m in $urlMatches) {
        $candidateUrl = $m.Value.TrimEnd('.', ',', ';', ')')
        $uri = $null
        if (-not [System.Uri]::TryCreate($candidateUrl, [System.UriKind]::Absolute, [ref]$uri)) { continue }
        if ([string]::IsNullOrWhiteSpace($uri.Query)) { continue }

        $queryString = $uri.Query.TrimStart('?')
        if ([string]::IsNullOrWhiteSpace($queryString)) { continue }

        foreach ($pair in ($queryString -split '&')) {
            $parts = $pair -split '=', 2
            if ($parts.Count -ne 2) { continue }

            $parameterName = $parts[0].ToLowerInvariant()
            if ($parameterName -notin @('token', 'sig', 'signature')) { continue }

            $parameterValue = [System.Uri]::UnescapeDataString($parts[1])
            if ($parameterValue.Length -lt 12) { continue }
            if (Test-IsPlaceholderSecretValue -Value $parameterValue) { continue }

            if ($uri.Host -imatch 'azure-automation\.net$' -and $parameterName -eq 'token') {
                $findings.Add(@{ Severity = 'ERROR'; Source = $source; Message = 'Azure Automation webhook URL detected — this bearer-secret URL must never be committed' })
            }
            else {
                $redactedUrl = "$($uri.Scheme)://$($uri.Host)$($uri.AbsolutePath)?$parameterName=<redacted>"
                $findings.Add(@{ Severity = 'ERROR'; Source = $source; Message = "Secret-bearing URL parameter '$parameterName' detected in '$redactedUrl' — replace with a placeholder or move it to a secure variable" })
            }

            break
        }
    }
}

# --- Output ---
Write-Output ''

if ($findings.Count -eq 0) {
    Write-Output '✅ PII scan passed — no issues found.'
    exit 0
}

$errors   = @($findings | Where-Object { $_.Severity -eq 'ERROR' })
$warnings = @($findings | Where-Object { $_.Severity -eq 'WARN'  })

Write-Output '──────────────────────────────────────────────'
Write-Output ' PII / Sensitive Data Scan Results'
Write-Output '──────────────────────────────────────────────'
Write-Output ''

foreach ($f in $findings) {
    Write-Output "[$($f.Severity)] $($f.Source)"
    Write-Output "       $($f.Message)"
    Write-Output ''
}

Write-Output '──────────────────────────────────────────────'

if ($errors.Count -gt 0) {
    Write-Output "❌ $($errors.Count) error(s) found — PR is blocked until resolved."
    Write-Output '   Replace real org emails / SIDs with fictional equivalents and remove live bearer-secret URLs.'
    Write-Output ''
    exit 1
}

Write-Output "⚠️  $($warnings.Count) warning(s) — review each item and confirm it is safe test data."
Write-Output '   Warnings do not block the PR but require manual sign-off.'
Write-Output ''
exit 0
