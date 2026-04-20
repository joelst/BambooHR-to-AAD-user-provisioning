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

$diffOutput = git diff "$BaseRef...HEAD" -- '*.ps1' '*.md' '*.json' '*.yml' '*.yaml' '*.txt' 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Warning "git diff returned exit code $LASTEXITCODE — falling back to two-dot diff."
    $diffOutput = git diff "$BaseRef" -- '*.ps1' '*.md' '*.json' '*.yml' '*.yaml' '*.txt' 2>&1
}

if (-not $diffOutput) {
    Write-Host 'ℹ️  No relevant file changes in diff. PII scan skipped.' -ForegroundColor Cyan
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
}

# --- Output ---
Write-Host ''

if ($findings.Count -eq 0) {
    Write-Host '✅ PII scan passed — no issues found.' -ForegroundColor Green
    exit 0
}

$errors   = @($findings | Where-Object { $_.Severity -eq 'ERROR' })
$warnings = @($findings | Where-Object { $_.Severity -eq 'WARN'  })

Write-Host '──────────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host ' PII / Sensitive Data Scan Results' -ForegroundColor Cyan
Write-Host '──────────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host ''

foreach ($f in $findings) {
    $color = if ($f.Severity -eq 'ERROR') { 'Red' } else { 'Yellow' }
    Write-Host "[$($f.Severity)] $($f.Source)" -ForegroundColor $color
    Write-Host "       $($f.Message)" -ForegroundColor $color
    Write-Host ''
}

Write-Host '──────────────────────────────────────────────' -ForegroundColor DarkGray

if ($errors.Count -gt 0) {
    Write-Host "❌ $($errors.Count) error(s) found — PR is blocked until resolved." -ForegroundColor Red
    Write-Host '   Replace real org emails / SIDs with fictional equivalents (e.g. contoso.com).' -ForegroundColor Red
    Write-Host ''
    exit 1
}

Write-Host "⚠️  $($warnings.Count) warning(s) — review each item and confirm it is safe test data." -ForegroundColor Yellow
Write-Host '   Warnings do not block the PR but require manual sign-off.' -ForegroundColor Yellow
Write-Host ''
exit 0
