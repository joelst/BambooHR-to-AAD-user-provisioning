import { execFile } from "node:child_process";
import { joinSession } from "@github/copilot-sdk/extension";

const isWindows = process.platform === "win32";
const shell = isWindows ? "pwsh" : "pwsh";

function runPowerShell(command) {
    return new Promise((resolve) => {
        execFile(
            shell,
            ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", command],
            { timeout: 60000 },
            (err, stdout, stderr) => {
                if (err && !stdout) resolve(`Error: ${stderr || err.message}`);
                else resolve(stdout || stderr || "No output.");
            }
        );
    });
}

const SECURITY_PATTERNS = [
    { pattern: "geckogreen\\.com", severity: "Critical", message: "Real organizational email domain — replace with 'contoso.com' or another fictional domain in test fixtures" },
    { pattern: "ConvertTo-SecureString.*-AsPlainText", severity: "High", message: "Plaintext password conversion — use Managed Identity or encrypted automation variables instead" },
    { pattern: "\\$(?:password|secret|apikey|token)\\s*=\\s*['\"]", severity: "Critical", message: "Hardcoded credential detected — secrets must come from Azure Automation encrypted variables" },
    { pattern: "Read-Host|\\$host\\.UI\\.Prompt", severity: "High", message: "Interactive prompt — Azure Automation runbooks run unattended" },
    { pattern: "Invoke-Expression|iex\\s", severity: "Critical", message: "Invoke-Expression enables script injection — use & (call operator) instead" },
    { pattern: "Write-Host\\b", severity: "Medium", message: "Write-Host bypasses the pipeline — use Write-Information, Write-Output, or Write-Verbose" },
    { pattern: "\\[switch\\]", severity: "High", message: "[switch] parameters do not work in Azure Automation — use [bool] with a default value" },
    { pattern: "\\$global:", severity: "Medium", message: "Global variable — pass data through parameters and return values instead" },
    { pattern: "Invoke-RestMethod|Invoke-WebRequest", severity: "Medium", message: "Raw REST call — prefer Microsoft Graph SDK cmdlets (parameterized) for Graph API calls" },
    { pattern: "\\$env:.*(?:KEY|SECRET|PASSWORD|TOKEN)", severity: "High", message: "Accessing credential from environment variable — use Azure Automation encrypted variables" },
    { pattern: "Export-Csv.*\\$(?:bhr|employee)", severity: "Medium", message: "Exporting employee data to file — PII must be processed in-memory only" },
    { pattern: "ConfirmImpact\\s*=\\s*['\"](?:High|Medium|Low)['\"]", severity: "High", message: "ConfirmImpact must be None for Azure Automation — any other value causes the runbook to hang" },
    { pattern: "Connect-(?:MgGraph|AzAccount|ExchangeOnline)(?:\\s|$|\\|)", exclude: "-(?:Identity|ManagedIdentity)|Get-Command|Disconnect-|Write-PSLog|Write-(?:Verbose|Information|Debug|Warning|Output)", severity: "High", message: "Connection without -Identity or -ManagedIdentity — will fail in Azure Automation unattended context" },
    { pattern: "ConvertFrom-SecureString", severity: "Medium", message: "ConvertFrom-SecureString produces a recoverable string on the same machine — avoid persisting credentials" },
    { pattern: "Start-Transcript", severity: "Medium", message: "Start-Transcript captures all output including secrets — avoid in runbooks handling credentials" },
    { pattern: "Send-MailMessage", severity: "Medium", message: "Send-MailMessage is deprecated and uses unauthenticated SMTP — use Send-MgUserMail with Graph API instead" },
    { pattern: "\\$ErrorActionPreference\\s*=\\s*['\"]SilentlyContinue['\"]", severity: "Medium", message: "SilentlyContinue hides errors including security failures — use Stop or Continue with explicit handling" },
    { pattern: "-SkipCertificateCheck|ServerCertificateValidationCallback", severity: "High", message: "Disabling TLS certificate validation enables man-in-the-middle attacks" },
    { pattern: "Write-(?:Output|Information|Warning|Error|Verbose|Debug).*\\$.*[Pp]ass(?:word)?", severity: "High", message: "Logging a variable that likely contains a password — secrets must never appear in output streams" },
];

const session = await joinSession({
    hooks: {
        onPostToolUse: async (input) => {
            if (input.toolName !== "edit" && input.toolName !== "create") return;

            const filePath = String(input.toolArgs?.path || "");
            if (!filePath.toLowerCase().endsWith(".ps1")) return;

            return {
                additionalContext:
                    "A .ps1 file was modified. Consider security implications: " +
                    "credential handling, PII exposure, input validation, " +
                    "ShouldProcess gates, and Azure Automation constraints. " +
                    "Use the run_security_review tool before committing if this is a substantive change.",
            };
        },
    },
    tools: [
        {
            name: "run_security_review",
            description:
                "Runs a security-focused review of a PowerShell script. " +
                "Executes PSScriptAnalyzer and checks for security anti-patterns " +
                "specific to Azure Automation runbooks handling PII (hardcoded " +
                "credentials, interactive prompts, [switch] params, raw REST " +
                "calls, PII exports, Write-Host, Invoke-Expression).",
            parameters: {
                type: "object",
                properties: {
                    path: {
                        type: "string",
                        description: "Absolute path to the .ps1 file to review",
                    },
                },
                required: ["path"],
            },
            handler: async (args) => {
                const filePath = args.path;

                if (!filePath.toLowerCase().endsWith(".ps1")) {
                    return {
                        textResultForLlm: "Skipped: not a PowerShell file.",
                        resultType: "rejected",
                    };
                }

                // Run PSScriptAnalyzer
                const analyzerCmd = [
                    `try {`,
                    `  Import-Module PSScriptAnalyzer -ErrorAction Stop`,
                    `  $results = Invoke-ScriptAnalyzer -Path '${filePath.replace(/'/g, "''")}' -Severity Warning,Error`,
                    `  if ($results) { $results | Format-Table -Property Severity,RuleName,Line,Message -AutoSize -Wrap | Out-String -Width 200 }`,
                    `  else { 'PSScriptAnalyzer: No warnings or errors.' }`,
                    `} catch { "PSScriptAnalyzer not available: $($_.Exception.Message)" }`,
                ].join("\n");

                const analyzerOutput = await runPowerShell(analyzerCmd);

                // Run pattern-based security checks
                const patternCmd = [
                    `$lines = Get-Content -Path '${filePath.replace(/'/g, "''")}'`,
                    `$findings = @()`,
                    ...SECURITY_PATTERNS.map((p) => {
                        if (p.exclude) {
                            return `$lines | Select-String -Pattern '${p.pattern.replace(/'/g, "''")}' | Where-Object { $_.Line -notmatch '${p.exclude.replace(/'/g, "''")}' } | ForEach-Object { $findings += "[${p.severity}] Line $($_.LineNumber): ${p.message.replace(/'/g, "''")}" }`;
                        }
                        return `$lines | Select-String -Pattern '${p.pattern.replace(/'/g, "''")}' | ForEach-Object { $findings += "[${p.severity}] Line $($_.LineNumber): ${p.message.replace(/'/g, "''")}" }`;
                    }),
                    `if ($findings.Count -gt 0) { $findings -join [Environment]::NewLine }`,
                    `else { 'Security patterns: No issues found.' }`,
                ].join("\n");

                const patternOutput = await runPowerShell(patternCmd);

                const report = [
                    "=== PSScriptAnalyzer Results ===",
                    analyzerOutput.trim(),
                    "",
                    "=== Security Pattern Scan ===",
                    patternOutput.trim(),
                ].join("\n");

                const hasIssues =
                    !analyzerOutput.includes("No warnings or errors") ||
                    !patternOutput.includes("No issues found");

                return {
                    textResultForLlm: report,
                    resultType: hasIssues ? "failure" : "success",
                };
            },
        },
    ],
});

await session.log("Security review extension loaded");
