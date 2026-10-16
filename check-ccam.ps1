$ErrorActionPreference = "Stop"

# ============================================================
# Paths
# ============================================================

$ProjectRoot = $PSScriptRoot

$ConfigPath = Join-Path $ProjectRoot "config.json"
$StatePath  = Join-Path $ProjectRoot "state\workflow.json"

Write-Host "CCAM Automation"
Write-Host "Project root: $ProjectRoot"


# ============================================================
# Validate required files
# ============================================================

if (-not (Test-Path $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}

if (-not (Test-Path $StatePath)) {
    throw "Workflow state file not found: $StatePath"
}


# ============================================================
# Load configuration and workflow state
# ============================================================

$Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$State  = Get-Content $StatePath -Raw | ConvertFrom-Json

Write-Host ""
Write-Host "Configuration loaded successfully."
Write-Host "CCAM URL: $($Config.source.ccamPageUrl)"

Write-Host ""
Write-Host "Workflow state loaded successfully."
Write-Host "Status: $($State.status)"
Write-Host "Processed version: $($State.processedVersion)"
Write-Host "Detected version: $($State.detectedVersion)"


# ============================================================
# Download current CCAM page
# ============================================================

Write-Host ""
Write-Host "Downloading CCAM page..."

$InputDirectory = Join-Path $ProjectRoot "input"
$PagePath = Join-Path $InputDirectory "ccam-current.html"

# Create input directory if it does not exist.
if (-not (Test-Path $InputDirectory)) {
    New-Item -ItemType Directory -Path $InputDirectory | Out-Null
}

& curl.exe `
    -sS `
    -L `
    $Config.source.ccamPageUrl `
    -o $PagePath

if ($LASTEXITCODE -ne 0) {
    throw "Failed to download CCAM page."
}

if (-not (Test-Path $PagePath)) {
    throw "CCAM page was not created."
}

$PageContent = Get-Content $PagePath -Raw

if ($PageContent -notmatch 'CCAM') {
    throw "Downloaded page does not appear to be a CCAM page."
}

if ($PageContent -notmatch 'Version actuelle') {
    throw "Downloaded page does not contain the expected CCAM current-version section."
}

Write-Host "CCAM page downloaded successfully."
Write-Host "Saved to: $PagePath"


# ============================================================
# Calculate page hash
# ============================================================

Write-Host ""
Write-Host "Calculating page hash..."

$CurrentHash = (Get-FileHash $PagePath -Algorithm SHA256).Hash

Write-Host "Current page hash: $CurrentHash"
Write-Host "Previous page hash: $($State.pageHash)"


# ============================================================
# Determine whether AI analysis is required
# ============================================================

$PageChanged = $false

if ($null -eq $State.pageHash) {

    Write-Host ""
    Write-Host "No previous page hash exists."
    Write-Host "Initial CCAM analysis is required."

    $PageChanged = $true
}
elseif ($CurrentHash -ne $State.pageHash) {

    Write-Host ""
    Write-Host "CCAM page has changed."

    $PageChanged = $true
}
else {

    Write-Host ""
    Write-Host "CCAM page has not changed."

    exit 0
}


# ============================================================
# Ask Claude to determine current CCAM version and Excel URL
# ============================================================

if ($PageChanged) {

    Write-Host "AI analysis is required."
    Write-Host "Calling Claude..."

    $Prompt = @"
Read the file at:

$PagePath

This file contains the official CCAM download page.

Identify:
1. The full CURRENT CCAM version from the current-version section.
2. The Excel (.xls or .xlsx) download link belonging to that version.

Use the complete version stated in the current-version description,
for example V84.00, not a shortened download label such as V84.

Do not use a historical CCAM release.
"@

    # Load and compact the JSON schema.
    $SchemaPath = Join-Path $ProjectRoot "schemas\ccam-source-response.schema.json"

    if (-not (Test-Path $SchemaPath)) {
        throw "Claude response schema not found: $SchemaPath"
    }

    $Schema = Get-Content $SchemaPath -Raw
    $Schema = ($Schema | ConvertFrom-Json | ConvertTo-Json -Compress)

    # Escape quotes so the JSON survives the cmd.exe boundary.
    $SchemaEscaped = $Schema.Replace('"', '\"')

    # Escape quotes in the prompt and collapse it to one line.
    $PromptEscaped = $Prompt.Replace('"', '\"').Replace("`r", "").Replace("`n", " ")

    $Command = 'claude -p --output-format json --json-schema "' +
               $SchemaEscaped +
               '" "' +
               $PromptEscaped +
               '"'

    $ClaudeRawResult = cmd.exe /d /s /c $Command

    if ($LASTEXITCODE -ne 0) {
        throw "Claude invocation failed with exit code $LASTEXITCODE."
    }

    # Parse Claude Code's JSON response envelope.
    $ClaudeResponse = $ClaudeRawResult | ConvertFrom-Json

    if ($ClaudeResponse.is_error -eq $true) {
        throw "Claude returned an error."
    }

    if ($null -eq $ClaudeResponse.structured_output) {
        throw "Claude did not return structured output."
    }

    $DetectedVersion = $ClaudeResponse.structured_output.version
    $ExcelUrl = $ClaudeResponse.structured_output.excelUrl

    Write-Host ""
    Write-Host "Validated Claude response."
    Write-Host "CCAM version: $DetectedVersion"
    Write-Host "Excel URL: $ExcelUrl"
}


# ============================================================
# Build Excel download URL
# ============================================================

$BaseUri = [Uri]$Config.source.ccamPageUrl
$ExcelDownloadUrl = [Uri]::new($BaseUri, $ExcelUrl).AbsoluteUri

Write-Host ""
Write-Host "Excel download URL: $ExcelDownloadUrl"


# ============================================================
# Download Excel source file
# ============================================================

$ExcelFileName = [System.IO.Path]::GetFileName($ExcelUrl)
$ExcelPath = Join-Path $InputDirectory $ExcelFileName

Write-Host "Downloading CCAM Excel file..."

& curl.exe `
    -sS `
    -L `
    $ExcelDownloadUrl `
    -o $ExcelPath

if ($LASTEXITCODE -ne 0) {
    throw "Failed to download CCAM Excel file."
}

if (-not (Test-Path $ExcelPath)) {
    throw "CCAM Excel file was not created."
}

Write-Host "CCAM Excel downloaded successfully."
Write-Host "Saved to: $ExcelPath"

$ExcelHash = (Get-FileHash $ExcelPath -Algorithm SHA256).Hash

Write-Host "Excel SHA-256: $ExcelHash"


# ======================================================
# First CCAM installation
# ============================================================

if ($null -eq $State.processedVersion) {

    Write-Host ""
    Write-Host "No CCAM version has been processed yet."
    Write-Host "Creating workflow for $DetectedVersion"

    $State.detectedVersion = $DetectedVersion
    $State.status = "DETECTED"
    $State.detectedAt = (Get-Date).ToUniversalTime().ToString("o")
    $State.pageHash = $CurrentHash
    $State.artifacts.sourceFile = $ExcelPath
    $State.artifacts.sourceUrl = $ExcelDownloadUrl
    $State.artifacts.sourceHash = $ExcelHash

    $State |
        ConvertTo-Json -Depth 10 |
        Set-Content $StatePath -Encoding UTF8

    Write-Host ""
    Write-Host "Workflow created."
    Write-Host "Status: DETECTED"
    Write-Host "Version: $DetectedVersion"

    exit 0
}


# ============================================================
# Existing installation
# ============================================================

Write-Host ""
Write-Host "Current processed version: $($State.processedVersion)"
Write-Host "Detected website version: $DetectedVersion"

# We have not implemented existing-version / supersession
# handling yet. That will be the next workflow step.