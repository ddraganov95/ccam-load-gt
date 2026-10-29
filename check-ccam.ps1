param(
    [Alias("v")]
    [switch]$Validate,
    [switch]$Test
)

$ErrorActionPreference = "Stop"

$Config = $null
$State = $null
$ProjectRoot = $PSScriptRoot
$ConfigPath = $null
$StatePath = $null

function Send-CcamFailureEmail {
    param(
        [Parameter(Mandatory)]
        [string]$Recipient,

        [Parameter(Mandatory)]
        [string]$Reason,

        [string]$WorkflowStatus = "UNKNOWN",
        [string]$DetectedVersion = "UNKNOWN"
    )

    if ([string]::IsNullOrWhiteSpace($Recipient)) {
        Write-Warning "Failure email not sent because no technical contact is configured."
        return
    }

    $Outlook = $null
    $Mail = $null

    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)
        $Mail.To = $Recipient
        $Mail.Subject = "CCAM automation failed on $env:COMPUTERNAME"
        $Mail.Body = @"
The CCAM automation requires attention.

Reason:
$Reason

Workflow status:
$WorkflowStatus

Detected CCAM version:
$DetectedVersion

Computer:
$env:COMPUTERNAME

User:
$env:USERNAME

Time:
$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))

Script:
$PSCommandPath
"@
        $Mail.Send()
        Write-Host ""
        Write-Host "Failure notification submitted to Outlook for $Recipient"
    }
    catch {
        Write-Warning "The failure notification could not be sent."
        Write-Warning $_.Exception.Message
    }
    finally {
        if ($null -ne $Mail) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail)
        }
        if ($null -ne $Outlook) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

try {
    if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
        throw "Unable to determine the project root from PSScriptRoot."
    }

    Set-Location -LiteralPath $ProjectRoot

    $ConfigPath = Join-Path $ProjectRoot "config.json"

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Configuration file not found: $ConfigPath"
}

try {
    $Config = Get-Content -LiteralPath $ConfigPath -Raw |
        ConvertFrom-Json
}
catch {
    throw "config.json contains invalid JSON: $($_.Exception.Message)"
}

$StatePath = Join-Path $ProjectRoot $Config.paths.state
$TransformSchemaPath = Join-Path $ProjectRoot $Config.paths.transformSchema
$TransformSkillPath = Join-Path $ProjectRoot $Config.paths.transformSkill
$ReviewReportScriptPath = Join-Path $ProjectRoot $Config.paths.reviewReportScript

    Write-Host "CCAM Automation"
    Write-Host "Project root: $ProjectRoot"
    Write-Host "Working directory: $((Get-Location).Path)"
    if (-not (Test-Path -LiteralPath $StatePath)) {
        throw "Workflow state file not found: $StatePath"
    }
    if (-not (Test-Path -LiteralPath $TransformSkillPath)) {
        throw "CCAM transformation skill not found: $TransformSkillPath"
    }
    if (-not (Test-Path -LiteralPath $ReviewReportScriptPath)) {
    throw "Review report generator not found: $ReviewReportScriptPath"
    }
    if (-not (Test-Path -LiteralPath $TransformSchemaPath)) {
        throw "Transformation response schema not found: $TransformSchemaPath"
    }

    Write-Host "Transformation skill found: $TransformSkillPath"

    try {
        $State = Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json
    }
    catch {
        throw "Workflow state file contains invalid JSON: $($_.Exception.Message)"
    }

    if ($null -eq $Config.connector) {
        throw "The connector section is missing from config.json."
    }
    if ([string]::IsNullOrWhiteSpace([string]$Config.connector.technicalContact)) {
        throw "connector.technicalContact is missing from config.json."
    }
    if ([string]::IsNullOrWhiteSpace([string]$Config.connector.reviewer)) {
        throw "connector.reviewer is missing from config.json."
    }
    if ([string]::IsNullOrWhiteSpace([string]$Config.source.ccamPageUrl)) {
        throw "source.ccamPageUrl is missing from config.json."
    }

    Write-Host ""
    Write-Host "Configuration loaded successfully."
    Write-Host "CCAM URL: $($Config.source.ccamPageUrl)"
    Write-Host ""
    Write-Host "Workflow state loaded successfully."
    Write-Host "Status: $($State.status)"
    Write-Host "Processed version: $($State.processedVersion)"
    Write-Host "Detected version: $($State.detectedVersion)"

    if ($Validate) {
        Write-Host ""
        Write-Host "Validation mode enabled."
        Write-Host "Checking Claude CLI and project skill visibility..."

        $ClaudeCommand = Get-Command claude -ErrorAction SilentlyContinue
        if ($null -eq $ClaudeCommand) {
            throw "Claude CLI was not found in PATH."
        }
        Write-Host "Claude CLI found: $($ClaudeCommand.Source)"

        $SkillVisibilityPrompt = @"
Inspect the current project directory.

Confirm whether this file exists and is readable:

.claude/skills/ccam-transform/SKILL.md

Do not transform any file.
Do not modify any file.

Return exactly:
SKILL_FOUND

If the file cannot be read, return exactly:
SKILL_NOT_FOUND
"@

        $SkillVisibilityResult = $SkillVisibilityPrompt | & claude -p --output-format text
        if ($LASTEXITCODE -ne 0) {
            throw "Claude visibility test failed with exit code $LASTEXITCODE."
        }
        $SkillVisibilityResult = $SkillVisibilityResult.Trim()
        Write-Host "Claude skill visibility result: $SkillVisibilityResult"
        if ($SkillVisibilityResult -ne "SKILL_FOUND") {
            throw "Claude could not confirm access to the project transformation skill."
        }

        Write-Host "Checking transformation response schema..."
        try {
            $null = Get-Content -LiteralPath $TransformSchemaPath -Raw | ConvertFrom-Json
        }
        catch {
            throw "Transformation response schema contains invalid JSON: $($_.Exception.Message)"
        }

        Write-Host "Transformation response schema is valid."
        Write-Host ""
        Write-Host "Validation completed successfully."
        exit 0
    }

    function Invoke-CcamTransformation {
        param(
            [Parameter(Mandatory)]
            [string]$SourceExcelPath,

            [Parameter(Mandatory)]
            [string]$Version
        )

        if (-not (Test-Path -LiteralPath $SourceExcelPath)) {
            throw "Source Excel file not found: $SourceExcelPath"
        }

        $OutputDirectory = Join-Path $ProjectRoot $Config.paths.outputDirectory
        if (-not (Test-Path -LiteralPath $OutputDirectory)) {
            New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
        }

        $SafeVersion = $Version -replace '[^A-Za-z0-9._-]', '_'
        $OutputPath = Join-Path $OutputDirectory "CCAM_${SafeVersion}_normalise.csv"

        try {
            $Schema = Get-Content -LiteralPath $TransformSchemaPath -Raw |
                ConvertFrom-Json |
                ConvertTo-Json -Depth 20 -Compress
        }
        catch {
            throw "Unable to load transformation schema: $($_.Exception.Message)"
        }

        $Prompt = @"
Use the ccam-transform skill available in this project.

Source Excel:
$SourceExcelPath

Required output:
$OutputPath

Transform the complete workbook and write the GraphTalk output
to the exact requested path.

Run every validation required by the skill.
Correct deterministic failures automatically.
Return the complete structured transformation report required by the skill and the supplied JSON schema.
"@

        Write-Host ""
        Write-Host "Calling Claude for CCAM transformation..."
        Write-Host "Source: $SourceExcelPath"
        Write-Host "Output: $OutputPath"

        $ClaudeRawResult = $Prompt |
    & claude `
        -p `
        --dangerously-skip-permissions `
        --output-format json `
        --json-schema $Schema

        if ($LASTEXITCODE -ne 0) {
            throw "Claude transformation failed with exit code $LASTEXITCODE."
        }

        try {
            $ClaudeResponse = $ClaudeRawResult | ConvertFrom-Json
        }
        catch {
            throw "Claude returned invalid JSON: $($_.Exception.Message)"
        }

        if ($ClaudeResponse.is_error -eq $true) {
            throw "Claude returned a transformation error."
        }
        if ($null -eq $ClaudeResponse.structured_output) {
            throw "Claude did not return structured transformation output."
        }

        $Result = $ClaudeResponse.structured_output
        Write-Host ""
        Write-Host "Claude structured result:"
        Write-Host ($Result | ConvertTo-Json -Depth 20)

        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw "Claude did not create the required output: $OutputPath"
        }

        return $Result
    }

 if ($State.status -eq "AWAITING_APPROVAL") {
    Write-Host ""
    Write-Warning "CCAM version $($State.detectedVersion) is awaiting approval."

    if ($null -ne $State.artifacts.candidateFile) {
        Write-Host "Candidate file: $($State.artifacts.candidateFile)"
    }

    if ($null -ne $State.review.reportFile) {
        Write-Host "Review report: $($State.review.reportFile)"
    }

    Write-Host "The daily CCAM page check will continue."
}

    $InputDirectory = Join-Path $ProjectRoot $Config.paths.inputDirectory
    if (-not (Test-Path -LiteralPath $InputDirectory)) {
        New-Item -ItemType Directory -Path $InputDirectory | Out-Null
    }

    if ($Test) {
        Write-Host ""
        Write-Warning "TEST MODE ENABLED"
        Write-Host "Ameli page check and download will be bypassed."

        if ($null -eq $Config.test) {
            throw "The test section is missing from config.json."
        }

        $DetectedVersion = [string]$Config.test.version
        if ([string]::IsNullOrWhiteSpace($DetectedVersion)) {
            throw "test.version is missing from config.json."
        }
        if ([string]::IsNullOrWhiteSpace([string]$Config.test.excelFile)) {
            throw "test.excelFile is missing from config.json."
        }

        $ExcelPath = Join-Path $ProjectRoot ([string]$Config.test.excelFile)
        if (-not (Test-Path -LiteralPath $ExcelPath)) {
            throw "Test Excel file not found: $ExcelPath"
        }

        $ExcelHash = (Get-FileHash -LiteralPath $ExcelPath -Algorithm SHA256).Hash
        $CurrentHash = "TEST-$DetectedVersion"
        $ExcelDownloadUrl = "LOCAL_TEST_FILE"

        Write-Host "Test CCAM version: $DetectedVersion"
        Write-Host "Test Excel: $ExcelPath"
        Write-Host "Excel SHA-256: $ExcelHash"
    }
    else {
        Write-Host ""
        Write-Host "Downloading CCAM page..."

        $PagePath = Join-Path $InputDirectory "ccam-current.html"
        $CurlErrorPath = Join-Path $InputDirectory "ccam-curl-error.txt"
        Remove-Item -LiteralPath $CurlErrorPath -Force -ErrorAction SilentlyContinue

        & curl.exe `
            -f `
            -sS `
            -L `
            --connect-timeout 10 `
            --max-time 30 `
            --retry 2 `
            --retry-delay 2 `
            $Config.source.ccamPageUrl `
            -o $PagePath `
            2> $CurlErrorPath

        $CurlExitCode = $LASTEXITCODE
        if ($CurlExitCode -ne 0) {
            $CurlError = ""
            if (Test-Path -LiteralPath $CurlErrorPath) {
                $CurlError = (Get-Content -LiteralPath $CurlErrorPath -Raw).Trim()
            }
            if ([string]::IsNullOrWhiteSpace($CurlError)) {
                $CurlError = "curl returned no additional error information."
            }
            throw "Failed to download CCAM page. curl exit code: $CurlExitCode. Details: $CurlError"
        }

        Remove-Item -LiteralPath $CurlErrorPath -Force -ErrorAction SilentlyContinue

        if (-not (Test-Path -LiteralPath $PagePath)) {
            throw "CCAM page was not created."
        }

        $PageContent = Get-Content -LiteralPath $PagePath -Raw
        if ($PageContent -notmatch 'CCAM') {
            throw "Downloaded page does not appear to be a CCAM page."
        }
        if ($PageContent -notmatch 'Version actuelle') {
            throw "Downloaded page does not contain the expected CCAM current-version section."
        }

        Write-Host "CCAM page downloaded successfully."
        Write-Host "Saved to: $PagePath"
        Write-Host ""
        Write-Host "Calculating page hash..."

        $CurrentHash = (Get-FileHash -LiteralPath $PagePath -Algorithm SHA256).Hash
        Write-Host "Current page hash: $CurrentHash"
        Write-Host "Previous page hash: $($State.pageHash)"

        if ($null -eq $State.pageHash) {
            Write-Host ""
            Write-Host "No previous page hash exists."
            Write-Host "Initial CCAM analysis is required."
        }
        elseif ($CurrentHash -ne $State.pageHash) {
            Write-Host ""
            Write-Host "CCAM page has changed."
        }
        else {
            Write-Host ""
            Write-Host "CCAM page has not changed."
            exit 0
        }

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

        $SchemaPath = Join-Path $ProjectRoot "schemas\ccam-source-response.schema.json"
        if (-not (Test-Path -LiteralPath $SchemaPath)) {
            throw "Claude response schema not found: $SchemaPath"
        }

        try {
            $Schema = Get-Content -LiteralPath $SchemaPath -Raw |
                ConvertFrom-Json |
                ConvertTo-Json -Depth 20 -Compress
        }
        catch {
            throw "Unable to load CCAM source response schema: $($_.Exception.Message)"
        }

        $ClaudeRawResult = $Prompt |
            & claude `
                -p `
                --output-format json `
                --json-schema $Schema

        if ($LASTEXITCODE -ne 0) {
            throw "Claude invocation failed with exit code $LASTEXITCODE."
        }

        try {
            $ClaudeResponse = $ClaudeRawResult | ConvertFrom-Json
        }
        catch {
            throw "Claude returned invalid JSON: $($_.Exception.Message)"
        }

        if ($ClaudeResponse.is_error -eq $true) {
            throw "Claude returned an error."
        }
        if ($null -eq $ClaudeResponse.structured_output) {
            throw "Claude did not return structured output."
        }

        $DetectedVersion = [string]$ClaudeResponse.structured_output.version
        $ExcelUrl = [string]$ClaudeResponse.structured_output.excelUrl

        if ([string]::IsNullOrWhiteSpace($DetectedVersion)) {
            throw "Claude did not return the current CCAM version."
        }
        if ([string]::IsNullOrWhiteSpace($ExcelUrl)) {
            throw "Claude did not return the CCAM Excel URL."
        }

        Write-Host ""
        Write-Host "Validated Claude response."
        Write-Host "CCAM version: $DetectedVersion"
        Write-Host "Excel URL: $ExcelUrl"

        $BaseUri = [Uri]$Config.source.ccamPageUrl
        $ExcelDownloadUrl = [Uri]::new($BaseUri, $ExcelUrl).AbsoluteUri
        Write-Host ""
        Write-Host "Excel download URL: $ExcelDownloadUrl"

        $ExcelFileName = [System.IO.Path]::GetFileName(([Uri]$ExcelDownloadUrl).AbsolutePath)
        if ([string]::IsNullOrWhiteSpace($ExcelFileName)) {
            throw "Unable to determine the Excel filename from URL: $ExcelDownloadUrl"
        }

        $ExcelPath = Join-Path $InputDirectory $ExcelFileName
        Write-Host "Downloading CCAM Excel file..."

        & curl.exe `
            -f `
            -sS `
            -L `
            --connect-timeout 10 `
            --max-time 120 `
            --retry 2 `
            --retry-delay 2 `
            $ExcelDownloadUrl `
            -o $ExcelPath

        $ExcelCurlExitCode = $LASTEXITCODE
        if ($ExcelCurlExitCode -ne 0) {
            throw "Failed to download CCAM Excel file. curl exit code: $ExcelCurlExitCode."
        }
        if (-not (Test-Path -LiteralPath $ExcelPath)) {
            throw "CCAM Excel file was not created."
        }

        Write-Host "CCAM Excel downloaded successfully."
        Write-Host "Saved to: $ExcelPath"
        $ExcelHash = (Get-FileHash -LiteralPath $ExcelPath -Algorithm SHA256).Hash
        Write-Host "Excel SHA-256: $ExcelHash"
    }

    if ($null -eq $State.processedVersion) {
        Write-Host ""
        Write-Host "No CCAM version has been processed yet."
        Write-Host "Creating workflow for $DetectedVersion"

        if ($null -eq $State.artifacts) {
            $State | Add-Member -MemberType NoteProperty -Name artifacts -Value ([pscustomobject]@{})
        }

        $State.detectedVersion = $DetectedVersion
        $State.status = "DETECTED"
        $State.detectedAt = (Get-Date).ToUniversalTime().ToString("o")
        $State.pageHash = $CurrentHash
        $State.artifacts.sourceFile = $ExcelPath
        $State.artifacts.sourceUrl = $ExcelDownloadUrl
        $State.artifacts.sourceHash = $ExcelHash

        $State | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StatePath -Encoding UTF8

        Write-Host ""
        Write-Host "Workflow created."
        Write-Host "Status: DETECTED"
        Write-Host "Version: $DetectedVersion"

        $State.status = "TRANSFORMING"
        $State | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StatePath -Encoding UTF8

        $TransformResult = Invoke-CcamTransformation `
            -SourceExcelPath $ExcelPath `
            -Version $DetectedVersion

        if ($TransformResult.status -ne "PASS") {
            $State.status = "TRANSFORMATION_FAILED"
            $State | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StatePath -Encoding UTF8
            throw "CCAM transformation failed validation."
        }

$CandidateFile = [string]$TransformResult.outputPath

if ([string]::IsNullOrWhiteSpace($CandidateFile)) {
    throw "Claude returned PASS without an outputPath."
}

if (-not (Test-Path -LiteralPath $CandidateFile)) {
    throw "Candidate output file does not exist: $CandidateFile"
}

$CandidateHash = (
    Get-FileHash `
        -LiteralPath $CandidateFile `
        -Algorithm SHA256
).Hash
$ReviewReport = & $ReviewReportScriptPath `
    -ProjectRoot $ProjectRoot `
    -TransformResult $TransformResult `
    -Version $DetectedVersion `
    -CandidateFile $CandidateFile `
    -CandidateHash $CandidateHash

if ($null -eq $ReviewReport) {
    throw "Review report generator returned no result."
}

if (-not (Test-Path -LiteralPath $ReviewReport.reportPath)) {
    throw "Review report was not created: $($ReviewReport.reportPath)"
}
$State.status = "AWAITING_APPROVAL"

$State.artifacts.candidateFile = $CandidateFile
$State.artifacts.candidateHash = $CandidateHash

$State.review.reportFile = $ReviewReport.reportPath
$State.review.reportHash = $ReviewReport.reportHash

$State.transformation = $TransformResult
$State.transformedAt = (Get-Date).ToUniversalTime().ToString("o")

$State |
    ConvertTo-Json -Depth 20 |
    Set-Content -LiteralPath $StatePath -Encoding UTF8

        Write-Host ""
        Write-Host "Transformation completed successfully."
        Write-Host "Status: AWAITING_APPROVAL"
        Write-Host "Candidate file: $CandidateFile"
        Write-Host "Candidate SHA-256: $CandidateHash"
        Write-Host "Review report: $($ReviewReport.reportPath)"
        Write-Host "Human review items: $($ReviewReport.reviewCount)"
        Write-Host "Records: $($TransformResult.totalRecords)"
        exit 0
    }

    Write-Host ""
    Write-Host "Current processed version: $($State.processedVersion)"
    Write-Host "Detected website version: $DetectedVersion"
}
catch {
    $FailureReason = $_.Exception.Message
    $FailurePosition = $_.InvocationInfo.PositionMessage
    $FailureDetails = $_ | Out-String

    Write-Host ""
    Write-Warning "CCAM automation failed: $FailureReason"
    Write-Warning "Location: $FailurePosition"

    $CurrentWorkflowStatus = "UNKNOWN"
    $CurrentDetectedVersion = "UNKNOWN"
    $TechnicalContact = $null

    if ($null -ne $State) {
        if (-not [string]::IsNullOrWhiteSpace([string]$State.status)) {
            $CurrentWorkflowStatus = [string]$State.status
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$State.detectedVersion)) {
            $CurrentDetectedVersion = [string]$State.detectedVersion
        }
    }

    if (
        $null -ne $Config -and
        $null -ne $Config.connector -and
        -not [string]::IsNullOrWhiteSpace([string]$Config.connector.technicalContact)
    ) {
        $TechnicalContact = [string]$Config.connector.technicalContact
    }

    $EmailReason = @"
$FailureReason

PowerShell location:
$FailurePosition

Full error:
$FailureDetails
"@

    if (-not [string]::IsNullOrWhiteSpace($TechnicalContact)) {
        Send-CcamFailureEmail `
            -Recipient $TechnicalContact `
            -Reason $EmailReason `
            -WorkflowStatus $CurrentWorkflowStatus `
            -DetectedVersion $CurrentDetectedVersion
    }
    else {
        Write-Warning "Failure email could not be sent because connector.technicalContact was not loaded."
    }

    exit 1
}
