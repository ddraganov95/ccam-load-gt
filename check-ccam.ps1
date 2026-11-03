param(
    [Alias("v")]
    [switch]$Validate,

    [switch]$Test
)
$LogFile = Join-Path $PSScriptRoot "task.log"
Start-Transcript -Path $LogFile -Force

Write-Host "=== CCAM START $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

$ErrorActionPreference = "Stop"
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

$OutputEncoding = $Utf8NoBom
[Console]::InputEncoding = $Utf8NoBom
[Console]::OutputEncoding = $Utf8NoBom
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
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Mail
            )
        }

        if ($null -ne $Outlook) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Outlook
            )
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Send-CcamReviewEmail {
    param(
        [Parameter(Mandatory)]
        [string]$Recipient,

        [Parameter(Mandatory)]
        [string]$Version,

        [Parameter(Mandatory)]
        [string]$ApplicabilityDate,

        [Parameter(Mandatory)]
        [string]$ReportPath
    )

    $Outlook = $null
    $Mail = $null

    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)

        $Mail.To = $Recipient
        $Mail.Subject = "CCAM $Version ready for review"

        $Mail.Body = @"
CCAM $Version is ready for review.

Applicable date:
$ApplicabilityDate

Review report:
$ReportPath

Open the review report to review the generated GraphTalk candidate.
"@

        $Mail.Send()

        Write-Host ""
        Write-Host "Review notification submitted to Outlook for $Recipient"
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

function Get-CcamWorkbookMetadata {
    param(
        [Parameter(Mandatory)]
        [string]$ExcelPath,

        [Parameter(Mandatory)]
        [string]$WorksheetName,

        [Parameter(Mandatory)]
        [string]$VersionCell
    )

    if (-not (Test-Path -LiteralPath $ExcelPath)) {
        throw "CCAM workbook not found: $ExcelPath"
    }

    $Excel = $null
    $Workbook = $null
    $Worksheet = $null
    $VersionRange = $null

    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        $Workbook = $Excel.Workbooks.Open(
            $ExcelPath,
            0,
            $true
        )

        $Worksheet = $Workbook.Worksheets.Item(
            $WorksheetName
        )

        $VersionRange = $Worksheet.Range(
            $VersionCell
        )

        $VersionHeader = [string]$VersionRange.Text

        if ([string]::IsNullOrWhiteSpace($VersionHeader)) {
            throw "CCAM version cell $WorksheetName!$VersionCell is empty."
        }

        $NormalizedHeader = (
            $VersionHeader -replace '\s+', ' '
        ).Trim()

        if (
            $NormalizedHeader -match
            'CCAM\s+version\s+(\d+(?:[.,]\d+)?)\s+applicable\s+(\d{2}/\d{2}/\d{4})'
        ) {
            $RawVersion = $Matches[1].Replace(",", ".")
            $ApplicabilityDate = $Matches[2]

            if ($RawVersion -match '^\d+$') {
                $DetectedVersion = "V$RawVersion.00"
            }
            else {
                $DetectedVersion = "V$RawVersion"
            }

            return [pscustomobject]@{
                version           = $DetectedVersion
                applicabilityDate = $ApplicabilityDate
                header            = $NormalizedHeader
            }
        }

        Write-Warning "CCAM metadata format was not recognized. Falling back to Claude."

        $MetadataSchemaPath = Join-Path `
            $ProjectRoot `
            "schemas\ccam-metadata-response.schema.json"

        if (-not (Test-Path -LiteralPath $MetadataSchemaPath)) {
            throw "CCAM metadata response schema not found: $MetadataSchemaPath"
        }

        $MetadataSchema = Get-Content `
            -LiteralPath $MetadataSchemaPath `
            -Raw |
            ConvertFrom-Json |
            ConvertTo-Json `
                -Depth 10 `
                -Compress
        $MetadataSchemaArgument = $MetadataSchema.Replace('"', '\"')
        $MetadataPrompt = @"
Extract the CCAM version and applicability date from this text:

$NormalizedHeader

Return:
- version in the format V84.00
- applicabilityDate in the format dd/MM/yyyy
"@

        $ClaudeRawResult = $MetadataPrompt |
            & claude `
                -p `
                --output-format json `
                --json-schema $MetadataSchemaArgument

        if ($LASTEXITCODE -ne 0) {
            throw "Claude metadata extraction failed with exit code $LASTEXITCODE."
        }

        try {
            $ClaudeResponse = $ClaudeRawResult |
                ConvertFrom-Json
        }
        catch {
            throw "Claude returned invalid metadata JSON: $($_.Exception.Message)"
        }

        if ($ClaudeResponse.is_error -eq $true) {
            throw "Claude returned a metadata extraction error."
        }

        if ($null -eq $ClaudeResponse.structured_output) {
            throw "Claude did not return structured CCAM metadata."
        }

        $Metadata = $ClaudeResponse.structured_output

        if (
            [string]::IsNullOrWhiteSpace(
                [string]$Metadata.version
            )
        ) {
            throw "Claude metadata response did not contain a version."
        }

        if (
            [string]::IsNullOrWhiteSpace(
                [string]$Metadata.applicabilityDate
            )
        ) {
            throw "Claude metadata response did not contain an applicability date."
        }

        return [pscustomobject]@{
            version           = [string]$Metadata.version
            applicabilityDate = [string]$Metadata.applicabilityDate
            header            = $NormalizedHeader
        }
    }
    finally {
        if ($null -ne $Workbook) {
            $Workbook.Close($false)
        }

        if ($null -ne $VersionRange) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $VersionRange
            )
        }

        if ($null -ne $Worksheet) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Worksheet
            )
        }

        if ($null -ne $Workbook) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Workbook
            )
        }

        if ($null -ne $Excel) {
            $Excel.Quit()

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Excel
            )
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

    $ConfigPath = Join-Path `
        $ProjectRoot `
        "config.json"

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Configuration file not found: $ConfigPath"
    }

    try {
        $Config = Get-Content `
            -LiteralPath $ConfigPath `
            -Raw |
            ConvertFrom-Json
    }
    catch {
        throw "config.json contains invalid JSON: $($_.Exception.Message)"
    }

    $StatePath = Join-Path `
        $ProjectRoot `
        $Config.paths.state

    $TransformSchemaPath = Join-Path `
        $ProjectRoot `
        $Config.paths.transformSchema

    $TransformSkillPath = Join-Path `
        $ProjectRoot `
        $Config.paths.transformSkill

    $ReviewReportScriptPath = Join-Path `
        $ProjectRoot `
        $Config.paths.reviewReportScript

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
        $State = Get-Content `
            -LiteralPath $StatePath `
            -Raw |
            ConvertFrom-Json
    }
    catch {
        throw "Workflow state file contains invalid JSON: $($_.Exception.Message)"
    }

    if ($null -eq $Config.connector) {
        throw "The connector section is missing from config.json."
    }

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$Config.connector.technicalContact
        )
    ) {
        throw "connector.technicalContact is missing from config.json."
    }

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$Config.connector.reviewer
        )
    ) {
        throw "connector.reviewer is missing from config.json."
    }

    if ($null -eq $Config.source) {
        throw "The source section is missing from config.json."
    }

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$Config.source.ccamExcelUrl
        )
    ) {
        throw "source.ccamExcelUrl is missing from config.json."
    }

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$Config.source.versionSheet
        )
    ) {
        throw "source.versionSheet is missing from config.json."
    }

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$Config.source.versionCell
        )
    ) {
        throw "source.versionCell is missing from config.json."
    }

    Write-Host ""
    Write-Host "Configuration loaded successfully."
    Write-Host "CCAM Excel URL: $($Config.source.ccamExcelUrl)"
    Write-Host "Version location: $($Config.source.versionSheet)!$($Config.source.versionCell)"
    Write-Host ""
    Write-Host "Workflow state loaded successfully."
    Write-Host "Status: $($State.status)"
    Write-Host "Detected version: $($State.detectedVersion)"

    if ($Validate) {
        Write-Host ""
        Write-Host "Validation mode enabled."
        Write-Host "Checking Claude CLI and project skill visibility..."

        $ClaudeCommand = Get-Command `
            claude `
            -ErrorAction SilentlyContinue

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

        $SkillVisibilityResult = $SkillVisibilityPrompt |
            & claude `
                -p `
                --output-format text

        if ($LASTEXITCODE -ne 0) {
            throw "Claude visibility test failed with exit code $LASTEXITCODE."
        }

        $SkillVisibilityResult = (
            $SkillVisibilityResult |
            Out-String
        ).Trim()

        Write-Host "Claude skill visibility result: $SkillVisibilityResult"

        if ($SkillVisibilityResult -ne "SKILL_FOUND") {
            throw "Claude could not confirm access to the project transformation skill."
        }

        Write-Host "Checking transformation response schema..."

        try {
            $null = Get-Content `
                -LiteralPath $TransformSchemaPath `
                -Raw |
                ConvertFrom-Json
        }
        catch {
            throw "Transformation response schema contains invalid JSON: $($_.Exception.Message)"
        }
        Write-Host "Checking Outlook review notification..."

        $ValidationReportPath = Join-Path `
            $ProjectRoot `
            "VALIDATION_ONLY_NO_REPORT.html"

        Send-CcamReviewEmail `
            -Recipient $Config.connector.reviewer `
            -Version "VALIDATION" `
            -ApplicabilityDate "N/A" `
            -ReportPath $ValidationReportPath

        Write-Host "Outlook review notification test completed."
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

$ReviewDirectory = [string]$Config.paths.reviewDirectory

if (-not [System.IO.Path]::IsPathRooted($ReviewDirectory)) {
    $ReviewDirectory = Join-Path `
        $ProjectRoot `
        $ReviewDirectory
}

        $SafeVersion = $Version -replace '[^A-Za-z0-9._-]', '_'

        $ReviewVersionDirectory = Join-Path `
            $ReviewDirectory `
            $SafeVersion

        if (-not (Test-Path -LiteralPath $ReviewVersionDirectory)) {
            New-Item `
                -ItemType Directory `
                -Path $ReviewVersionDirectory `
                -Force |
                Out-Null
        }

        $OutputPath = Join-Path `
            $ReviewVersionDirectory `
            "CCAM_${SafeVersion}_normalise.csv"

        try {
            $Schema = Get-Content `
                -LiteralPath $TransformSchemaPath `
                -Raw |
                ConvertFrom-Json |
                ConvertTo-Json `
                    -Depth 20 `
                    -Compress
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

        $ClaudeErrorPath = Join-Path `
            $ReviewVersionDirectory `
            "claude-error.txt"

        Remove-Item `
            -LiteralPath $ClaudeErrorPath `
            -Force `
            -ErrorAction SilentlyContinue

        $ClaudeRawResult = $Prompt |
            & claude `
                -p `
                --dangerously-skip-permissions `
                --output-format json `
                --json-schema $Schema `
                2> $ClaudeErrorPath

        $ClaudeExitCode = $LASTEXITCODE

        if ($ClaudeExitCode -ne 0) {
            $ClaudeError = ""

            if (Test-Path -LiteralPath $ClaudeErrorPath) {
                $ClaudeErrorContent = Get-Content `
                    -LiteralPath $ClaudeErrorPath `
                    -Raw

                if ($null -ne $ClaudeErrorContent) {
                    $ClaudeError = $ClaudeErrorContent.Trim()
                }
            }

            $ClaudeStdout = ""

            if ($null -ne $ClaudeRawResult) {
                $ClaudeStdout = (
                    $ClaudeRawResult |
                    Out-String
                ).Trim()
            }

            $ClaudeSummary = $null

            if (-not [string]::IsNullOrWhiteSpace($ClaudeStdout)) {
                try {
                    $ClaudeFailureResponse = $ClaudeStdout |
                        ConvertFrom-Json

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$ClaudeFailureResponse.result
                        )
                    ) {
                        $ClaudeSummary = [string]$ClaudeFailureResponse.result
                    }
                    elseif (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$ClaudeFailureResponse.api_error
                        )
                    ) {
                        $ClaudeSummary = [string]$ClaudeFailureResponse.api_error
                    }
                }
                catch {
                    # Use the raw stdout fallback below.
                }
            }

            if ([string]::IsNullOrWhiteSpace($ClaudeSummary)) {
                if (-not [string]::IsNullOrWhiteSpace($ClaudeError)) {
                    $ClaudeSummary = $ClaudeError
                }
                elseif (-not [string]::IsNullOrWhiteSpace($ClaudeStdout)) {
                    $ClaudeSummary = $ClaudeStdout
                }
                else {
                    $ClaudeSummary = "Claude returned no error details."
                }
            }

            throw @"
Claude transformation failed.

Exit code:
$ClaudeExitCode

Claude error:
$ClaudeSummary
"@
        }

        Remove-Item `
            -LiteralPath $ClaudeErrorPath `
            -Force `
            -ErrorAction SilentlyContinue

        try {
            $ClaudeResponse = $ClaudeRawResult |
                ConvertFrom-Json
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

        Write-Host "The daily CCAM source check will continue."
    }

    $InputDirectory = Join-Path `
        $ProjectRoot `
        $Config.paths.inputDirectory

    if (-not (Test-Path -LiteralPath $InputDirectory)) {
        New-Item `
            -ItemType Directory `
            -Path $InputDirectory `
            -Force |
            Out-Null
    }

    if ($Test) {
        Write-Host ""
        Write-Warning "TEST MODE ENABLED"
        Write-Host "CCAM download will be bypassed."

        if ($null -eq $Config.test) {
            throw "The test section is missing from config.json."
        }

        if (
            [string]::IsNullOrWhiteSpace(
                [string]$Config.test.excelFile
            )
        ) {
            throw "test.excelFile is missing from config.json."
        }

        $ExcelPath = Join-Path `
            $ProjectRoot `
            $Config.test.excelFile

        if (-not (Test-Path -LiteralPath $ExcelPath)) {
            throw "Test Excel file not found: $ExcelPath"
        }

        $WorkbookMetadata = Get-CcamWorkbookMetadata `
            -ExcelPath $ExcelPath `
            -WorksheetName $Config.source.versionSheet `
            -VersionCell $Config.source.versionCell

        $DetectedVersion = [string]$WorkbookMetadata.version
        $ApplicabilityDate = [string]$WorkbookMetadata.applicabilityDate

        $ExcelHash = (
            Get-FileHash `
                -LiteralPath $ExcelPath `
                -Algorithm SHA256
        ).Hash

        $ExcelSourceUrl = "LOCAL_TEST_FILE"

        Write-Host "Test CCAM version: $DetectedVersion"
        Write-Host "Applicable date: $ApplicabilityDate"
        Write-Host "Test Excel: $ExcelPath"
        Write-Host "Excel SHA-256: $ExcelHash"
    }
    else {
        Write-Host ""
        Write-Host "Downloading current CCAM Excel..."

        $ExcelPath = Join-Path `
            $InputDirectory `
            "ccam-current.xls"

        $CurlErrorPath = Join-Path `
            $InputDirectory `
            "ccam-curl-error.txt"

        Remove-Item `
            -LiteralPath $CurlErrorPath `
            -Force `
            -ErrorAction SilentlyContinue

        & curl.exe `
            -f `
            -sS `
            -L `
            --connect-timeout 10 `
            --max-time 120 `
            --retry 2 `
            --retry-delay 2 `
            $Config.source.ccamExcelUrl `
            -o $ExcelPath `
            2> $CurlErrorPath

        $CurlExitCode = $LASTEXITCODE

        if ($CurlExitCode -ne 0) {
            $CurlError = ""

            if (Test-Path -LiteralPath $CurlErrorPath) {
                $CurlErrorContent = Get-Content `
                    -LiteralPath $CurlErrorPath `
                    -Raw

                if ($null -ne $CurlErrorContent) {
                    $CurlError = $CurlErrorContent.Trim()
                }
            }

            if ([string]::IsNullOrWhiteSpace($CurlError)) {
                $CurlError = "curl returned no additional error information."
            }

            throw "Failed to download CCAM Excel. curl exit code: $CurlExitCode. Details: $CurlError"
        }

        Remove-Item `
            -LiteralPath $CurlErrorPath `
            -Force `
            -ErrorAction SilentlyContinue

        if (-not (Test-Path -LiteralPath $ExcelPath)) {
            throw "CCAM Excel file was not created."
        }

        Write-Host "CCAM Excel downloaded successfully."
        Write-Host "Saved to: $ExcelPath"
        Write-Host ""
        Write-Host "Calculating Excel hash..."

        $ExcelHash = (
            Get-FileHash `
                -LiteralPath $ExcelPath `
                -Algorithm SHA256
        ).Hash

        Write-Host "Current Excel hash: $ExcelHash"
        Write-Host "Previous Excel hash: $($State.artifacts.sourceHash)"

        if (
            $null -ne $State.artifacts.sourceHash -and
            $ExcelHash -eq $State.artifacts.sourceHash
        ) {
            Write-Host ""
            Write-Host "CCAM Excel has not changed."
            exit 0
        }

        Write-Host ""
        Write-Host "New CCAM Excel detected."

        $WorkbookMetadata = Get-CcamWorkbookMetadata `
            -ExcelPath $ExcelPath `
            -WorksheetName $Config.source.versionSheet `
            -VersionCell $Config.source.versionCell

        $DetectedVersion = [string]$WorkbookMetadata.version
        $ApplicabilityDate = [string]$WorkbookMetadata.applicabilityDate
        $ExcelSourceUrl = [string]$Config.source.ccamExcelUrl

        Write-Host "CCAM version: $DetectedVersion"
        Write-Host "Applicable date: $ApplicabilityDate"
        Write-Host "Excel SHA-256: $ExcelHash"
    }

    
        Write-Host ""
        Write-Host "No CCAM version has been processed yet."
        Write-Host "Creating workflow for $DetectedVersion"

        if ($null -eq $State.artifacts) {
            $State |
                Add-Member `
                    -MemberType NoteProperty `
                    -Name artifacts `
                    -Value ([pscustomobject]@{})
        }
        $PreviousVersion = [string]$State.detectedVersion
        $PreviousApplicabilityDate = [string]$State.applicabilityDate

        if ([string]::IsNullOrWhiteSpace($PreviousVersion)) {
            $PreviousVersion = "N/A"
        }

        if ([string]::IsNullOrWhiteSpace($PreviousApplicabilityDate)) {
            $PreviousApplicabilityDate = "N/A"
        }

        $State.detectedVersion = $DetectedVersion
        $State.applicabilityDate = $ApplicabilityDate
        $State.status = "DETECTED"
        $State.detectedAt = (
            Get-Date
        ).ToUniversalTime().ToString("o")

        $State.artifacts.sourceFile = $ExcelPath
        $State.artifacts.sourceUrl = $ExcelSourceUrl
        $State.artifacts.sourceHash = $ExcelHash

        $State |
            ConvertTo-Json -Depth 20 |
            Set-Content `
                -LiteralPath $StatePath `
                -Encoding UTF8

        Write-Host ""
        Write-Host "Workflow created."
        Write-Host "Status: DETECTED"
        Write-Host "Version: $DetectedVersion"

        $State.status = "TRANSFORMING"

        $State |
            ConvertTo-Json -Depth 20 |
            Set-Content `
                -LiteralPath $StatePath `
                -Encoding UTF8
        $WorkDirectory = Join-Path $ProjectRoot "work"

if (Test-Path -LiteralPath $WorkDirectory) {
    Write-Host ""
    Write-Host "Cleaning transformation work directory..."

    Get-ChildItem `
        -LiteralPath $WorkDirectory `
        -Force |
        Remove-Item `
            -Recurse `
            -Force `
            -ErrorAction Stop

    Write-Host "Work directory cleaned."
}
        $TransformResult = Invoke-CcamTransformation `
            -SourceExcelPath $ExcelPath `
            -Version $DetectedVersion

        if ($TransformResult.status -ne "PASS") {
            $State.status = "TRANSFORMATION_FAILED"

            $State |
                ConvertTo-Json -Depth 20 |
                Set-Content `
                    -LiteralPath $StatePath `
                    -Encoding UTF8

            throw "CCAM transformation failed validation."
        }

        $CandidateFile = [string]$TransformResult.outputPath

        if ([string]::IsNullOrWhiteSpace($CandidateFile)) {
            throw "Claude returned PASS without an outputPath."
        }

        if (-not (Test-Path -LiteralPath $CandidateFile)) {
            throw "Candidate output file does not exist: $CandidateFile"
        }

        $ReviewReport = & $ReviewReportScriptPath `
            -ProjectRoot $ProjectRoot `
            -TransformResult $TransformResult `
            -Version $DetectedVersion `
            -ApplicabilityDate $ApplicabilityDate `
            -PreviousVersion $PreviousVersion `
            -PreviousApplicabilityDate $PreviousApplicabilityDate `
            -CandidateFile $CandidateFile

        if ($null -eq $ReviewReport) {
            throw "Review report generator returned no result."
        }

        if (-not (Test-Path -LiteralPath $ReviewReport.reportPath)) {
            throw "Review report was not created: $($ReviewReport.reportPath)"
        }

        $State.status = "AWAITING_APPROVAL"
        $State.artifacts.candidateFile = $CandidateFile
        $State.review.reportFile = $ReviewReport.reportPath
        $State.transformation = $TransformResult
        $State.transformedAt = (
            Get-Date
        ).ToUniversalTime().ToString("o")

        $State |
            ConvertTo-Json -Depth 20 |
            Set-Content `
                -LiteralPath $StatePath `
                -Encoding UTF8
       
        Send-CcamReviewEmail `
            -Recipient $Config.connector.reviewer `
            -Version $DetectedVersion `
            -ApplicabilityDate $ApplicabilityDate `
            -ReportPath $ReviewReport.reportPath
        Write-Host ""
        Write-Host "Transformation completed successfully."
        Write-Host "Status: AWAITING_APPROVAL"
        Write-Host "Candidate file: $CandidateFile"
        Write-Host "Review report: $($ReviewReport.reportPath)"
        Write-Host "Human review items: $($ReviewReport.reviewCount)"
        Write-Host "Records: $($TransformResult.totalRecords)"

        exit 0

    Write-Host ""
    Write-Host "Detected CCAM version: $DetectedVersion"
}
catch {
    $FailureReason = $_.Exception.Message
    $FailurePosition = $_.InvocationInfo.PositionMessage

    Write-Host ""
    Write-Warning "CCAM automation failed: $FailureReason"
    Write-Warning "Location: $FailurePosition"

    $CurrentWorkflowStatus = "UNKNOWN"
    $CurrentDetectedVersion = "UNKNOWN"
    $TechnicalContact = $null

    if ($null -ne $State) {
        if (
            -not [string]::IsNullOrWhiteSpace(
                [string]$State.status
            )
        ) {
            $CurrentWorkflowStatus = [string]$State.status
        }

        if (
            -not [string]::IsNullOrWhiteSpace(
                [string]$State.detectedVersion
            )
        ) {
            $CurrentDetectedVersion = [string]$State.detectedVersion
        }
    }

    if (
        $null -ne $Config -and
        $null -ne $Config.connector -and
        -not [string]::IsNullOrWhiteSpace(
            [string]$Config.connector.technicalContact
        )
    ) {
        $TechnicalContact = [string]$Config.connector.technicalContact
    }

    $EmailReason = @"
$FailureReason

PowerShell location:
$FailurePosition
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
finally {
    Write-Host "=== CCAM END $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

    try {
        Stop-Transcript
    }
    catch {
        # Do not fail the CCAM workflow just because transcript cleanup failed.
    }
}