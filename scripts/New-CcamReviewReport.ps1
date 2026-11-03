param(
    [Parameter(Mandatory)]
    [string]$ProjectRoot,

    [Parameter(Mandatory)]
    [object]$TransformResult,

    [Parameter(Mandatory)]
    [string]$Version,

    [Parameter(Mandatory)]
    [string]$CandidateFile,

    [Parameter(Mandatory)]
    [string]$ApplicabilityDate,
    
    [Parameter(Mandatory)]
    [string]$PreviousVersion,

    [Parameter(Mandatory)]
    [string]$PreviousApplicabilityDate
)

$ErrorActionPreference = "Stop"

function ConvertTo-HtmlText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

$TemplatePath = Join-Path `
    $ProjectRoot `
    "Context\ccam-review-report.template.html"

if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "CCAM review template not found: $TemplatePath"
}

if (-not (Test-Path -LiteralPath $CandidateFile)) {
    throw "Candidate file not found: $CandidateFile"
}

$CandidateDirectory = Split-Path `
    -Parent `
    $CandidateFile
$ConfigPath = Join-Path $ProjectRoot "config.json"
$Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$GraphTalkEndpoint = [string]$Config.graphTalk.endpoint
if ([string]::IsNullOrWhiteSpace($CandidateDirectory)) {
    throw "Unable to determine candidate directory from: $CandidateFile"
}

if (-not (Test-Path -LiteralPath $CandidateDirectory)) {
    throw "Candidate directory not found: $CandidateDirectory"
}

$SafeVersion = $Version -replace '[^A-Za-z0-9._-]', '_'

$ReportPath = Join-Path `
    $CandidateDirectory `
    "CCAM_${SafeVersion}_review.html"
$LogoSourcePath = Join-Path `
    $ProjectRoot `
    "Context\logo.jpg"

$LogoTargetPath = Join-Path `
    $CandidateDirectory `
    "logo.jpg"

if (-not (Test-Path -LiteralPath $LogoSourcePath)) {
    throw "DXC logo not found: $LogoSourcePath"
}

Copy-Item `
    -LiteralPath $LogoSourcePath `
    -Destination $LogoTargetPath `
    -Force
$Html = Get-Content `
    -LiteralPath $TemplatePath `
    -Raw

$GeneratedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$RunId = (Get-Date).ToUniversalTime().ToString("yyyyMMdd-HHmmss")
$CurrentFileName = [System.IO.Path]::GetFileName($CandidateFile)
$CandidateFileUri = ([System.Uri]$CandidateFile).AbsoluteUri
$CandidateFileJs = $CandidateFile.Replace('\', '\\')
$PreviousDate = $PreviousApplicabilityDate
$CurrentDate = $ApplicabilityDate

$ReviewRows = @()

foreach ($Item in @($TransformResult.humanReviewItems)) {
    if ($null -eq $Item) {
        continue
    }

    $ReviewReason = switch -Regex ([string]$Item.category) {
        'multiline'       { 'Ambiguous label'; break }
        'nmbr|tooth|tail' { 'Ambiguous Nmbr_tooth'; break }
        default           { [string]$Item.category }
    }

    $ReviewRows += @"
<tr>
  <td class="code">$(ConvertTo-HtmlText $Item.code)</td>
  <td>$(ConvertTo-HtmlText $ReviewReason)</td>
  <td class="before">$(ConvertTo-HtmlText $Item.sourceText)</td>
  <td class="after">$(ConvertTo-HtmlText $Item.decision)</td>
  <td>$(ConvertTo-HtmlText $Item.reason)</td>
</tr>
"@
}

$ReviewCount = $ReviewRows.Count

if ($ReviewCount -eq 0) {
    $ReviewRowsHtml = ""
    $EmptyReviewHtml = '<tr class="empty"><td colspan="5">No items require human review.</td></tr>'
}
else {
    $ReviewRowsHtml = $ReviewRows -join "`r`n"
    $EmptyReviewHtml = ""
}

$Values = [ordered]@{
    '{{REPORT_GENERATED_AT}}' = ConvertTo-HtmlText $GeneratedAt
    '{{RUN_ID}}' = ConvertTo-HtmlText $RunId
    '{{PREVIOUS_VERSION}}' = ConvertTo-HtmlText $PreviousVersion
    '{{PREVIOUS_DATE}}' = ConvertTo-HtmlText $PreviousDate
    '{{CURRENT_VERSION}}' = ConvertTo-HtmlText $Version
    '{{CURRENT_DATE}}' = ConvertTo-HtmlText $CurrentDate
    '{{CURRENT_FILE_NAME}}' = ConvertTo-HtmlText $CurrentFileName
    '{{CANDIDATE_FILE_URI}}' = ConvertTo-HtmlText $CandidateFileUri
    '{{CANDIDATE_FILE_PATH}}' = ConvertTo-HtmlText $CandidateFileJs
    '{{GRAPHTALK_ENDPOINT}}' = ConvertTo-HtmlText $GraphTalkEndpoint
    '{{TOTAL_RECORDS}}' = [string]$TransformResult.totalRecords
    '{{UNIQUE_CODES}}' = [string]$TransformResult.uniqueCodes
    '{{MULTI_ACTIVITY_CODES}}' = [string]$TransformResult.multiActivityCodes
    '{{PHASED_RECORDS}}' = [string]$TransformResult.phasedRecords
    '{{POPULATED_NMBR_TOOTH}}' = [string]$TransformResult.populatedNmbrTooth
    '{{EMPTY_TARIFF_RECORDS}}' = [string]$TransformResult.emptyTariffRecords
    '{{REVIEW_COUNT}}' = [string]$ReviewCount
    '{{REVIEW_ROWS}}' = $ReviewRowsHtml
    '{{EMPTY_REVIEW}}' = $EmptyReviewHtml
}

foreach ($Token in $Values.Keys) {
    $Html = $Html.Replace($Token, [string]$Values[$Token])
}

Set-Content -LiteralPath $ReportPath -Value $Html -Encoding UTF8

$ReportHash = (Get-FileHash -LiteralPath $ReportPath -Algorithm SHA256).Hash

[pscustomobject]@{
    reportPath  = $ReportPath
    reportHash  = $ReportHash
    reviewCount = $ReviewCount
    runId       = $RunId
}
