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
    [string]$CandidateHash
)

$ErrorActionPreference = "Stop"

function ConvertTo-HtmlText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

$TemplatePath = Join-Path $ProjectRoot "Context\ccam-review-report.template.html"
$ReviewDirectory = Join-Path $ProjectRoot "review"

if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "CCAM review template not found: $TemplatePath"
}

if (-not (Test-Path -LiteralPath $CandidateFile)) {
    throw "Candidate file not found: $CandidateFile"
}

if (-not (Test-Path -LiteralPath $ReviewDirectory)) {
    New-Item -ItemType Directory -Path $ReviewDirectory | Out-Null
}

$SafeVersion = $Version -replace '[^A-Za-z0-9._-]', '_'
$ReportPath = Join-Path $ReviewDirectory "CCAM_${SafeVersion}_review.html"
$Html = Get-Content -LiteralPath $TemplatePath -Raw

$GeneratedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$RunId = (Get-Date).ToUniversalTime().ToString("yyyyMMdd-HHmmss")
$CurrentFileName = [System.IO.Path]::GetFileName($CandidateFile)

# Previous-version comparison will populate these later.
$PreviousVersion = "N/A"
$PreviousDate = "N/A"
$CurrentDate = "N/A"
$AddedCount = 0
$RemovedCount = 0
$ChangedCount = 0

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

$AddedRows = '<tr class="empty"><td colspan="3">Previous-version comparison not available yet.</td></tr>'
$RemovedRows = '<tr class="empty"><td colspan="3">Previous-version comparison not available yet.</td></tr>'
$ChangedRows = '<tr class="empty"><td colspan="5">Previous-version comparison not available yet.</td></tr>'

$Values = [ordered]@{
    '{{REPORT_GENERATED_AT}}' = ConvertTo-HtmlText $GeneratedAt
    '{{RUN_ID}}' = ConvertTo-HtmlText $RunId
    '{{PREVIOUS_VERSION}}' = ConvertTo-HtmlText $PreviousVersion
    '{{PREVIOUS_DATE}}' = ConvertTo-HtmlText $PreviousDate
    '{{CURRENT_VERSION}}' = ConvertTo-HtmlText $Version
    '{{CURRENT_DATE}}' = ConvertTo-HtmlText $CurrentDate
    '{{CURRENT_FILE_NAME}}' = ConvertTo-HtmlText $CurrentFileName
    '{{CURRENT_FILE_SHA256}}' = ConvertTo-HtmlText $CandidateHash
    '{{ADDED_COUNT}}' = [string]$AddedCount
    '{{REMOVED_COUNT}}' = [string]$RemovedCount
    '{{CHANGED_COUNT}}' = [string]$ChangedCount
    '{{REVIEW_COUNT}}' = [string]$ReviewCount
    '{{REVIEW_ROWS}}' = $ReviewRowsHtml
    '{{EMPTY_REVIEW}}' = $EmptyReviewHtml
    '{{ADDED_ROWS}}' = $AddedRows
    '{{REMOVED_ROWS}}' = $RemovedRows
    '{{CHANGED_ROWS}}' = $ChangedRows
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
