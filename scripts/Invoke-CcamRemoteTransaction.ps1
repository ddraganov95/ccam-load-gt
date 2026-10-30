param(
    [Parameter(Mandatory = $true)]
    [string]$BatchUrl = "http://issapp014.csccha.fr:13161",

    [Parameter(Mandatory)]
    [string]$CcamFile,

    [switch]$Async
)

$ErrorActionPreference = "Stop"

if (-not [Uri]::IsWellFormedUriString($BatchUrl, [UriKind]::Absolute)) {
    throw "BatchUrl must be a complete URL, for example http://server:12345"
}

$uri = [Uri]$BatchUrl

Write-Host ""
Write-Host "GraphTalk CCAM transaction test"
Write-Host "Endpoint : $BatchUrl"
Write-Host "File     : $CcamFile"
Write-Host "Mode     : $(if ($Async) { 'asynchronous' } else { 'synchronous' })"
Write-Host ""

# Escape only values inserted into XML.
$escapedFile = [System.Security.SecurityElement]::Escape($CcamFile)
$runDate = Get-Date -Format "dd-MM-yyyy"

$soap = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <soap:Body>
    <GraphTalkRunTrx xmlns="http://www.csc.com/graphtalk/aia/"
                     xmlns:gtx="graphtalk.aia.extension">
      <Environment>
        <EnvDateOverload>"$runDate"</EnvDateOverload>
      </Environment>
      <EnvironmentVariables/>
      <RequestVariables/>
      <TrxPackage>rfrnc</TrxPackage>
      <TrxTheme>trx_upgrade</TrxTheme>
      <TrxName>load_rfrnc_param_vals</TrxName>
      <TrxType>batch</TrxType>
      <TrxTrace>no_tracing</TrxTrace>
      <TrxMode>save</TrxMode>
      <TrxParams>[file_name="$escapedFile",type="ccam_ro"]</TrxParams>
      <TrxDirectAccessClasses>[]</TrxDirectAccessClasses>
    </GraphTalkRunTrx>
  </soap:Body>
</soap:Envelope>
"@

$headers = @{
    "GtMonitor-Async" = if ($Async) { "true" } else { "false" }
}

try {
    $response = Invoke-WebRequest `
        -Uri $BatchUrl `
        -Method Post `
        -Headers $headers `
        -ContentType "text/xml; charset=utf-8" `
        -Body $soap `
        -UseBasicParsing

    Write-Host "HTTP status: $($response.StatusCode)"

    if ($Async) {
        if ($response.StatusCode -ne 202) {
            throw "Expected HTTP 202 for asynchronous execution, received $($response.StatusCode)."
        }

        $location = $response.Headers["Location"]

        if ([string]::IsNullOrWhiteSpace($location)) {
            throw "GT Monitor returned 202 but no Location header."
        }

        $resultUrl = [Uri]::new($uri, $location).AbsoluteUri

        Write-Host "Accepted asynchronously."
        Write-Host "Result URL: $resultUrl"
        Write-Host ""

        do {
            Start-Sleep -Seconds 2

            try {
                $result = Invoke-WebRequest `
                    -Uri $resultUrl `
                    -Method Get `
                    -Headers @{ "GtMonitor-Timeout" = "28" } `
                    -UseBasicParsing

                $status = $result.StatusCode
            }
            catch {
                if ($_.Exception.Response) {
                    $status = [int]$_.Exception.Response.StatusCode
                }
                else {
                    throw
                }
            }

            Write-Host "Poll status: $status"
        }
        while ($status -eq 202)

        if ($status -ne 200) {
            throw "Unexpected result status: $status"
        }

        $responseContent = $result.Content
    }
    else {
        if ($response.StatusCode -ne 200) {
            throw "Expected HTTP 200 for synchronous execution, received $($response.StatusCode)."
        }

        $responseContent = $response.Content
    }

    [xml]$responseXml = $responseContent

    $namespaceManager = [System.Xml.XmlNamespaceManager]::new(
        $responseXml.NameTable
    )
    $namespaceManager.AddNamespace(
        "aia",
        "http://www.csc.com/graphtalk/aia/"
    )

    $returnCodeNode = $responseXml.SelectSingleNode(
        "//*[local-name()='ReturnCode']"
    )
    $executionIdNode = $responseXml.SelectSingleNode(
        "//*[local-name()='ExecutionId']"
    )
    $durationNode = $responseXml.SelectSingleNode(
        "//*[local-name()='ExecutionDuration']"
    )
    $stepsInErrorNode = $responseXml.SelectSingleNode(
        "//*[local-name()='StepsInError']"
    )
    $executionStatusNode = $responseXml.SelectSingleNode(
        "//*[local-name()='ExecutionStatus']"
    )

    $returnCode = if ($returnCodeNode) {
        $returnCodeNode.InnerText
    }
    else {
        $null
    }

    $executionId = if ($executionIdNode) {
        $executionIdNode.InnerText
    }
    else {
        $null
    }

    $duration = if ($durationNode) {
        $durationNode.InnerText
    }
    else {
        $null
    }

    $stepsInError = if ($stepsInErrorNode) {
        $stepsInErrorNode.InnerText
    }
    else {
        $null
    }

    $executionStatus = if ($executionStatusNode) {
        $executionStatusNode.InnerText
    }
    else {
        $null
    }

    Write-Host ""
    Write-Host "GraphTalk response"
    Write-Host "Return code     : $returnCode"
    Write-Host "Execution ID    : $ecution status: $executionStatus"
    Write-Host "Execution status: $executionStatus"
    Write-Host "Steps in error  : $stepsInError"
    Write-Host "Duration        : $duration"
    Write-Host ""

    $successful = (
        $returnCode -eq "0" -and
        ($null -eq $stepsInError -or $stepsInError -eq "0") -and
        (
            $null -eq $executionStatus -or
            $executionStatus -eq "completed"
        )
    )

    if (-not $successful) {
        Write-Host "Full response:"
        Write-Host $responseContent
        throw "GraphTalk transaction did not report success."
    }

    Write-Host "SUCCESS: GraphTalk completed load_rfrnc_param_vals."
    [pscustomobject]@{
    success           = $true
    asynchronous      = [bool]$Async
    endpoint          = $BatchUrl
    fileName          = $CcamFile
    returnCode        = $returnCode
    executionId       = $executionId
    executionStatus   = $executionStatus
    stepsInError      = $stepsInError
    executionDuration = $duration
}
}
catch {
    Write-Host ""
    Write-Host "FAILED"
    Write-Host $_.Exception.Message

    if ($_.ErrorDetails.Message) {
        Write-Host ""
        Write-Host "Server response:"
        Write-Host $_.ErrorDetails.Message
    }

    exit 1
}