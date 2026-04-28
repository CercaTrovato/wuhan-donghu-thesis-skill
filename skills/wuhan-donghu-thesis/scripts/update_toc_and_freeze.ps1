param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [string]$PythonExe = "python",

    [int]$TocTabPosTwips = 8504,

    [int]$TocLineTwips = 460
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $DocxPath)) {
    throw "DocxPath not found: $DocxPath"
}

$postprocessScript = Join-Path $PSScriptRoot "postprocess_toc_xml.py"
if (-not (Test-Path -LiteralPath $postprocessScript)) {
    throw "postprocess_toc_xml.py not found next to this script"
}

$wdDoNotSaveChanges = 0

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($DocxPath, $false, $false)
    $doc.Repaginate()
    [void]$doc.Fields.Update()
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        $doc.TablesOfContents.Item($i).Update()
    }
    $doc.Repaginate()
    $tocCount = $doc.TablesOfContents.Count
    $pageCount = $doc.ComputeStatistics(2)
    $doc.Save()
    $doc.Close([ref]$wdDoNotSaveChanges)
}
finally {
    $word.Quit()
}

& $PythonExe $postprocessScript $DocxPath --tab-pos $TocTabPosTwips --line $TocLineTwips

Write-Output "TOC_COUNT=$tocCount"
Write-Output "PAGE_COUNT=$pageCount"
Write-Output $DocxPath
