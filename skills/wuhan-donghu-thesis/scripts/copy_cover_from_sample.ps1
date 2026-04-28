param(
    [Parameter(Mandatory = $true)]
    [string]$SamplePath,

    [Parameter(Mandatory = $true)]
    [string]$DraftPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$ReplacementsJson = "",

    [int]$CoverParagraphCount = 47
)

$ErrorActionPreference = "Stop"

$wdGoToPage = 1
$wdGoToAbsolute = 1
$wdFormatXMLDocument = 16
$wdDoNotSaveChanges = 0

if (-not (Test-Path -LiteralPath $SamplePath)) {
    throw "SamplePath not found: $SamplePath"
}
if (-not (Test-Path -LiteralPath $DraftPath)) {
    throw "DraftPath not found: $DraftPath"
}
if ($ReplacementsJson -and -not (Test-Path -LiteralPath $ReplacementsJson)) {
    throw "ReplacementsJson not found: $ReplacementsJson"
}

$workPath = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    ("cover_work_{0}.docx" -f ([System.Guid]::NewGuid().ToString("N")))
)

Copy-Item -LiteralPath $DraftPath -Destination $workPath -Force

$replacements = @()
if ($ReplacementsJson) {
    $raw = Get-Content -LiteralPath $ReplacementsJson -Raw -Encoding UTF8
    $replacements = @($raw | ConvertFrom-Json)
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $sample = $word.Documents.Open($SamplePath, $false, $true)
    $doc = $word.Documents.Open($workPath, $false, $false)

    $samplePage3 = $sample.GoTo($wdGoToPage, $wdGoToAbsolute, 3)
    $sampleCoverRange = $sample.Range(0, $samplePage3.Start)

    $docPage3 = $doc.GoTo($wdGoToPage, $wdGoToAbsolute, 3)
    $docCoverRange = $doc.Range(0, $docPage3.Start)
    [void]$docCoverRange.Delete()

    $insertRange = $doc.Range(0, 0)
    $insertRange.FormattedText = $sampleCoverRange.FormattedText

    foreach ($pair in $replacements) {
        $findText = [string]$pair.find
        $replaceText = [string]$pair.replace
        if ($findText.Length -eq 0) {
            continue
        }
        $find = $doc.Content.Find
        [void]$find.ClearFormatting()
        [void]$find.Replacement.ClearFormatting()
        $find.Text = $findText
        $find.Replacement.Text = $replaceText
        $find.Forward = $true
        $find.Wrap = 1
        [void]$find.Execute($findText, $false, $false, $false, $false, $false, $true, 1, $false, $replaceText, 2)
    }

    $maxParagraph = [Math]::Min($CoverParagraphCount, [Math]::Min($sample.Paragraphs.Count, $doc.Paragraphs.Count))
    for ($i = 1; $i -le $maxParagraph; $i++) {
        $src = $sample.Paragraphs.Item($i).Format
        $dst = $doc.Paragraphs.Item($i).Format

        $dst.Alignment = $src.Alignment
        $dst.SpaceBefore = $src.SpaceBefore
        $dst.SpaceAfter = $src.SpaceAfter
        $dst.LineSpacingRule = $src.LineSpacingRule
        $dst.LineSpacing = $src.LineSpacing
        $dst.LeftIndent = $src.LeftIndent
        $dst.RightIndent = $src.RightIndent
        $dst.FirstLineIndent = $src.FirstLineIndent
        $dst.KeepTogether = $src.KeepTogether
        $dst.KeepWithNext = $src.KeepWithNext
        $dst.PageBreakBefore = $src.PageBreakBefore
        $dst.WidowControl = $src.WidowControl
    }

    $doc.SaveAs([ref]$workPath, [ref]$wdFormatXMLDocument)
    $doc.Close([ref]$wdDoNotSaveChanges)
    $sample.Close([ref]$wdDoNotSaveChanges)
}
finally {
    $word.Quit()
}

Copy-Item -LiteralPath $workPath -Destination $OutputPath -Force
Remove-Item -LiteralPath $workPath -Force -ErrorAction SilentlyContinue
Write-Output $OutputPath
