param(
    [Parameter(Mandatory = $true)]
    [string]$SamplePath,

    [Parameter(Mandatory = $true)]
    [string]$CandidatePath,

    [int[]]$Paragraphs = @(2, 3, 9, 11, 12, 15, 16, 17, 18, 22, 28, 32, 38, 39, 40, 46)
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $SamplePath)) {
    throw "SamplePath not found: $SamplePath"
}
if (-not (Test-Path -LiteralPath $CandidatePath)) {
    throw "CandidatePath not found: $CandidatePath"
}

$wdDoNotSaveChanges = 0
$wdUndefined = 9999999

function Is-ConcreteFontValue {
    param($Value)
    if ($null -eq $Value) {
        return $false
    }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }
    if ($text.StartsWith("+")) {
        return $false
    }
    return $true
}

function Add-Failure {
    param(
        [System.Collections.Generic.List[string]]$Failures,
        [int]$Paragraph,
        [string]$Property,
        $Expected,
        $Actual
    )
    $Failures.Add(("P{0} {1}: expected [{2}], actual [{3}]" -f $Paragraph, $Property, $Expected, $Actual))
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$failures = [System.Collections.Generic.List[string]]::new()

try {
    $sample = $word.Documents.Open($SamplePath, $false, $true)
    $candidate = $word.Documents.Open($CandidatePath, $false, $true)

    foreach ($i in $Paragraphs) {
        if ($i -gt $sample.Paragraphs.Count -or $i -gt $candidate.Paragraphs.Count) {
            $failures.Add(("P{0}: paragraph missing in sample or candidate" -f $i))
            continue
        }

        $srcFont = $sample.Paragraphs.Item($i).Range.Font
        $dstFont = $candidate.Paragraphs.Item($i).Range.Font

        foreach ($property in @("Name", "NameAscii", "NameFarEast", "NameOther")) {
            $expected = $srcFont.$property
            $actual = $dstFont.$property
            if ((Is-ConcreteFontValue $expected) -and ([string]$actual -ne [string]$expected)) {
                Add-Failure -Failures $failures -Paragraph $i -Property $property -Expected $expected -Actual $actual
            }
        }

        if ($srcFont.Size -gt 0 -and [double]$dstFont.Size -ne [double]$srcFont.Size) {
            Add-Failure -Failures $failures -Paragraph $i -Property "Size" -Expected $srcFont.Size -Actual $dstFont.Size
        }
        if ($srcFont.Bold -ne $wdUndefined -and [int]$dstFont.Bold -ne [int]$srcFont.Bold) {
            Add-Failure -Failures $failures -Paragraph $i -Property "Bold" -Expected $srcFont.Bold -Actual $dstFont.Bold
        }
    }

    $candidate.Close([ref]$wdDoNotSaveChanges)
    $sample.Close([ref]$wdDoNotSaveChanges)
}
finally {
    $word.Quit()
}

if ($failures.Count -gt 0) {
    Write-Output "COVER_FONT_COMPARISON=FAIL"
    foreach ($failure in $failures) {
        Write-Output "- $failure"
    }
    exit 1
}

Write-Output "COVER_FONT_COMPARISON=PASS"
