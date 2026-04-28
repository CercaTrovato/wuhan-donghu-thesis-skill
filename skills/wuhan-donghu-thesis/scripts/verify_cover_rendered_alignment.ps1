param(
    [Parameter(Mandatory = $true)]
    [string]$SamplePath,

    [Parameter(Mandatory = $true)]
    [string]$CandidatePath,

    [double]$TolerancePt = 2.5
)

$ErrorActionPreference = "Stop"

$wdHorizontalPositionRelativeToPage = 5
$wdVerticalPositionRelativeToPage = 6
$wdDoNotSaveChanges = 0

if (-not (Test-Path -LiteralPath $SamplePath)) {
    throw "SamplePath not found: $SamplePath"
}
if (-not (Test-Path -LiteralPath $CandidatePath)) {
    throw "CandidatePath not found: $CandidatePath"
}

function Get-ParagraphPlainText {
    param($Paragraph)
    return ([string]$Paragraph.Range.Text).TrimEnd([char]13, [char]7)
}

function Get-InsertionPointGeometry {
    param(
        [Parameter(Mandatory = $true)]
        $Document,

        [Parameter(Mandatory = $true)]
        [int]$Position
    )

    $range = $Document.Range($Position, $Position)
    $range.Select()

    return [pscustomobject]@{
        X = [double]$Document.Application.Selection.Information($wdHorizontalPositionRelativeToPage)
        Y = [double]$Document.Application.Selection.Information($wdVerticalPositionRelativeToPage)
    }
}

function Get-FieldGeometry {
    param(
        [Parameter(Mandatory = $true)]
        $Document,

        [Parameter(Mandatory = $true)]
        [int]$ParagraphIndex
    )

    $paragraph = $Document.Paragraphs.Item($ParagraphIndex)
    $text = Get-ParagraphPlainText -Paragraph $paragraph
    $colonIndex = $text.IndexOf([char]0xFF1A)
    if ($colonIndex -lt 0) {
        throw "Paragraph $ParagraphIndex does not contain full-width colon."
    }

    $lineStart = [int]$paragraph.Range.Start
    $suffixStart = [int]($paragraph.Range.Start + $colonIndex + 1)
    $suffixEnd = [int]($paragraph.Range.End - 1)

    $lineStartGeometry = Get-InsertionPointGeometry -Document $Document -Position $lineStart
    $suffixStartGeometry = Get-InsertionPointGeometry -Document $Document -Position $suffixStart
    $suffixEndGeometry = Get-InsertionPointGeometry -Document $Document -Position $suffixEnd

    return [pscustomobject]@{
        ParagraphIndex = $ParagraphIndex
        Text = $text
        LineStartX = $lineStartGeometry.X
        SuffixStartX = $suffixStartGeometry.X
        SuffixEndX = $suffixEndGeometry.X
        BaselineY = $suffixStartGeometry.Y
    }
}

function Compare-FieldGeometry {
    param(
        [System.Collections.ArrayList]$Failures,

        [Parameter(Mandatory = $true)]
        $SampleGeometry,

        [Parameter(Mandatory = $true)]
        $CandidateGeometry,

        [Parameter(Mandatory = $true)]
        [double]$TolerancePt
    )

    $checks = @(
        [pscustomobject]@{ Name = "lineStartX"; Sample = $SampleGeometry.LineStartX; Candidate = $CandidateGeometry.LineStartX },
        [pscustomobject]@{ Name = "suffixStartX"; Sample = $SampleGeometry.SuffixStartX; Candidate = $CandidateGeometry.SuffixStartX },
        [pscustomobject]@{ Name = "suffixEndX"; Sample = $SampleGeometry.SuffixEndX; Candidate = $CandidateGeometry.SuffixEndX }
    )

    foreach ($check in $checks) {
        $delta = [Math]::Abs($check.Candidate - $check.Sample)
        if ($delta -gt $TolerancePt) {
            [void]$Failures.Add((
                "P{0} {1}: candidate {2:N2}pt differs from sample {3:N2}pt by {4:N2}pt" -f
                    $SampleGeometry.ParagraphIndex,
                    $check.Name,
                    $check.Candidate,
                    $check.Sample,
                    $delta
            ))
        }
    }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$sample = $null
$candidate = $null
$failures = [System.Collections.ArrayList]::new()

try {
    $sample = $word.Documents.Open($SamplePath, $false, $true)
    $candidate = $word.Documents.Open($CandidatePath, $false, $true)

    $sample.Activate()
    $sample.Repaginate()
    $candidate.Activate()
    $candidate.Repaginate()

    foreach ($paragraphIndex in @(2, 3, 15, 16, 17, 18)) {
        $sample.Activate()
        $sampleGeometry = Get-FieldGeometry -Document $sample -ParagraphIndex $paragraphIndex

        $candidate.Activate()
        $candidateGeometry = Get-FieldGeometry -Document $candidate -ParagraphIndex $paragraphIndex

        Compare-FieldGeometry `
            -Failures $failures `
            -SampleGeometry $sampleGeometry `
            -CandidateGeometry $candidateGeometry `
            -TolerancePt $TolerancePt
    }
}
finally {
    if ($null -ne $candidate) {
        $candidate.Close([ref]$wdDoNotSaveChanges)
    }
    if ($null -ne $sample) {
        $sample.Close([ref]$wdDoNotSaveChanges)
    }
    $word.Quit()
}

if ($failures.Count -gt 0) {
    Write-Output "COVER_RENDERED_ALIGNMENT=FAIL"
    foreach ($failure in $failures) {
        Write-Output "- $failure"
    }
    exit 1
}

Write-Output "COVER_RENDERED_ALIGNMENT=PASS"
