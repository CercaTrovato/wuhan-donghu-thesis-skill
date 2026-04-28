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
$wdUndefined = 9999999

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
    $parsedReplacements = $raw | ConvertFrom-Json
    if ($null -ne $parsedReplacements) {
        $replacements = $parsedReplacements
    }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

function Set-FontPropertyIfConcrete {
    param(
        [Parameter(Mandatory = $true)]
        $Font,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        $Value
    )

    if ($null -eq $Value) {
        return
    }

    $stringValue = [string]$Value
    if ([string]::IsNullOrWhiteSpace($stringValue)) {
        return
    }
    if ($stringValue.StartsWith("+")) {
        return
    }

    $Font.$PropertyName = $Value
}

function Copy-CoverParagraphFont {
    param(
        [Parameter(Mandatory = $true)]
        $SourceParagraph,

        [Parameter(Mandatory = $true)]
        $TargetParagraph
    )

    $srcFont = $SourceParagraph.Range.Font
    $dstRange = $TargetParagraph.Range
    $dstFont = $dstRange.Font

    Set-FontPropertyIfConcrete -Font $dstFont -PropertyName "Name" -Value $srcFont.Name
    Set-FontPropertyIfConcrete -Font $dstFont -PropertyName "NameAscii" -Value $srcFont.NameAscii
    Set-FontPropertyIfConcrete -Font $dstFont -PropertyName "NameFarEast" -Value $srcFont.NameFarEast
    Set-FontPropertyIfConcrete -Font $dstFont -PropertyName "NameOther" -Value $srcFont.NameOther

    if ($srcFont.Size -gt 0) {
        $dstFont.Size = $srcFont.Size
    }
    if ($srcFont.Bold -ne $wdUndefined) {
        $dstFont.Bold = $srcFont.Bold
    }
    if ($srcFont.Italic -ne $wdUndefined) {
        $dstFont.Italic = $srcFont.Italic
    }

    # Cover pages are fixed template pages; spell-check marks should not affect the delivered visual.
    $dstRange.NoProofing = $true
}

function Get-ParagraphPlainText {
    param($Paragraph)
    return ([string]$Paragraph.Range.Text).TrimEnd([char]13, [char]7)
}

function Get-FieldLabel {
    param($Paragraph)

    $text = Get-ParagraphPlainText -Paragraph $Paragraph
    $colon = [char]0xFF1A
    $index = $text.IndexOf($colon)
    if ($index -lt 0) {
        return ""
    }
    return $text.Substring(0, $index + 1)
}

function Get-SuffixValue {
    param(
        $Paragraph
    )

    $label = Get-FieldLabel -Paragraph $Paragraph
    if ([string]::IsNullOrEmpty($label)) {
        return ""
    }
    $text = Get-ParagraphPlainText -Paragraph $Paragraph
    $index = $text.IndexOf($label)
    if ($index -lt 0) {
        return ""
    }
    return $text.Substring($index + $label.Length).Trim()
}

function Set-UnderlinedFieldSlotText {
    param(
        $Document,
        [int]$ParagraphIndex,
        [string]$SlotText
    )

    $paragraph = $Document.Paragraphs.Item($ParagraphIndex)
    $label = Get-FieldLabel -Paragraph $paragraph
    if ([string]::IsNullOrEmpty($label)) {
        return
    }
    $text = Get-ParagraphPlainText -Paragraph $paragraph
    $labelIndex = $text.IndexOf($label)
    if ($labelIndex -lt 0) {
        return
    }

    $suffixStart = $paragraph.Range.Start + $labelIndex + $label.Length
    $suffixEnd = $paragraph.Range.End - 1
    if ($suffixEnd -lt $suffixStart) {
        $suffixEnd = $suffixStart
    }

    $suffixRange = $Document.Range($suffixStart, $suffixEnd)
    $suffixRange.Text = $SlotText

    $paragraph = $Document.Paragraphs.Item($ParagraphIndex)
    $label = Get-FieldLabel -Paragraph $paragraph
    if ([string]::IsNullOrEmpty($label)) {
        return
    }
    $text = Get-ParagraphPlainText -Paragraph $paragraph
    $labelIndex = $text.IndexOf($label)
    if ($labelIndex -lt 0) {
        return
    }

    $suffixStart = $paragraph.Range.Start + $labelIndex + $label.Length
    $suffixEnd = $paragraph.Range.End - 1
    if ($suffixEnd -lt $suffixStart) {
        $suffixEnd = $suffixStart
    }

    $labelRange = $Document.Range($paragraph.Range.Start, $suffixStart)
    $suffixRange = $Document.Range($suffixStart, $suffixEnd)
    $labelRange.Font.Underline = 0
    $suffixRange.Font.Underline = 1
    $paragraph.Range.NoProofing = $true
}

function Get-InsertionPointGeometry {
    param(
        $Document,
        [int]$Position
    )

    $range = $Document.Range($Position, $Position)
    $range.Select()

    return [pscustomobject]@{
        X = [double]$Document.Application.Selection.Information(5)
        Y = [double]$Document.Application.Selection.Information(6)
    }
}

function Get-FieldGeometry {
    param(
        $Document,
        [int]$ParagraphIndex,
        [string]$Value = ""
    )

    $paragraph = $Document.Paragraphs.Item($ParagraphIndex)
    $text = Get-ParagraphPlainText -Paragraph $paragraph
    $label = Get-FieldLabel -Paragraph $paragraph
    if ([string]::IsNullOrEmpty($label)) {
        return $null
    }

    $labelIndex = $text.IndexOf($label)
    if ($labelIndex -lt 0) {
        return $null
    }

    $suffixText = $text.Substring($labelIndex + $label.Length)
    $lineStart = [int]$paragraph.Range.Start
    $suffixStart = [int]($paragraph.Range.Start + $labelIndex + $label.Length)
    $suffixEnd = [int]($paragraph.Range.End - 1)

    $lineStartGeometry = Get-InsertionPointGeometry -Document $Document -Position $lineStart
    $suffixStartGeometry = Get-InsertionPointGeometry -Document $Document -Position $suffixStart
    $suffixEndGeometry = Get-InsertionPointGeometry -Document $Document -Position $suffixEnd

    $valueStartX = $null
    $valueEndX = $null
    if (-not [string]::IsNullOrEmpty($Value)) {
        $valueIndex = $suffixText.IndexOf($Value)
        if ($valueIndex -ge 0) {
            $valueStart = [int]($suffixStart + $valueIndex)
            $valueEnd = [int]($valueStart + $Value.Length)
            $valueStartGeometry = Get-InsertionPointGeometry -Document $Document -Position $valueStart
            $valueEndGeometry = Get-InsertionPointGeometry -Document $Document -Position $valueEnd
            $valueStartX = $valueStartGeometry.X
            $valueEndX = $valueEndGeometry.X
        }
    }

    return [pscustomobject]@{
        ParagraphIndex = $ParagraphIndex
        LineStartX = $lineStartGeometry.X
        SuffixStartX = $suffixStartGeometry.X
        SuffixEndX = $suffixEndGeometry.X
        ValueStartX = $valueStartX
        ValueEndX = $valueEndX
    }
}

function Get-GeometryEdgeScore {
    param(
        $Geometry,
        $TargetGeometry
    )

    if ($null -eq $Geometry -or $null -eq $TargetGeometry) {
        return [double]::PositiveInfinity
    }

    return (
        [Math]::Abs($Geometry.LineStartX - $TargetGeometry.LineStartX) +
        [Math]::Abs($Geometry.SuffixStartX - $TargetGeometry.SuffixStartX) +
        [Math]::Abs($Geometry.SuffixEndX - $TargetGeometry.SuffixEndX)
    )
}

function Get-ValueCenterScore {
    param(
        $Geometry,
        $TargetGeometry,
        [string]$Value
    )

    if ([string]::IsNullOrEmpty($Value)) {
        return 0.0
    }
    if ($null -eq $Geometry.ValueStartX -or $null -eq $Geometry.ValueEndX) {
        return 10000.0
    }

    $targetCenter = ($TargetGeometry.SuffixStartX + $TargetGeometry.SuffixEndX) / 2.0
    $valueCenter = ($Geometry.ValueStartX + $Geometry.ValueEndX) / 2.0
    return [Math]::Abs($valueCenter - $targetCenter)
}

function Set-MeasuredUnderlinedField {
    param(
        $Document,
        [int]$ParagraphIndex,
        $TargetGeometry,
        [int]$MaxSpaces = 96
    )

    $paragraph = $Document.Paragraphs.Item($ParagraphIndex)
    $value = Get-SuffixValue -Paragraph $paragraph
    $cleanValue = $value.Trim()

    $bestTotalSpaces = 0
    $bestEdgeScore = [double]::PositiveInfinity

    for ($totalSpaces = 0; $totalSpaces -le $MaxSpaces; $totalSpaces++) {
        $leftPad = [Math]::Floor($totalSpaces / 2)
        $rightPad = $totalSpaces - $leftPad
        $slotText = (" " * $leftPad) + $cleanValue + (" " * $rightPad)
        Set-UnderlinedFieldSlotText -Document $Document -ParagraphIndex $ParagraphIndex -SlotText $slotText
        $geometry = Get-FieldGeometry -Document $Document -ParagraphIndex $ParagraphIndex -Value $cleanValue
        $edgeScore = Get-GeometryEdgeScore -Geometry $geometry -TargetGeometry $TargetGeometry

        if ($edgeScore -lt $bestEdgeScore) {
            $bestEdgeScore = $edgeScore
            $bestTotalSpaces = $totalSpaces
        }
    }

    if ([string]::IsNullOrEmpty($cleanValue)) {
        Set-UnderlinedFieldSlotText `
            -Document $Document `
            -ParagraphIndex $ParagraphIndex `
            -SlotText (" " * $bestTotalSpaces)
        return
    }

    $bestLeftPad = [Math]::Floor($bestTotalSpaces / 2)
    $bestRightPad = $bestTotalSpaces - $bestLeftPad
    $bestScore = [double]::PositiveInfinity

    for ($leftPad = 0; $leftPad -le $bestTotalSpaces; $leftPad++) {
        $rightPad = $bestTotalSpaces - $leftPad
        $slotText = (" " * $leftPad) + $cleanValue + (" " * $rightPad)
        Set-UnderlinedFieldSlotText -Document $Document -ParagraphIndex $ParagraphIndex -SlotText $slotText
        $geometry = Get-FieldGeometry -Document $Document -ParagraphIndex $ParagraphIndex -Value $cleanValue
        $edgeScore = Get-GeometryEdgeScore -Geometry $geometry -TargetGeometry $TargetGeometry
        $centerScore = Get-ValueCenterScore -Geometry $geometry -TargetGeometry $TargetGeometry -Value $cleanValue
        $score = $edgeScore + ($centerScore * 0.25)

        if ($score -lt $bestScore) {
            $bestScore = $score
            $bestLeftPad = $leftPad
            $bestRightPad = $rightPad
        }
    }

    $finalSlotText = (" " * $bestLeftPad) + $cleanValue + (" " * $bestRightPad)
    Set-UnderlinedFieldSlotText -Document $Document -ParagraphIndex $ParagraphIndex -SlotText $finalSlotText
}

function Normalize-FixedUnderlineFields {
    param(
        $SampleDocument,
        $TargetDocument
    )

    $identifierFields = @(
        [pscustomobject]@{ ParagraphIndex = 2 },
        [pscustomobject]@{ ParagraphIndex = 3 }
    )
    $coverInfoFields = @(
        [pscustomobject]@{ ParagraphIndex = 15 },
        [pscustomobject]@{ ParagraphIndex = 16 },
        [pscustomobject]@{ ParagraphIndex = 17 },
        [pscustomobject]@{ ParagraphIndex = 18 }
    )

    $targetGeometries = @{}
    $SampleDocument.Activate()
    foreach ($field in @($identifierFields + $coverInfoFields)) {
        $paragraphIndex = [int]$field.ParagraphIndex
        $targetGeometries[$paragraphIndex] = Get-FieldGeometry `
            -Document $SampleDocument `
            -ParagraphIndex $paragraphIndex
    }

    $TargetDocument.Activate()
    foreach ($field in $identifierFields) {
        $paragraphIndex = [int]$field.ParagraphIndex
        Set-MeasuredUnderlinedField `
            -Document $TargetDocument `
            -ParagraphIndex $paragraphIndex `
            -TargetGeometry $targetGeometries[$paragraphIndex]
    }
    foreach ($field in $coverInfoFields) {
        $paragraphIndex = [int]$field.ParagraphIndex
        Set-MeasuredUnderlinedField `
            -Document $TargetDocument `
            -ParagraphIndex $paragraphIndex `
            -TargetGeometry $targetGeometries[$paragraphIndex]
    }
}

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

        Copy-CoverParagraphFont -SourceParagraph $sample.Paragraphs.Item($i) -TargetParagraph $doc.Paragraphs.Item($i)
    }

    Normalize-FixedUnderlineFields -SampleDocument $sample -TargetDocument $doc

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
