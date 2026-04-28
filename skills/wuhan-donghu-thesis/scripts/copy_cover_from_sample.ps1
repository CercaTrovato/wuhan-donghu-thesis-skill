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

function Get-DisplayWidth {
    param([string]$Text)

    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        $code = [int][char]$ch
        if ($ch -eq "`t" -or $ch -eq "`r" -or $ch -eq "`n") {
            continue
        }
        if ($ch -eq " " -or $code -lt 128) {
            $width += 1
        }
        else {
            $width += 2
        }
    }
    return $width
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

function Get-SuffixDisplayWidth {
    param(
        $Paragraph
    )

    $label = Get-FieldLabel -Paragraph $Paragraph
    if ([string]::IsNullOrEmpty($label)) {
        return 0
    }
    $text = Get-ParagraphPlainText -Paragraph $Paragraph
    $index = $text.IndexOf($label)
    if ($index -lt 0) {
        return 0
    }
    $suffix = $text.Substring($index + $label.Length)
    return Get-DisplayWidth -Text $suffix
}

function New-CenteredUnderlineSlot {
    param(
        [string]$Value,
        [int]$SlotWidth
    )

    $cleanValue = $Value.Trim()
    $valueWidth = Get-DisplayWidth -Text $cleanValue
    if ($valueWidth -ge $SlotWidth) {
        return $cleanValue
    }

    $pad = $SlotWidth - $valueWidth
    $leftPad = [Math]::Floor($pad / 2)
    $rightPad = $pad - $leftPad
    return (" " * $leftPad) + $cleanValue + (" " * $rightPad)
}

function Set-FixedUnderlinedField {
    param(
        $Document,
        [int]$ParagraphIndex,
        [int]$SlotWidth
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

    $value = Get-SuffixValue -Paragraph $paragraph
    $slotText = New-CenteredUnderlineSlot -Value $value -SlotWidth $SlotWidth

    $suffixStart = $paragraph.Range.Start + $labelIndex + $label.Length
    $suffixEnd = $paragraph.Range.End - 1
    if ($suffixEnd -lt $suffixStart) {
        $suffixEnd = $suffixStart
    }

    $suffixRange = $Document.Range($suffixStart, $suffixEnd)
    $suffixRange.Text = $slotText

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

function Get-MaxSuffixWidth {
    param(
        $Document,
        [object[]]$Fields
    )

    $maxWidth = 0
    foreach ($field in $Fields) {
        $paragraph = $Document.Paragraphs.Item([int]$field.ParagraphIndex)
        $width = Get-SuffixDisplayWidth -Paragraph $paragraph
        if ($width -gt $maxWidth) {
            $maxWidth = $width
        }
    }
    return $maxWidth
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

    $identifierSlotWidth = Get-MaxSuffixWidth -Document $SampleDocument -Fields $identifierFields
    $coverInfoSlotWidth = Get-MaxSuffixWidth -Document $SampleDocument -Fields $coverInfoFields

    foreach ($field in $identifierFields) {
        Set-FixedUnderlinedField -Document $TargetDocument -ParagraphIndex ([int]$field.ParagraphIndex) -SlotWidth $identifierSlotWidth
    }
    foreach ($field in $coverInfoFields) {
        Set-FixedUnderlinedField -Document $TargetDocument -ParagraphIndex ([int]$field.ParagraphIndex) -SlotWidth $coverInfoSlotWidth
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
