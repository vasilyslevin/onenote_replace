#requires -version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:stage = 'startup'

if ($env:OS -ne 'Windows_NT') { throw 'This utility requires Windows and desktop OneNote.' }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

function Show-Exception($Exception) {
    $code = '0x{0:X8}' -f ($Exception.HResult -band 0xFFFFFFFFL)
    $message = "Failed stage: $script:stage`r`n`r`nHRESULT: $code`r`n$($Exception.Message)"
    [Windows.Forms.MessageBox]::Show(
        $message,
        'OneNote Replace',
        [Windows.Forms.MessageBoxButtons]::OK,
        [Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Get-OneNoteApplication {
    $script:stage = 'create OneNote COM object'
    return New-Object -ComObject OneNote.Application
}

function Get-CurrentPageId($OneNote) {
    $script:stage = 'GetHierarchy for currently viewed page'
    [string]$hierarchyXml = ''
    $OneNote.GetHierarchy($null, 4, [ref]$hierarchyXml)

    $document = New-Object System.Xml.XmlDocument
    $document.LoadXml($hierarchyXml)
    $manager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
    $manager.AddNamespace('one', [string]$document.DocumentElement.NamespaceURI)
    $pages = @($document.SelectNodes('//one:Page[@isCurrentlyViewed="true"]', $manager))

    if ($pages.Count -eq 0) { throw 'No currently viewed page was found. Click inside a OneNote page and try again.' }
    if ($pages.Count -gt 1) {
        $names = @($pages | ForEach-Object { $_.GetAttribute('name') }) -join ', '
        throw "Multiple pages are currently viewed: $names. Close extra OneNote windows and try again."
    }

    $id = [string]$pages[0].GetAttribute('ID')
    if ([string]::IsNullOrWhiteSpace($id)) { throw 'The currently viewed page did not contain a page ID.' }
    return $id
}

function Get-PageSnapshot($OneNote, [string]$PageId) {
    $script:stage = 'GetPageContent'
    [string]$pageXml = ''
    $OneNote.GetPageContent($PageId, [ref]$pageXml, 0, 2)

    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $true
    $document.LoadXml($pageXml)
    $modifiedText = [string]$document.DocumentElement.GetAttribute('lastModifiedTime')
    if ([string]::IsNullOrWhiteSpace($modifiedText)) { throw 'The page XML did not contain a last-modified timestamp.' }

    $modifiedDate = [DateTime]::Parse(
        $modifiedText,
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind
    )

    return [pscustomobject]@{
        PageId = $PageId
        Xml = $pageXml
        Document = $document
        LastModifiedText = $modifiedText
        LastModifiedDate = $modifiedDate
    }
}

function New-FindRegex([string]$Find, [bool]$CaseSensitive, [bool]$WholeWord) {
    if ([string]::IsNullOrEmpty($Find)) { throw 'Enter text to find.' }
    $pattern = [Regex]::Escape($Find)
    if ($WholeWord) { $pattern = '(?<![\p{L}\p{Nd}_])' + $pattern + '(?![\p{L}\p{Nd}_])' }
    $options = [Text.RegularExpressions.RegexOptions]::CultureInvariant
    if (-not $CaseSensitive) { $options = $options -bor [Text.RegularExpressions.RegexOptions]::IgnoreCase }
    return New-Object Text.RegularExpressions.Regex($pattern, $options)
}

function Analyze-Page($Snapshot, $Regex, [string]$Replacement) {
    $script:stage = 'analyze page XML'
    $document = $Snapshot.Document
    $manager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
    $manager.AddNamespace('one', [string]$document.DocumentElement.NamespaceURI)
    $blocks = @($document.SelectNodes('//one:T[not(ancestor::one:Title)]', $manager))

    $total = 0
    $changedBlocks = 0
    $hyperlinkMatches = 0
    $skippedBlocks = 0
    $previews = New-Object System.Collections.Generic.List[string]

    foreach ($block in $blocks) {
        $fragment = [string]$block.InnerText
        $wrapper = New-Object System.Xml.XmlDocument
        $wrapper.PreserveWhitespace = $true
        try { $wrapper.LoadXml('<root>' + $fragment.Replace('&nbsp;', '&#160;') + '</root>') }
        catch { $skippedBlocks++; continue }

        $before = $wrapper.DocumentElement.InnerText
        $blockCount = 0
        $blockLinkCount = 0

        foreach ($node in @($wrapper.SelectNodes('//text()'))) {
            $matches = $Regex.Matches($node.Value)
            if ($matches.Count -eq 0) { continue }

            $insideLink = $false
            $parent = $node.ParentNode
            while ($null -ne $parent) {
                if ($parent.LocalName -eq 'a') { $insideLink = $true; break }
                $parent = $parent.ParentNode
            }

            $blockCount += $matches.Count
            if ($insideLink) { $blockLinkCount += $matches.Count }
            $node.Value = $Regex.Replace(
                $node.Value,
                [Text.RegularExpressions.MatchEvaluator]{ param($match) $Replacement }
            )
        }

        if ($blockCount -eq 0) { continue }
        $total += $blockCount
        $hyperlinkMatches += $blockLinkCount
        $changedBlocks++
        $after = $wrapper.DocumentElement.InnerText
        $content = $wrapper.DocumentElement.InnerXml.Replace('&#xA0;', '&nbsp;').Replace('&#160;', '&nbsp;')
        $block.RemoveAll()
        $block.AppendChild($document.CreateCDataSection($content)) | Out-Null

        if ($previews.Count -lt 20) {
            $beforeLine = ($before -replace '\s+', ' ').Trim()
            $afterLine = ($after -replace '\s+', ' ').Trim()
            if ($beforeLine.Length -gt 150) { $beforeLine = $beforeLine.Substring(0, 147) + '...' }
            if ($afterLine.Length -gt 150) { $afterLine = $afterLine.Substring(0, 147) + '...' }
            $label = if ($blockLinkCount -gt 0) { '[INCLUDES HYPERLINK NAME]' } else { '[TEXT]' }
            $previews.Add("$label`r`nBEFORE: $beforeLine`r`nAFTER:  $afterLine")
        }
    }

    return [pscustomobject]@{
        Count = $total
        HyperlinkMatches = $hyperlinkMatches
        ChangedBlocks = $changedBlocks
        SkippedBlocks = $skippedBlocks
        Preview = $previews
        ModifiedXml = $document.OuterXml
    }
}

function Save-Backup($Snapshot) {
    $script:stage = 'save local backup'
    $folder = Join-Path $env:LOCALAPPDATA 'OneNoteLocalFindReplace\Backups'
    [IO.Directory]::CreateDirectory($folder) | Out-Null
    $safeId = $Snapshot.PageId -replace '[^A-Za-z0-9_-]', '_'
    $path = Join-Path $folder "$(Get-Date -Format 'yyyyMMdd-HHmmss-fff')-$safeId.xml"
    [IO.File]::WriteAllText($path, $Snapshot.Xml, (New-Object Text.UTF8Encoding($false)))
    return $path
}

function Get-FormAnalysis {
    $oneNote = Get-OneNoteApplication
    $pageId = Get-CurrentPageId $oneNote
    $snapshot = Get-PageSnapshot $oneNote $pageId
    $regex = New-FindRegex $findBox.Text $caseBox.Checked $wordBox.Checked
    $analysis = Analyze-Page $snapshot $regex $replaceBox.Text
    return [pscustomobject]@{ OneNote = $oneNote; Snapshot = $snapshot; Analysis = $analysis }
}

$form = New-Object Windows.Forms.Form
$form.Text = 'OneNote Replace'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object Drawing.Size(800, 630)
$form.MinimumSize = New-Object Drawing.Size(816, 669)
$form.Font = New-Object Drawing.Font('Segoe UI', 9)

$intro = New-Object Windows.Forms.Label
$intro.Location = New-Object Drawing.Point(18, 14)
$intro.Size = New-Object Drawing.Size(760, 40)
$intro.Text = 'Current page only. Hyperlink display names are included; destinations are unchanged. Titles, attachments, and OCR metadata are excluded.'
$form.Controls.Add($intro)

$findLabel = New-Object Windows.Forms.Label
$findLabel.Location = New-Object Drawing.Point(18, 64)
$findLabel.Size = New-Object Drawing.Size(95, 22)
$findLabel.Text = 'Find:'
$form.Controls.Add($findLabel)
$findBox = New-Object Windows.Forms.TextBox
$findBox.Location = New-Object Drawing.Point(118, 61)
$findBox.Size = New-Object Drawing.Size(660, 25)
$form.Controls.Add($findBox)

$replaceLabel = New-Object Windows.Forms.Label
$replaceLabel.Location = New-Object Drawing.Point(18, 100)
$replaceLabel.Size = New-Object Drawing.Size(95, 22)
$replaceLabel.Text = 'Replace with:'
$form.Controls.Add($replaceLabel)
$replaceBox = New-Object Windows.Forms.TextBox
$replaceBox.Location = New-Object Drawing.Point(118, 97)
$replaceBox.Size = New-Object Drawing.Size(660, 25)
$form.Controls.Add($replaceBox)

$caseBox = New-Object Windows.Forms.CheckBox
$caseBox.Location = New-Object Drawing.Point(118, 134)
$caseBox.Size = New-Object Drawing.Size(135, 24)
$caseBox.Text = 'Case-sensitive'
$form.Controls.Add($caseBox)
$wordBox = New-Object Windows.Forms.CheckBox
$wordBox.Location = New-Object Drawing.Point(270, 134)
$wordBox.Size = New-Object Drawing.Size(145, 24)
$wordBox.Text = 'Whole words only'
$form.Controls.Add($wordBox)

$previewButton = New-Object Windows.Forms.Button
$previewButton.Location = New-Object Drawing.Point(18, 173)
$previewButton.Size = New-Object Drawing.Size(135, 35)
$previewButton.Text = 'Preview (dry run)'
$form.Controls.Add($previewButton)
$applyButton = New-Object Windows.Forms.Button
$applyButton.Location = New-Object Drawing.Point(164, 173)
$applyButton.Size = New-Object Drawing.Size(135, 35)
$applyButton.Text = 'Apply replacement'
$form.Controls.Add($applyButton)
$statusLabel = New-Object Windows.Forms.Label
$statusLabel.Location = New-Object Drawing.Point(318, 179)
$statusLabel.Size = New-Object Drawing.Size(460, 24)
$statusLabel.Text = 'Ready. Select a page in desktop OneNote.'
$form.Controls.Add($statusLabel)

$previewBox = New-Object Windows.Forms.TextBox
$previewBox.Location = New-Object Drawing.Point(18, 222)
$previewBox.Size = New-Object Drawing.Size(760, 383)
$previewBox.Anchor = 'Top, Bottom, Left, Right'
$previewBox.Multiline = $true
$previewBox.ReadOnly = $true
$previewBox.ScrollBars = 'Vertical'
$previewBox.Font = New-Object Drawing.Font('Consolas', 9)
$form.Controls.Add($previewBox)

$previewButton.Add_Click({
    try {
        $form.UseWaitCursor = $true
        $bundle = Get-FormAnalysis
        $a = $bundle.Analysis
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('DRY RUN - no changes made')
        $lines.Add("All matches: $($a.Count)")
        $lines.Add("Matches in hyperlink names: $($a.HyperlinkMatches)")
        $lines.Add("Changed text blocks: $($a.ChangedBlocks)")
        $lines.Add("Unparseable blocks skipped: $($a.SkippedBlocks)")
        $lines.Add('')
        foreach ($item in $a.Preview) { $lines.Add($item); $lines.Add(('-' * 72)) }
        if ($a.ChangedBlocks -gt 20) { $lines.Add('Preview is limited to the first 20 changed text blocks.') }
        if ($a.Count -eq 0) { $lines.Add('No eligible matches were found.') }
        $previewBox.Text = $lines -join "`r`n"
        $statusLabel.Text = "Found $($a.Count) match(es), including $($a.HyperlinkMatches) in hyperlink names."
    }
    catch { Show-Exception $_.Exception }
    finally { $form.UseWaitCursor = $false }
})

$applyButton.Add_Click({
    try {
        $form.UseWaitCursor = $true
        $bundle = Get-FormAnalysis
        $oneNote = $bundle.OneNote
        $snapshot = $bundle.Snapshot
        $a = $bundle.Analysis

        if ($a.Count -eq 0) {
            [Windows.Forms.MessageBox]::Show('No eligible matches were found.', 'OneNote Replace', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }

        $question = "Replace $($a.Count) occurrence(s), including $($a.HyperlinkMatches) in hyperlink names?`r`n`r`nHyperlink destinations will remain unchanged. A local XML backup will be created first."
        $answer = [Windows.Forms.MessageBox]::Show($question, 'Confirm replacement', [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Warning, [Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }

        $latestId = Get-CurrentPageId $oneNote
        $latest = Get-PageSnapshot $oneNote $latestId
        if ($latest.PageId -ne $snapshot.PageId -or $latest.LastModifiedText -ne $snapshot.LastModifiedText) {
            throw 'The page changed after analysis. Nothing was written. Preview again.'
        }

        $backup = Save-Backup $snapshot
        $script:stage = 'UpdatePageContent'
        $oneNote.UpdatePageContent($a.ModifiedXml, $snapshot.LastModifiedDate, 2, $false)

        $statusLabel.Text = "Applied $($a.Count) replacement(s)."
        $previewBox.Text = "COMPLETED`r`n`r`nAll replacements: $($a.Count)`r`nHyperlink-name replacements: $($a.HyperlinkMatches)`r`nBackup: $backup"
        [Windows.Forms.MessageBox]::Show("Completed $($a.Count) replacement(s).`r`nHyperlink-name replacements: $($a.HyperlinkMatches)`r`n`r`nBackup:`r`n$backup", 'OneNote Replace', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    catch { Show-Exception $_.Exception }
    finally { $form.UseWaitCursor = $false }
})

$form.Add_Shown({ $findBox.Focus() })
[void]$form.ShowDialog()
