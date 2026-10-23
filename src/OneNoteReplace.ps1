#requires -version 5.1
param(
    [switch]$FunctionsOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:stage = 'startup'

if ($env:OS -ne 'Windows_NT') { throw 'This utility requires Windows and desktop OneNote.' }

function Get-AppIcon {
    Add-Type -AssemblyName System.Drawing
    $bytes = [Convert]::FromBase64String('AAABAAEAICAAAAEAIAAlAQAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAgAAAAIAgGAAAAc3p69AAAAOxJREFUeNpjYBjpgBFdIEao7T8tLVzyrooRpwNwWS7NrkI1Bzz9eQfFEUwDHQWD0wHmCg4M5goOdHEACymKO56FkmVJhdRq0hxw8sEBukUBC7V8MrwSIT0BC65cgC0t0CIRDs4QwJULRhPhkCwJCUUbTeqCWOF2oh1NlZIQZgmyxTD24reVDB3PQnHqZcKVCyitDxYfeUD9NEAo2NEtXXzkAUT8bSXtHbD4bSVKFCw+8oAh1kYBr+U0zYbEWE61XABLYIvfVsItJcZyqoYAcirHZjmu7EiXZjmy5bHC7fRvlpPcJqSnIwa8azYKADUcYxQbgZRMAAAAAElFTkSuQmCC')
    $stream = New-Object IO.MemoryStream(,$bytes)
    $icon = $null
    try {
        $icon = New-Object Drawing.Icon($stream)
        return $icon.Clone()
    } finally {
        if ($null -ne $icon) { $icon.Dispose() }
        $stream.Dispose()
    }
}

if ($FunctionsOnly) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()

function Show-Exception($Exception) {
    $code = '0x{0:X8}' -f ($Exception.HResult -band 0xFFFFFFFFL)
    $message = "Failed stage: $script:stage`r`n`r`nHRESULT: $code`r`n$($Exception.Message)"
    [Windows.Forms.MessageBox]::Show($message, 'OneNote Replace', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error) | Out-Null
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
    $modifiedDate = [DateTime]::Parse($modifiedText, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
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

function Replace-TextInMarkup([string]$Fragment, $Regex, [string]$EscapedReplacement) {
    $entityPattern = '^(&#[0-9]+;|&#x[0-9A-Fa-f]+;|&[A-Za-z][A-Za-z0-9]+;)$'
    $output = New-Object Text.StringBuilder
    $buffer = New-Object Text.StringBuilder
    $inTag = $false
    $quote = [char]0
    $linkDepth = 0
    $count = 0
    $linkCount = 0

    for ($i = 0; $i -lt $Fragment.Length; $i++) {
        $character = $Fragment[$i]
        if (-not $inTag) {
            if ($character -eq '<') {
                $textPart = $buffer.ToString()
                $buffer.Clear() | Out-Null
                if ($textPart.Length -gt 0) {
                    foreach ($chunkValue in [Text.RegularExpressions.Regex]::Split($textPart, '(&#[0-9]+;|&#x[0-9A-Fa-f]+;|&[A-Za-z][A-Za-z0-9]+;)')) {
                        $chunk = [string]$chunkValue
                        if ($chunk -match $entityPattern) { $output.Append($chunk) | Out-Null; continue }
                        $matches = $Regex.Matches($chunk)
                        if ($matches.Count -gt 0) {
                            $count += $matches.Count
                            if ($linkDepth -gt 0) { $linkCount += $matches.Count }
                            $chunk = $Regex.Replace($chunk, [Text.RegularExpressions.MatchEvaluator]{ param($match) $EscapedReplacement })
                        }
                        $output.Append($chunk) | Out-Null
                    }
                }
                $inTag = $true
                $quote = [char]0
                $buffer.Append($character) | Out-Null
            }
            else { $buffer.Append($character) | Out-Null }
        }
        else {
            $buffer.Append($character) | Out-Null
            if ($quote -ne [char]0) {
                if ($character -eq $quote) { $quote = [char]0 }
            }
            elseif ($character -eq '"' -or $character -eq "'") { $quote = $character }
            elseif ($character -eq '>') {
                $tag = $buffer.ToString()
                $buffer.Clear() | Out-Null
                $output.Append($tag) | Out-Null
                if ($tag -match '^<\s*/\s*a\s*>') {
                    if ($linkDepth -gt 0) { $linkDepth-- }
                }
                elseif ($tag -match '^<\s*a(?:\s|>)' -and $tag -notmatch '/\s*>$') { $linkDepth++ }
                $inTag = $false
            }
        }
    }

    if ($inTag) { return [pscustomobject]@{ Success = $false; Count = 0; LinkCount = 0; Content = $Fragment } }

    $textPart = $buffer.ToString()
    if ($textPart.Length -gt 0) {
        foreach ($chunkValue in [Text.RegularExpressions.Regex]::Split($textPart, '(&#[0-9]+;|&#x[0-9A-Fa-f]+;|&[A-Za-z][A-Za-z0-9]+;)')) {
            $chunk = [string]$chunkValue
            if ($chunk -match $entityPattern) { $output.Append($chunk) | Out-Null; continue }
            $matches = $Regex.Matches($chunk)
            if ($matches.Count -gt 0) {
                $count += $matches.Count
                if ($linkDepth -gt 0) { $linkCount += $matches.Count }
                $chunk = $Regex.Replace($chunk, [Text.RegularExpressions.MatchEvaluator]{ param($match) $EscapedReplacement })
            }
            $output.Append($chunk) | Out-Null
        }
    }

    return [pscustomobject]@{ Success = $true; Count = $count; LinkCount = $linkCount; Content = $output.ToString() }
}

function Analyze-Page($Snapshot, $Regex, [string]$Replacement) {
    $script:stage = 'analyze page and build minimal outline update'
    $document = $Snapshot.Document
    $namespace = [string]$document.DocumentElement.NamespaceURI
    $manager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
    $manager.AddNamespace('one', $namespace)
    $blocks = @($document.SelectNodes('//one:T[not(ancestor::one:Title)]', $manager))
    $escapedReplacement = [Security.SecurityElement]::Escape($Replacement)
    if ($null -eq $escapedReplacement) { $escapedReplacement = '' }

    $total = 0
    $changedBlocks = 0
    $hyperlinkMatches = 0
    $skippedBlocks = 0
    $previews = New-Object System.Collections.Generic.List[string]
    $changedOutlineIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $changedOutlines = New-Object 'System.Collections.Generic.List[System.Xml.XmlElement]'

    foreach ($block in $blocks) {
        $fragment = [string]$block.InnerText
        $result = Replace-TextInMarkup $fragment $Regex $escapedReplacement
        if (-not $result.Success) { $skippedBlocks++; continue }
        if ($result.Count -eq 0) { continue }

        $outline = $block.SelectSingleNode('ancestor::one:Outline[1]', $manager)
        if ($null -eq $outline) { $skippedBlocks++; continue }
        $outlineId = [string]$outline.GetAttribute('objectID')
        if ([string]::IsNullOrWhiteSpace($outlineId)) { $skippedBlocks++; continue }

        $total += $result.Count
        $hyperlinkMatches += $result.LinkCount
        $changedBlocks++
        $newFragment = [string]$result.Content

        if ($block.ChildNodes.Count -eq 1 -and $block.FirstChild.NodeType -eq [System.Xml.XmlNodeType]::CDATA) {
            $block.FirstChild.Value = $newFragment
        }
        else {
            $block.RemoveAll()
            $block.AppendChild($document.CreateCDataSection($newFragment)) | Out-Null
        }

        if ($changedOutlineIds.Add($outlineId)) { $changedOutlines.Add($outline) }

        if ($previews.Count -lt 20) {
            $beforePlain = [Net.WebUtility]::HtmlDecode(($fragment -replace '<[^>]+>', ''))
            $afterPlain = [Net.WebUtility]::HtmlDecode(($newFragment -replace '<[^>]+>', ''))
            $beforeLine = ($beforePlain -replace '\s+', ' ').Trim()
            $afterLine = ($afterPlain -replace '\s+', ' ').Trim()
            if ($beforeLine.Length -gt 150) { $beforeLine = $beforeLine.Substring(0, 147) + '...' }
            if ($afterLine.Length -gt 150) { $afterLine = $afterLine.Substring(0, 147) + '...' }
            $label = if ($result.LinkCount -gt 0) { '[INCLUDES HYPERLINK NAME]' } else { '[TEXT]' }
            $previews.Add("$label`r`nBEFORE: $beforeLine`r`nAFTER:  $afterLine")
        }
    }

    $updateDocument = New-Object System.Xml.XmlDocument
    $updateDocument.PreserveWhitespace = $true
    $page = $updateDocument.CreateElement('one', 'Page', $namespace)
    $page.SetAttribute('ID', $Snapshot.PageId)
    $updateDocument.AppendChild($page) | Out-Null
    foreach ($outline in $changedOutlines) {
        $page.AppendChild($updateDocument.ImportNode($outline, $true)) | Out-Null
    }

    return [pscustomobject]@{
        Count = $total
        HyperlinkMatches = $hyperlinkMatches
        ChangedBlocks = $changedBlocks
        ChangedOutlines = $changedOutlines.Count
        SkippedBlocks = $skippedBlocks
        Preview = $previews
        UpdateXml = $updateDocument.OuterXml
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
    $oneNote = New-Object -ComObject OneNote.Application
    $pageId = Get-CurrentPageId $oneNote
    $snapshot = Get-PageSnapshot $oneNote $pageId
    $regex = New-FindRegex $findBox.Text $caseBox.Checked $wordBox.Checked
    $analysis = Analyze-Page $snapshot $regex $replaceBox.Text
    return [pscustomobject]@{ OneNote = $oneNote; Snapshot = $snapshot; Analysis = $analysis }
}

$form = New-Object Windows.Forms.Form
$form.Text = 'OneNote Replace v3.3'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object Drawing.Size(800, 640)
$form.MinimumSize = New-Object Drawing.Size(816, 679)
$icon = Get-AppIcon
$form.Icon = $icon
$form.Font = New-Object Drawing.Font('Segoe UI', 9)

$intro = New-Object Windows.Forms.Label
$intro.Location = New-Object Drawing.Point(18, 14)
$intro.Size = New-Object Drawing.Size(760, 42)
$intro.Text = 'Current page only. Only outlines containing matches are updated. Hyperlink names are included; destinations are preserved.'
$form.Controls.Add($intro)

$findLabel = New-Object Windows.Forms.Label
$findLabel.Location = New-Object Drawing.Point(18, 66)
$findLabel.Size = New-Object Drawing.Size(95, 22)
$findLabel.Text = 'Find:'
$form.Controls.Add($findLabel)
$findBox = New-Object Windows.Forms.TextBox
$findBox.Location = New-Object Drawing.Point(118, 63)
$findBox.Size = New-Object Drawing.Size(660, 25)
$form.Controls.Add($findBox)

$replaceLabel = New-Object Windows.Forms.Label
$replaceLabel.Location = New-Object Drawing.Point(18, 102)
$replaceLabel.Size = New-Object Drawing.Size(95, 22)
$replaceLabel.Text = 'Replace with:'
$form.Controls.Add($replaceLabel)
$replaceBox = New-Object Windows.Forms.TextBox
$replaceBox.Location = New-Object Drawing.Point(118, 99)
$replaceBox.Size = New-Object Drawing.Size(660, 25)
$form.Controls.Add($replaceBox)

$caseBox = New-Object Windows.Forms.CheckBox
$caseBox.Location = New-Object Drawing.Point(118, 136)
$caseBox.Size = New-Object Drawing.Size(135, 24)
$caseBox.Text = 'Case-sensitive'
$form.Controls.Add($caseBox)
$wordBox = New-Object Windows.Forms.CheckBox
$wordBox.Location = New-Object Drawing.Point(270, 136)
$wordBox.Size = New-Object Drawing.Size(145, 24)
$wordBox.Text = 'Whole words only'
$form.Controls.Add($wordBox)

$previewButton = New-Object Windows.Forms.Button
$previewButton.Location = New-Object Drawing.Point(18, 175)
$previewButton.Size = New-Object Drawing.Size(135, 35)
$previewButton.Text = 'Preview (dry run)'
$form.Controls.Add($previewButton)
$applyButton = New-Object Windows.Forms.Button
$applyButton.Location = New-Object Drawing.Point(164, 175)
$applyButton.Size = New-Object Drawing.Size(135, 35)
$applyButton.Text = 'Apply replacement'
$form.Controls.Add($applyButton)
$statusLabel = New-Object Windows.Forms.Label
$statusLabel.Location = New-Object Drawing.Point(318, 181)
$statusLabel.Size = New-Object Drawing.Size(460, 24)
$statusLabel.Text = 'Ready. Use a white page canvas for updates.'
$form.Controls.Add($statusLabel)

$previewBox = New-Object Windows.Forms.TextBox
$previewBox.Location = New-Object Drawing.Point(18, 224)
$previewBox.Size = New-Object Drawing.Size(760, 391)
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
        $lines.Add("Page-level outlines to update: $($a.ChangedOutlines)")
        $lines.Add("Unsafe blocks skipped: $($a.SkippedBlocks)")
        $lines.Add('')
        foreach ($item in $a.Preview) { $lines.Add($item); $lines.Add(('-' * 72)) }
        if ($a.ChangedBlocks -gt 20) { $lines.Add('Preview is limited to the first 20 changed text blocks.') }
        if ($a.Count -eq 0) { $lines.Add('No eligible matches were found.') }
        $previewBox.Text = $lines -join "`r`n"
        $statusLabel.Text = "Found $($a.Count) match(es) in $($a.ChangedOutlines) outline(s)."
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

        $question = "Replace $($a.Count) occurrence(s) in $($a.ChangedOutlines) outline(s)?`r`n`r`nUse a white page canvas. Hyperlink destinations will remain unchanged. A local XML backup will be created first."
        $answer = [Windows.Forms.MessageBox]::Show($question, 'Confirm replacement', [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Warning, [Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }

        $latestId = Get-CurrentPageId $oneNote
        $latest = Get-PageSnapshot $oneNote $latestId
        if ($latest.PageId -ne $snapshot.PageId -or $latest.LastModifiedText -ne $snapshot.LastModifiedText) {
            throw 'The page changed after analysis. Nothing was written. Preview again.'
        }

        $backup = Save-Backup $snapshot
        $script:stage = 'UpdatePageContent with changed outlines only'
        $oneNote.UpdatePageContent($a.UpdateXml, $snapshot.LastModifiedDate, 2, $false)

        $statusLabel.Text = "Applied $($a.Count) replacement(s)."
        $previewBox.Text = "COMPLETED`r`n`r`nAll replacements: $($a.Count)`r`nHyperlink-name replacements: $($a.HyperlinkMatches)`r`nUpdated outlines: $($a.ChangedOutlines)`r`nBackup: $backup"
        [Windows.Forms.MessageBox]::Show("Completed $($a.Count) replacement(s) in $($a.ChangedOutlines) outline(s).`r`n`r`nBackup:`r`n$backup", 'OneNote Replace', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    catch { Show-Exception $_.Exception }
    finally { $form.UseWaitCursor = $false }
})

$form.Add_Shown({ $findBox.Focus() })
try { [void]$form.ShowDialog() } finally { $form.Dispose(); $icon.Dispose() }
