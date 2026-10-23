#requires -version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$inputFile = Join-Path $repoRoot 'src\OneNoteReplace.ps1'
$outputFolder = Join-Path $repoRoot 'dist'
$outputFile = Join-Path $outputFolder 'OneNoteReplace.exe'
$iconFile = Join-Path $outputFolder 'OneNoteReplace.ico'

if (-not (Test-Path -LiteralPath $inputFile)) {
    throw "Source file not found: $inputFile"
}

. $inputFile -FunctionsOnly

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host 'Installing ps2exe for the current user...'
    Install-Module ps2exe -Scope CurrentUser -Force
}

Import-Module ps2exe
[IO.Directory]::CreateDirectory($outputFolder) | Out-Null

$icon = Get-AppIcon
$iconStream = [IO.File]::Create($iconFile)
try {
    $icon.Save($iconStream)
} finally {
    $iconStream.Dispose()
    $icon.Dispose()
}

Invoke-ps2exe `
    -InputFile $inputFile `
    -OutputFile $outputFile `
    -IconFile $iconFile `
    -NoConsole `
    -Title 'OneNote Replace' `
    -Product 'OneNote Replace' `
    -Description 'Local find and replace for the current OneNote page' `
    -Company 'Vasily S. Levin' `
    -Version '0.1.0'

$hash = Get-FileHash -Algorithm SHA256 -LiteralPath $outputFile
$hashLine = "$($hash.Hash.ToLowerInvariant())  $([IO.Path]::GetFileName($outputFile))"
[IO.File]::WriteAllText((Join-Path $outputFolder 'SHA256SUMS.txt'), $hashLine + [Environment]::NewLine)
Write-Host "Built: $outputFile"
Write-Host "SHA-256: $($hash.Hash)"
