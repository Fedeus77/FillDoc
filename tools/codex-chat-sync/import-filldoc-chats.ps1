param(
    [string]$CodexDir = "$HOME\.codex",
    [string]$SyncDir = "$HOME\Yandex.Disk\CODEX\FillDocChats"
)

$ErrorActionPreference = "Stop"

$SessionsDir = Join-Path $CodexDir "sessions"
$IndexPath = Join-Path $CodexDir "session_index.jsonl"
$SyncSessionsDir = Join-Path $SyncDir "sessions"
$SyncIndexPath = Join-Path $SyncDir "session_index.filldoc.jsonl"

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\') + '\'
    $targetFullPath = [System.IO.Path]::GetFullPath($TargetPath)
    $baseUri = New-Object System.Uri($baseFullPath)
    $targetUri = New-Object System.Uri($targetFullPath)

    return [System.Uri]::UnescapeDataString(
        $baseUri.MakeRelativeUri($targetUri).ToString()
    ).Replace('/', '\')
}

if (-not (Test-Path -LiteralPath $SyncSessionsDir)) {
    throw "Sync sessions directory not found: $SyncSessionsDir"
}

New-Item -ItemType Directory -Force $SessionsDir | Out-Null

$sessionFiles = Get-ChildItem -LiteralPath $SyncSessionsDir -Recurse -File -Filter "rollout-*.jsonl"
$copied = 0

foreach ($file in $sessionFiles) {
    $relativePath = Get-RelativePath -BasePath $SyncSessionsDir -TargetPath $file.FullName
    $targetPath = Join-Path $SessionsDir $relativePath
    $targetDir = Split-Path -Parent $targetPath

    New-Item -ItemType Directory -Force $targetDir | Out-Null

    if (-not (Test-Path -LiteralPath $targetPath)) {
        Copy-Item -LiteralPath $file.FullName -Destination $targetPath
        $copied += 1
        continue
    }

    $sourceInfo = Get-Item -LiteralPath $file.FullName
    $targetInfo = Get-Item -LiteralPath $targetPath

    if ($sourceInfo.Length -gt $targetInfo.Length) {
        Copy-Item -LiteralPath $file.FullName -Destination $targetPath -Force
        $copied += 1
    }
}

$existingIds = New-Object System.Collections.Generic.HashSet[string]
$mergedLines = New-Object System.Collections.Generic.List[string]

if (Test-Path -LiteralPath $IndexPath) {
    $localLines = Get-Content -LiteralPath $IndexPath -Encoding UTF8

    foreach ($line in $localLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $entry = $line | ConvertFrom-Json
        }
        catch {
            $mergedLines.Add($line)
            continue
        }

        if ($entry.id) {
            [void]$existingIds.Add($entry.id)
        }

        $mergedLines.Add($line)
    }
}

$addedIndexLines = 0

if (Test-Path -LiteralPath $SyncIndexPath) {
    $syncLines = Get-Content -LiteralPath $SyncIndexPath -Encoding UTF8

    foreach ($line in $syncLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $entry = $line | ConvertFrom-Json
        }
        catch {
            continue
        }

        if ($entry.id -and -not $existingIds.Contains($entry.id)) {
            $mergedLines.Add($line)
            [void]$existingIds.Add($entry.id)
            $addedIndexLines += 1
        }
    }

    $mergedLines | Set-Content -LiteralPath $IndexPath -Encoding UTF8
}

Write-Host "Imported or updated session files: $copied"
Write-Host "Added index entries: $addedIndexLines"
Write-Host "Codex directory: $CodexDir"
