param(
    [string]$CodexDir = "$HOME\.codex",
    [string]$SyncDir = "$HOME\Yandex.Disk\CODEX\FillDocChats",
    [string]$ProjectPath = "E:\GitHub Projects\FillDoc"
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

New-Item -ItemType Directory -Force $SyncSessionsDir | Out-Null

if (-not (Test-Path -LiteralPath $SessionsDir)) {
    throw "Sessions directory not found: $SessionsDir"
}

$projectPathNormalized = [System.IO.Path]::GetFullPath($ProjectPath)
$exportedIds = New-Object System.Collections.Generic.HashSet[string]

$sessionFiles = Get-ChildItem -LiteralPath $SessionsDir -Recurse -File -Filter "rollout-*.jsonl"

foreach ($file in $sessionFiles) {
    $firstLine = Get-Content -LiteralPath $file.FullName -TotalCount 1 -Encoding UTF8

    if (-not $firstLine) {
        continue
    }

    try {
        $meta = $firstLine | ConvertFrom-Json
    }
    catch {
        continue
    }

    $cwd = $meta.payload.cwd

    if (-not $cwd) {
        continue
    }

    $cwdNormalized = [System.IO.Path]::GetFullPath($cwd)

    if ($cwdNormalized -ne $projectPathNormalized) {
        continue
    }

    $relativePath = Get-RelativePath -BasePath $SessionsDir -TargetPath $file.FullName
    $targetPath = Join-Path $SyncSessionsDir $relativePath
    $targetDir = Split-Path -Parent $targetPath

    New-Item -ItemType Directory -Force $targetDir | Out-Null
    Copy-Item -LiteralPath $file.FullName -Destination $targetPath -Force

    [void]$exportedIds.Add($meta.payload.id)
}

if (Test-Path -LiteralPath $IndexPath) {
    $indexLines = Get-Content -LiteralPath $IndexPath -Encoding UTF8
    $selectedLines = New-Object System.Collections.Generic.List[string]

    foreach ($line in $indexLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $entry = $line | ConvertFrom-Json
        }
        catch {
            continue
        }

        if ($entry.id -and $exportedIds.Contains($entry.id)) {
            $selectedLines.Add($line)
        }
    }

    $selectedLines | Set-Content -LiteralPath $SyncIndexPath -Encoding UTF8
}

Write-Host "Exported FillDoc chats: $($exportedIds.Count)"
Write-Host "Sync directory: $SyncDir"
