param(
    [ValidateSet("status", "pull", "push", "sync", "rebuild-index", "repair-paths")]
    [string]$Action = "status",

    [string]$ProjectPath = "C:\Projects\FillDoc",
    [string[]]$ProjectAliases = @("C:\Projects\FillDoc", "E:\GitHub Projects\FillDoc"),
    [string]$CodexDir = (Join-Path $env:USERPROFILE ".codex"),
    [string]$SyncRoot = (Join-Path $env:USERPROFILE "YandexDisk\CODEX\FillDocChats"),
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

function New-Utf8NoBomEncoding {
    return New-Object System.Text.UTF8Encoding($false)
}

function Normalize-PathText {
    param([string]$PathText)

    if ([string]::IsNullOrWhiteSpace($PathText)) {
        return $null
    }

    $clean = $PathText.Trim()
    if ($clean.StartsWith("\\?\")) {
        $clean = $clean.Substring(4)
    }

    try {
        $clean = [System.IO.Path]::GetFullPath($clean)
    } catch {
    }

    return $clean.TrimEnd("\").ToLowerInvariant()
}

function Test-ProjectPath {
    param([string]$Candidate)

    $candidatePath = Normalize-PathText $Candidate
    if ($candidatePath -eq $null) {
        return $false
    }

    $aliases = @($ProjectPath) + @($ProjectAliases)
    foreach ($alias in $aliases) {
        if ($candidatePath -eq (Normalize-PathText $alias)) {
            return $true
        }
    }

    return $false
}

function Test-OldProjectAlias {
    param([string]$Candidate)

    $candidatePath = Normalize-PathText $Candidate
    if ($candidatePath -eq $null) {
        return $false
    }

    foreach ($alias in @($ProjectAliases)) {
        $aliasPath = Normalize-PathText $alias
        if ($aliasPath -ne (Normalize-PathText $ProjectPath) -and $candidatePath -eq $aliasPath) {
            return $true
        }
    }

    return $false
}

function Get-RelativePathText {
    param(
        [string]$BasePath,
        [string]$FullPath
    )

    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd("\") + "\"
    $full = [System.IO.Path]::GetFullPath($FullPath)
    if (-not $full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Path '$full' is not under '$base'."
    }

    return $full.Substring($base.Length)
}

function ConvertTo-IsoUtc {
    param([datetime]$DateTime)

    return $DateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
}

function Get-JsonLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $null
    }

    try {
        return $Line | ConvertFrom-Json
    } catch {
        return $null
    }
}

function Get-RolloutMeta {
    param([string]$Path)

    $meta = [ordered]@{
        Id = $null
        Cwd = $null
        CreatedAt = $null
        UpdatedAt = (Get-Item -LiteralPath $Path).LastWriteTimeUtc
        ThreadName = $null
        Path = $Path
    }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    try {
        $lineCount = 0
        while (($line = $reader.ReadLine()) -ne $null) {
            $lineCount += 1
            $item = Get-JsonLine $line
            if ($item -eq $null) {
                continue
            }

            if ($item.timestamp) {
                try {
                    $meta.UpdatedAt = [datetime]$item.timestamp
                } catch {
                }
            }

            if ($item.type -eq "session_meta" -and $item.payload) {
                $meta.Id = $item.payload.id
                $meta.Cwd = $item.payload.cwd
                if ($item.payload.timestamp) {
                    try {
                        $meta.CreatedAt = [datetime]$item.payload.timestamp
                    } catch {
                    }
                }
            }

            if ($meta.ThreadName -eq $null) {
                $candidate = Get-UserMessageText $item
                if ($candidate) {
                    $meta.ThreadName = $candidate
                }
            }

            if ($lineCount -gt 400 -and $meta.Id -and $meta.Cwd -and $meta.ThreadName) {
                break
            }
        }
    } finally {
        $reader.Dispose()
        $stream.Dispose()
    }

    if (-not $meta.CreatedAt) {
        $meta.CreatedAt = (Get-Item -LiteralPath $Path).CreationTimeUtc
    }

    if (-not $meta.Id) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        if ($name -match "([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$") {
            $meta.Id = $Matches[1]
        }
    }

    if (-not $meta.ThreadName) {
        $meta.ThreadName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    }

    return [pscustomobject]$meta
}

function Get-UserMessageText {
    param($Item)

    $message = $null
    if ($Item.type -eq "response_item" -and $Item.payload -and $Item.payload.type -eq "message") {
        $message = $Item.payload
    } elseif ($Item.type -eq "message" -and $Item.role) {
        $message = $Item
    }

    if ($message -eq $null -or $message.role -ne "user" -or $message.content -eq $null) {
        return $null
    }

    foreach ($part in @($message.content)) {
        $text = $null
        if ($part.text) {
            $text = $part.text
        } elseif ($part.type -eq "input_text" -and $part.PSObject.Properties.Name -contains "text") {
            $text = $part.text
        }

        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $oneLine = ($text -replace "\s+", " ").Trim()
            if ($oneLine.Length -gt 120) {
                return $oneLine.Substring(0, 120)
            }
            return $oneLine
        }
    }

    return $null
}

function Get-ProjectRollouts {
    param(
        [string]$Root,
        [string]$RelativeDir
    )

    $dir = Join-Path $Root $RelativeDir
    if (-not (Test-Path -LiteralPath $dir)) {
        return @()
    }

    $items = @()
    Get-ChildItem -LiteralPath $dir -Recurse -File -Filter "rollout-*.jsonl" | ForEach-Object {
        $meta = Get-RolloutMeta $_.FullName
        if ($meta.Id -and (Test-ProjectPath $meta.Cwd)) {
            $rel = Get-RelativePathText $Root $_.FullName
            $items += [pscustomobject]@{
                Id = $meta.Id
                SourcePath = $_.FullName
                RelativePath = $rel
                UpdatedAt = $meta.UpdatedAt.ToUniversalTime()
                ThreadName = $meta.ThreadName
                IsArchived = $RelativeDir -eq "archived_sessions"
            }
        }
    }

    return $items
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($DryRun) {
            Write-Host "DRY create dir $Path"
        } else {
            New-Item -ItemType Directory -Force -Path $Path | Out-Null
        }
    }
}

function Copy-ChangedFile {
    param(
        [string]$Source,
        [string]$Destination
    )

    $sourceItem = Get-Item -LiteralPath $Source
    $shouldCopy = $true
    if (Test-Path -LiteralPath $Destination) {
        $destItem = Get-Item -LiteralPath $Destination
        $destNewer = $destItem.LastWriteTimeUtc -gt $sourceItem.LastWriteTimeUtc.AddSeconds(1)
        $sameSize = $destItem.Length -eq $sourceItem.Length
        if ($sameSize -and -not ($sourceItem.LastWriteTimeUtc -gt $destItem.LastWriteTimeUtc.AddSeconds(1))) {
            $shouldCopy = $false
        } elseif ($destNewer) {
            $shouldCopy = $false
            Write-Host "skip newer destination: $Destination"
        }
    }

    if (-not $shouldCopy) {
        return $false
    }

    Ensure-Directory ([System.IO.Path]::GetDirectoryName($Destination))
    if ($DryRun) {
        Write-Host "DRY copy $Source -> $Destination"
    } else {
        Copy-Item -LiteralPath $Source -Destination $Destination -Force
        (Get-Item -LiteralPath $Destination).LastWriteTimeUtc = $sourceItem.LastWriteTimeUtc
    }

    return $true
}

function Read-SessionIndex {
    param([string]$Path)

    $map = @{}
    if (-not (Test-Path -LiteralPath $Path)) {
        return $map
    }

    Get-Content -LiteralPath $Path -Encoding UTF8 | ForEach-Object {
        $item = Get-JsonLine $_
        if ($item -and $item.id) {
            $map[$item.id] = $item
        }
    }

    return $map
}

function Rebuild-SessionIndex {
    $indexPath = Join-Path $CodexDir "session_index.jsonl"
    $existing = Read-SessionIndex $indexPath
    $projectItems = @()
    $projectItems += Get-ProjectRollouts $CodexDir "sessions"
    $projectItems += Get-ProjectRollouts $CodexDir "archived_sessions"

    $projectById = @{}
    foreach ($item in $projectItems) {
        if (-not $projectById.ContainsKey($item.Id) -or $item.UpdatedAt -gt $projectById[$item.Id].UpdatedAt) {
            $projectById[$item.Id] = $item
        }
    }

    foreach ($id in @($projectById.Keys)) {
        if ($existing.ContainsKey($id)) {
            $projectById[$id].ThreadName = $existing[$id].thread_name
        }
    }

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($key in ($existing.Keys | Sort-Object)) {
        if (-not $projectById.ContainsKey($key)) {
            $lines.Add(($existing[$key] | ConvertTo-Json -Compress -Depth 8))
        }
    }

    foreach ($item in ($projectById.Values | Sort-Object UpdatedAt)) {
        $entry = [ordered]@{
            id = $item.Id
            thread_name = $item.ThreadName
            updated_at = (ConvertTo-IsoUtc $item.UpdatedAt)
        }
        $lines.Add(($entry | ConvertTo-Json -Compress -Depth 8))
    }

    if ($DryRun) {
        Write-Host "DRY rebuild $indexPath with $($lines.Count) entries ($($projectById.Count) FillDoc)"
    } else {
        [System.IO.File]::WriteAllLines($indexPath, [string[]]$lines, (New-Utf8NoBomEncoding))
    }

    Write-Host "rebuilt index: $($projectById.Count) FillDoc sessions, $($lines.Count) total entries"
}

function Get-TextWithSharedRead {
    param([string]$Path)

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Set-TextWithSharedWrite {
    param(
        [string]$Path,
        [string]$Text
    )

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    $writer = New-Object System.IO.StreamWriter($stream, (New-Utf8NoBomEncoding))
    try {
        $writer.Write($Text)
    } finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Get-EscapedPathVariants {
    param([string]$PathText)

    $jsonEscaped = $PathText.Replace("\", "\\")
    $doubleEscaped = $jsonEscaped.Replace("\", "\\")
    return @($PathText, $jsonEscaped, $doubleEscaped)
}

function Repair-RolloutPaths {
    param(
        [string]$Root,
        [string]$BackupLabel
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        Write-Host "skip missing root: $Root"
        return 0
    }

    $oldPaths = @($ProjectAliases | Where-Object { (Normalize-PathText $_) -ne (Normalize-PathText $ProjectPath) })
    if ($oldPaths.Count -eq 0) {
        Write-Host "no old project aliases configured"
        return 0
    }

    $backupRoot = Join-Path $Root ("backup-filldoc-path-repair-" + (Get-Date -Format "yyyyMMdd-HHmmss") + "-" + $BackupLabel)
    $changed = 0
    $newVariants = Get-EscapedPathVariants $ProjectPath

    $files = @()
    foreach ($dirName in @("sessions", "archived_sessions")) {
        $dir = Join-Path $Root $dirName
        if (Test-Path -LiteralPath $dir) {
            $files += Get-ChildItem -LiteralPath $dir -Recurse -File -Filter "rollout-*.jsonl"
        }
    }

    foreach ($file in $files) {
        $meta = Get-RolloutMeta $file.FullName
        if (-not (Test-OldProjectAlias $meta.Cwd)) {
            continue
        }

        $text = Get-TextWithSharedRead $file.FullName
        $updated = $text

        foreach ($oldPath in $oldPaths) {
            $oldVariants = Get-EscapedPathVariants $oldPath
            for ($i = 0; $i -lt $oldVariants.Count; $i += 1) {
                $updated = $updated.Replace($oldVariants[$i], $newVariants[$i])
            }
        }

        if ($updated -eq $text) {
            continue
        }

        $relative = Get-RelativePathText $Root $file.FullName
        $backupPath = Join-Path $backupRoot $relative
        Ensure-Directory ([System.IO.Path]::GetDirectoryName($backupPath))

        if ($DryRun) {
            Write-Host "DRY repair $($file.FullName)"
        } else {
            Copy-Item -LiteralPath $file.FullName -Destination $backupPath -Force
            Set-TextWithSharedWrite $file.FullName $updated
        }

        $changed += 1
    }

    Write-Host "path repair in ${Root}: $changed files changed"
    if ($changed -gt 0) {
        Write-Host "backup: $backupRoot"
    }

    return $changed
}

function Push-ProjectRollouts {
    Ensure-Directory $SyncRoot
    $items = @()
    $items += Get-ProjectRollouts $CodexDir "sessions"
    $items += Get-ProjectRollouts $CodexDir "archived_sessions"

    $copied = 0
    foreach ($item in $items) {
        $destination = Join-Path $SyncRoot $item.RelativePath
        if (Copy-ChangedFile $item.SourcePath $destination) {
            $copied += 1
        }
    }

    Write-Host "push complete: $copied copied, $($items.Count) FillDoc rollout files found"
}

function Pull-ProjectRollouts {
    if (-not (Test-Path -LiteralPath $SyncRoot)) {
        throw "SyncRoot does not exist: $SyncRoot"
    }

    $items = @()
    $items += Get-ProjectRollouts $SyncRoot "sessions"
    $items += Get-ProjectRollouts $SyncRoot "archived_sessions"

    $copied = 0
    foreach ($item in $items) {
        $destination = Join-Path $CodexDir $item.RelativePath
        if (Copy-ChangedFile $item.SourcePath $destination) {
            $copied += 1
        }
    }

    Write-Host "pull complete: $copied copied, $($items.Count) cloud FillDoc rollout files found"
}

function Show-Status {
    $local = @()
    $local += Get-ProjectRollouts $CodexDir "sessions"
    $local += Get-ProjectRollouts $CodexDir "archived_sessions"

    $cloud = @()
    if (Test-Path -LiteralPath $SyncRoot) {
        $cloud += Get-ProjectRollouts $SyncRoot "sessions"
        $cloud += Get-ProjectRollouts $SyncRoot "archived_sessions"
    }

    $localIds = @{}
    foreach ($item in $local) {
        $localIds[$item.Id] = $true
    }

    $cloudIds = @{}
    foreach ($item in $cloud) {
        $cloudIds[$item.Id] = $true
    }

    $missingLocal = @($cloudIds.Keys | Where-Object { -not $localIds.ContainsKey($_) } | Sort-Object)
    $missingCloud = @($localIds.Keys | Where-Object { -not $cloudIds.ContainsKey($_) } | Sort-Object)

    Write-Host "ProjectPath: $ProjectPath"
    Write-Host "CodexDir:    $CodexDir"
    Write-Host "SyncRoot:    $SyncRoot"
    Write-Host "Local FillDoc sessions: $($localIds.Count)"
    Write-Host "Cloud FillDoc sessions: $($cloudIds.Count)"
    Write-Host "Missing locally: $($missingLocal.Count)"
    foreach ($id in ($missingLocal | Select-Object -First 10)) {
        Write-Host "  local <- $id"
    }
    Write-Host "Missing in cloud: $($missingCloud.Count)"
    foreach ($id in ($missingCloud | Select-Object -First 10)) {
        Write-Host "  cloud <- $id"
    }
}

Write-Host "Codex FillDoc chat sync: $Action"

switch ($Action) {
    "status" {
        Show-Status
    }
    "pull" {
        Pull-ProjectRollouts
        Rebuild-SessionIndex
    }
    "push" {
        Push-ProjectRollouts
    }
    "sync" {
        Pull-ProjectRollouts
        Rebuild-SessionIndex
        Push-ProjectRollouts
    }
    "rebuild-index" {
        Rebuild-SessionIndex
    }
    "repair-paths" {
        Repair-RolloutPaths $CodexDir "local" | Out-Null
        Repair-RolloutPaths $SyncRoot "cloud" | Out-Null
        Rebuild-SessionIndex
        Push-ProjectRollouts
    }
}
