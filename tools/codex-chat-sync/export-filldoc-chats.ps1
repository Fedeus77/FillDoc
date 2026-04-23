param(
    [Parameter(Mandatory = $true)]
    [string]$SyncDir,

    [Parameter(Mandatory = $true)]
    [string]$ProjectPath,

    [string]$CodexDir = "$HOME\.codex"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "sync_filldoc_chats.py"

if (-not (Test-Path $PythonScript)) {
    throw "Не найден $PythonScript"
}

python $PythonScript export --sync-dir $SyncDir --project-path $ProjectPath --codex-dir $CodexDir
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
