param(
    [Parameter(Mandatory = $true)]
    [string]$OldProjectPath,

    [Parameter(Mandatory = $true)]
    [string]$NewProjectPath,

    [string]$CodexDir = "$HOME\.codex"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "sync_filldoc_chats.py"

if (-not (Test-Path $PythonScript)) {
    throw "Не найден $PythonScript"
}

python $PythonScript repair-paths --old-project-path $OldProjectPath --new-project-path $NewProjectPath --codex-dir $CodexDir
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
