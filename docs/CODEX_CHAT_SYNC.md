# Codex chat sync for FillDoc

This project keeps Codex chats in:

- `C:\Users\<user>\.codex\sessions`
- `C:\Users\<user>\.codex\archived_sessions`
- `C:\Users\<user>\.codex\session_index.jsonl`

Do not sync the whole `.codex` directory. It also contains auth, caches, sqlite state,
logs, plugins, and local machine state.

The repo script syncs only FillDoc rollout files whose `cwd` is one of the known
FillDoc paths. By default it includes the current `C:\Projects\FillDoc` path and the
older `E:\GitHub Projects\FillDoc` path used by earlier chats. This works on all three
machines because the current project path is the same, while the user profile can be
either `fedus` or `fedeus`.

## Paths

Default local Codex directory:

```powershell
%USERPROFILE%\.codex
```

Default Yandex Disk sync directory:

```powershell
%USERPROFILE%\YandexDisk\CODEX\FillDocChats
```

## Daily use

Double-click launchers are available in the project root:

```text
codex-chat-pull.bat
codex-chat-push.bat
codex-chat-sync.bat
codex-chat-status.bat
codex-chat-repair-paths.bat
```

Before starting work on a machine:

```powershell
.\tools\codex-chat-sync.ps1 -Action pull
```

After finishing work on a machine:

```powershell
.\tools\codex-chat-sync.ps1 -Action push
```

For a two-way sync:

```powershell
.\tools\codex-chat-sync.ps1 -Action sync
```

To only inspect differences:

```powershell
.\tools\codex-chat-sync.ps1 -Action status
```

To repair old chats that still point to an obsolete FillDoc path:

```powershell
.\tools\codex-chat-sync.ps1 -Action repair-paths
```

If Codex still shows old chats outside the FillDoc project after that, repair the local
Codex UI cache:

```powershell
uv run --no-project python .\tools\codex-sqlite-path-repair.py
```

## Notes

- Run the commands from `C:\Projects\FillDoc`.
- Let Yandex Disk finish syncing before switching to another computer.
- If Codex is open while running `pull`, restart Codex if the imported chats do not
  appear immediately.
- Old files like `threads.filldoc.jsonl` and `manifest.filldoc.json` are legacy export
  artifacts. The new flow does not need them.
