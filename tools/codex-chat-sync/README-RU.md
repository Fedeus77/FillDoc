# Прод-синхронизация локальных Codex-чатов по проекту FillDoc

Этот комплект делает 3 вещи:

1. Экспортирует только чаты нужного проекта из `threads` + `sessions`.
2. Импортирует их на другое устройство.
3. Переписывает `cwd` и `sandbox_policy`, чтобы Codex UI показывал чаты в текущем workspace.

## Что входит

- `sync_filldoc_chats.py` — основная логика.
- `export-filldoc-chats.ps1` — экспорт на исходном устройстве.
- `import-filldoc-chats.ps1` — импорт на втором устройстве.
- `repair-filldoc-chat-paths.ps1` — аварийный фикс путей в уже существующей SQLite-базе.

## Куда положить

Папка в проекте:

```text
C:\Projects\FillDoc\tools\codex-chat-sync\
```

## Обязательные условия

- На обоих устройствах должен быть установлен Python.
- Codex при импорте и repair-paths должен быть закрыт.
- На втором устройстве нужно хотя бы один раз открыть Codex, чтобы появился `state_*.sqlite`.

## Экспорт на первом устройстве

Пример для первого компьютера:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Projects\FillDoc\tools\codex-chat-sync\export-filldoc-chats.ps1" -SyncDir "C:\Users\fedus\YandexDisk\CODEX\FillDocChats" -ProjectPath "E:\GitHub Projects\FillDoc"
```

Если проект на первом устройстве лежит в другом месте — подставь свой путь.

## Импорт на втором устройстве

Пример для второго компьютера:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Projects\FillDoc\tools\codex-chat-sync\import-filldoc-chats.ps1" -SyncDir "C:\Users\fedus\YandexDisk\CODEX\FillDocChats" -ProjectPath "C:\Projects\FillDoc"
```

Скрипт:

- копирует `sessions`;
- обновляет `session_index.jsonl`;
- делает upsert в `threads`;
- меняет `cwd` на локальный путь проекта;
- переписывает `sandbox_policy.writable_roots`.

## Аварийный фикс путей

Если чаты уже импортированы, но проект переехал в другую папку:

```powershell
powershell -ExecutionPolicy Bypass -File "C:\Projects\FillDoc\tools\codex-chat-sync\repair-filldoc-chat-paths.ps1" -OldProjectPath "E:\GitHub Projects\FillDoc" -NewProjectPath "C:\Projects\FillDoc"
```

## Файлы в папке синхронизации

После нового экспорта в папке обмена должны быть:

- `manifest.filldoc.json`
- `session_index.filldoc.jsonl`
- `threads.filldoc.jsonl`
- `sessions\...`

## Важный нюанс

Старая версия экспорта, которая копировала только `sessions` и `session_index.filldoc.jsonl`, уже недостаточна.
Для полного отображения чатов в UI нужен еще `threads.filldoc.jsonl`.

## Диагностика

Если импорт пишет, что `threads.filldoc.jsonl` не найден:

1. Замени export-скрипт на новую версию.
2. Перезапусти экспорт на первом устройстве.
3. Дождись синхронизации Яндекс.Диска.
4. Повтори импорт на втором устройстве.
