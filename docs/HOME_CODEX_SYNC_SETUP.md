# Настройка дома после исправления main

Эта памятка нужна для домашнего компьютера, где проекта еще нет в правильном
состоянии после переписывания ветки `main`.

Цель: получить версию проекта `a280973 + файлы синхронизации Codex-чатов`.

## Важно

На рабочем компьютере история `main` была переписана. Поэтому на домашнем
компьютере обычный `Pull` может не сработать или GitHub Desktop может показать
предупреждение о расхождении истории.

Это нормально.

## Вариант 1: через GitHub Desktop

1. Открой GitHub Desktop.
2. Выбери репозиторий `FillDoc`.
3. Убедись, что путь проекта:

```text
C:\Projects\FillDoc
```

4. Нажми `Fetch origin`.
5. Если появилась кнопка `Pull origin`, нажми ее.
6. Если GitHub Desktop пишет, что история расходится или не может сделать Pull,
   нужно привести локальный `main` к `origin/main`.

В GitHub Desktop это может называться примерно так:

```text
Reset to origin/main
Discard local changes
Force checkout
```

Смысл действия: локальный `main` должен стать точной копией `origin/main`.

После этого в корне проекта должны появиться файлы:

```text
codex-chat-sync.bat
codex-chat-status.bat
codex-chat-pull.bat
codex-chat-push.bat
codex-chat-repair-paths.bat
```

## Вариант 2: через PowerShell

Если GitHub Desktop путает или не дает нормально обновиться, открой PowerShell и
выполни:

```powershell
cd C:\Projects\FillDoc
git fetch origin
git reset --hard origin/main
```

После этого проверь:

```powershell
git log --oneline -3
```

Наверху должно быть примерно:

```text
defb7f1 Add Codex chat sync launchers
a280973 22:47 22/04/2026
09f1600 Add FillDoc IDs and repository conflict checks
```

## Настройка Codex-чатов

После обновления проекта:

1. Закрой Codex полностью.
2. Дождись, пока Яндекс.Диск досинхронизирует папку:

```text
C:\Users\fedeus\YandexDisk\CODEX\FillDocChats
```

3. Запусти двойным кликом:

```text
codex-chat-sync.bat
```

4. Если старые чаты находятся через поиск, но не отображаются в проекте FillDoc,
   запусти один раз:

```text
codex-chat-repair-paths.bat
```

5. После `repair-paths` снова полностью перезапусти Codex.

## Обычный порядок работы дальше

Перед началом работы на любом компьютере:

```text
codex-chat-sync.bat
```

После окончания работы на этом компьютере:

```text
codex-chat-sync.bat
```

Потом дождись, пока Яндекс.Диск закончит синхронизацию.

## Если что-то пошло не так

Проверь статус:

```text
codex-chat-status.bat
```

Нормальное состояние:

```text
Missing locally: 0
Missing in cloud: 0
```

Если GitHub Desktop снова тянет ненужные изменения, значит домашний репозиторий
не был приведен к новому `origin/main`. Повтори вариант 2 через PowerShell:

```powershell
cd C:\Projects\FillDoc
git fetch origin
git reset --hard origin/main
```
