#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
import sqlite3
import sys
from pathlib import Path
from typing import Any, Iterable

THREAD_COLUMNS = [
    "id",
    "rollout_path",
    "created_at",
    "updated_at",
    "source",
    "model_provider",
    "cwd",
    "title",
    "sandbox_policy",
    "approval_mode",
    "tokens_used",
    "has_user_event",
    "archived",
    "archived_at",
    "git_sha",
    "git_branch",
    "git_origin_url",
    "cli_version",
    "first_user_message",
    "agent_nickname",
    "agent_role",
    "memory_mode",
    "model",
    "reasoning_effort",
    "agent_path",
    "created_at_ms",
    "updated_at_ms",
]


def eprint(*args: Any) -> None:
    print(*args, file=sys.stderr)


def canonical_path(path: str | Path) -> str:
    s = str(Path(path).expanduser().resolve())
    if s.startswith('\\\\?\\'):
        s = s[4:]
    return str(Path(s))


def extended_windows_path(path: str | Path) -> str:
    s = canonical_path(path)
    if s.startswith('\\\\?\\'):
        return s
    return '\\\\?\\' + s


def project_aliases(project_path: str | Path) -> set[str]:
    raw = canonical_path(project_path)
    aliases = {
        raw,
        extended_windows_path(raw),
        raw.replace('/', '\\'),
        extended_windows_path(raw).replace('/', '\\'),
    }
    return aliases


def detect_codex_dir(explicit: str | None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve()
    return (Path.home() / '.codex').resolve()


def latest_state_db(codex_dir: Path) -> Path:
    candidates = sorted(codex_dir.glob('state_*.sqlite'), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise SystemExit(
            f"Не найден state_*.sqlite в {codex_dir}. Сначала открой Codex хотя бы один раз и создай/открой любой чат."
        )
    return candidates[0]


def iso_from_seconds(seconds: int | None, millis: int | None = None) -> str:
    if millis:
        d = dt.datetime.fromtimestamp(millis / 1000, tz=dt.timezone.utc)
    elif seconds:
        d = dt.datetime.fromtimestamp(seconds, tz=dt.timezone.utc)
    else:
        d = dt.datetime.now(tz=dt.timezone.utc)
    return d.isoformat().replace('+00:00', 'Z')


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def load_jsonl(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    rows: list[dict[str, Any]] = []
    for line in path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if not line:
            continue
        rows.append(json.loads(line))
    return rows


def dump_jsonl(path: Path, rows: Iterable[dict[str, Any]]) -> None:
    with path.open('w', encoding='utf-8', newline='\n') as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + '\n')


def rewrite_sandbox_policy(policy_text: str, new_project_path: str) -> str:
    if not policy_text:
        return policy_text
    try:
        payload = json.loads(policy_text)
    except json.JSONDecodeError:
        return policy_text
    writable_roots = payload.get('writable_roots')
    if isinstance(writable_roots, list) and writable_roots:
        payload['writable_roots'] = [new_project_path]
    return json.dumps(payload, ensure_ascii=False, separators=(',', ':'))


def path_with_rebased_prefix(value: str | None, old_root: str, new_root: str) -> str | None:
    if not value:
        return value
    canonical_old = canonical_path(old_root)
    canonical_new = canonical_path(new_root)
    canonical_value = canonical_path(value)
    if canonical_value == canonical_old:
        return new_root
    if canonical_value.startswith(canonical_old + '\\'):
        suffix = canonical_value[len(canonical_old):].lstrip('\\/')
        return str(Path(canonical_new) / suffix)
    return value


def detect_workspace_cwd_representation(con: sqlite3.Connection, requested_project_path: str) -> str:
    aliases = project_aliases(requested_project_path)
    rows = con.execute(
        'SELECT cwd, updated_at_ms, updated_at FROM threads ORDER BY COALESCE(updated_at_ms, updated_at * 1000) DESC'
    ).fetchall()
    for cwd, _updated_at_ms, _updated_at in rows:
        if cwd in aliases or canonical_path(cwd) == canonical_path(requested_project_path):
            return cwd
    return canonical_path(requested_project_path)


def select_threads_for_project(con: sqlite3.Connection, project_path: str) -> list[dict[str, Any]]:
    aliases = project_aliases(project_path)
    cur = con.execute(f"SELECT {', '.join(THREAD_COLUMNS)} FROM threads")
    rows: list[dict[str, Any]] = []
    for db_row in cur.fetchall():
        row = dict(zip(THREAD_COLUMNS, db_row))
        cwd = row.get('cwd')
        if cwd in aliases or canonical_path(cwd) == canonical_path(project_path):
            rows.append(row)
    rows.sort(key=lambda r: (r.get('updated_at_ms') or (r.get('updated_at') or 0) * 1000), reverse=True)
    return rows


def export_threads(sync_dir: Path, codex_dir: Path, project_path: str) -> None:
    state_db = latest_state_db(codex_dir)
    with sqlite3.connect(state_db) as con:
        threads = select_threads_for_project(con, project_path)

    if not threads:
        raise SystemExit(
            'Не найдено ни одного чата для указанного ProjectPath в таблице threads. '\
            'Проверь путь проекта и открой хотя бы один чат Codex из нужного workspace.'
        )

    ensure_dir(sync_dir)
    ensure_dir(sync_dir / 'sessions')

    exported_rows: list[dict[str, Any]] = []
    index_rows: list[dict[str, Any]] = []
    copied_sessions = 0

    for row in threads:
        src_rollout = Path(str(row['rollout_path']))
        try:
            relative_rollout = src_rollout.relative_to(codex_dir)
        except ValueError:
            if 'sessions' in src_rollout.parts:
                idx = src_rollout.parts.index('sessions')
                relative_rollout = Path(*src_rollout.parts[idx:])
            else:
                raise SystemExit(f"Не удалось вычислить относительный путь rollout_path: {src_rollout}")

        dst_rollout = sync_dir / relative_rollout
        ensure_dir(dst_rollout.parent)
        if src_rollout.exists():
            shutil.copy2(src_rollout, dst_rollout)
            copied_sessions += 1
        else:
            eprint(f"[WARN] Файл сессии не найден, пропускаю копирование: {src_rollout}")

        export_row = dict(row)
        export_row['rollout_rel_path'] = str(relative_rollout).replace('/', '\\')
        export_row['source_project_path'] = canonical_path(project_path)
        exported_rows.append(export_row)

        index_rows.append(
            {
                'id': row['id'],
                'thread_name': row['title'],
                'updated_at': iso_from_seconds(row.get('updated_at'), row.get('updated_at_ms')),
            }
        )

    dump_jsonl(sync_dir / 'threads.filldoc.jsonl', exported_rows)
    dump_jsonl(sync_dir / 'session_index.filldoc.jsonl', index_rows)

    manifest = {
        'schema_version': 2,
        'exported_at': dt.datetime.now(tz=dt.timezone.utc).isoformat().replace('+00:00', 'Z'),
        'codex_dir': str(codex_dir),
        'project_path': canonical_path(project_path),
        'thread_count': len(exported_rows),
        'copied_session_files': copied_sessions,
    }
    (sync_dir / 'manifest.filldoc.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')

    print(f'Exported threads: {len(exported_rows)}')
    print(f'Copied session files: {copied_sessions}')
    print(f'Sync directory: {sync_dir}')
    print(f'State DB: {state_db}')


def read_existing_index(index_path: Path) -> dict[str, dict[str, Any]]:
    rows = load_jsonl(index_path)
    return {str(row['id']): row for row in rows if 'id' in row}


def merge_index(index_path: Path, imported_rows: list[dict[str, Any]]) -> int:
    merged = read_existing_index(index_path)
    before = len(merged)
    for row in imported_rows:
        merged[str(row['id'])] = row
    sorted_rows = sorted(
        merged.values(),
        key=lambda x: x.get('updated_at', ''),
        reverse=True,
    )
    dump_jsonl(index_path, sorted_rows)
    return len(merged) - before


def thread_exists(con: sqlite3.Connection, thread_id: str) -> bool:
    row = con.execute('SELECT 1 FROM threads WHERE id = ? LIMIT 1', (thread_id,)).fetchone()
    return row is not None


def upsert_threads(con: sqlite3.Connection, rows: list[dict[str, Any]]) -> tuple[int, int]:
    inserted = 0
    updated = 0
    placeholders = ', '.join(['?'] * len(THREAD_COLUMNS))
    update_clause = ', '.join([f'{col}=excluded.{col}' for col in THREAD_COLUMNS if col != 'id'])
    sql = f"""
        INSERT INTO threads ({', '.join(THREAD_COLUMNS)})
        VALUES ({placeholders})
        ON CONFLICT(id) DO UPDATE SET
        {update_clause}
    """
    for row in rows:
        existed = thread_exists(con, str(row['id']))
        values = [row.get(col) for col in THREAD_COLUMNS]
        con.execute(sql, values)
        if existed:
            updated += 1
        else:
            inserted += 1
    return inserted, updated


def import_threads(sync_dir: Path, codex_dir: Path, project_path: str) -> None:
    threads_path = sync_dir / 'threads.filldoc.jsonl'
    if not threads_path.exists():
        raise SystemExit(
            f'В {sync_dir} нет файла threads.filldoc.jsonl. '\
            'Нужно заменить export-скрипт на новую версию и заново выполнить экспорт на исходном устройстве.'
        )

    imported_threads = load_jsonl(threads_path)
    imported_index = load_jsonl(sync_dir / 'session_index.filldoc.jsonl')
    if not imported_threads:
        raise SystemExit('threads.filldoc.jsonl пустой — импортировать нечего.')

    state_db = latest_state_db(codex_dir)
    ensure_dir(codex_dir / 'sessions')

    copied_files = 0
    transformed_rows: list[dict[str, Any]] = []

    with sqlite3.connect(state_db) as con:
        local_workspace_cwd = detect_workspace_cwd_representation(con, project_path)
        raw_project_path = canonical_path(project_path)

        for row in imported_threads:
            old_project = row.get('source_project_path') or row.get('cwd') or raw_project_path
            rollout_rel = row.get('rollout_rel_path')
            if not rollout_rel:
                raise SystemExit(f"У записи thread {row.get('id')} отсутствует rollout_rel_path")

            src_rollout = sync_dir / Path(rollout_rel)
            dst_rollout = codex_dir / Path(rollout_rel)
            ensure_dir(dst_rollout.parent)
            if src_rollout.exists():
                shutil.copy2(src_rollout, dst_rollout)
                copied_files += 1
            else:
                eprint(f"[WARN] Не найден rollout-файл в sync dir: {src_rollout}")

            new_row = {col: row.get(col) for col in THREAD_COLUMNS}
            new_row['rollout_path'] = str(dst_rollout)
            new_row['cwd'] = local_workspace_cwd
            new_row['sandbox_policy'] = rewrite_sandbox_policy(str(new_row.get('sandbox_policy') or ''), raw_project_path)
            new_row['agent_path'] = path_with_rebased_prefix(new_row.get('agent_path'), str(old_project), raw_project_path)
            transformed_rows.append(new_row)

        inserted, updated = upsert_threads(con, transformed_rows)
        con.commit()

    added_index_entries = merge_index(codex_dir / 'session_index.jsonl', imported_index)

    print(f'Imported or updated session files: {copied_files}')
    print(f'Inserted thread rows: {inserted}')
    print(f'Updated thread rows: {updated}')
    print(f'Added index entries: {added_index_entries}')
    print(f'Codex directory: {codex_dir}')
    print(f'State DB: {state_db}')
    print(f'Workspace cwd used: {canonical_path(project_path)}')


def repair_paths(codex_dir: Path, old_project_path: str, new_project_path: str) -> None:
    state_db = latest_state_db(codex_dir)
    with sqlite3.connect(state_db) as con:
        local_workspace_cwd = detect_workspace_cwd_representation(con, new_project_path)
        raw_project_path = canonical_path(new_project_path)
        cur = con.execute('SELECT id, cwd, sandbox_policy, agent_path FROM threads')
        rows = cur.fetchall()
        updated = 0
        for thread_id, cwd, sandbox_policy, agent_path in rows:
            if canonical_path(cwd) != canonical_path(old_project_path):
                continue
            con.execute(
                'UPDATE threads SET cwd = ?, sandbox_policy = ?, agent_path = ? WHERE id = ?',
                (
                    local_workspace_cwd,
                    rewrite_sandbox_policy(sandbox_policy or '', raw_project_path),
                    path_with_rebased_prefix(agent_path, old_project_path, raw_project_path),
                    thread_id,
                ),
            )
            updated += 1
        con.commit()
    print(f'Repaired thread rows: {updated}')
    print(f'State DB: {state_db}')


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description='Прод-синхронизация локальных Codex-чатов по конкретному проекту.')
    sub = parser.add_subparsers(dest='command', required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--sync-dir', required=True, help='Папка обмена через Яндекс.Диск / Google Drive / другую синхронизацию.')
    common.add_argument('--project-path', required=True, help='Путь к проекту на текущем устройстве.')
    common.add_argument('--codex-dir', default=None, help='Необязательно. По умолчанию берется ~/.codex')

    export_p = sub.add_parser('export', parents=[common])
    export_p.set_defaults(func=lambda a: export_threads(Path(a.sync_dir).expanduser().resolve(), detect_codex_dir(a.codex_dir), a.project_path))

    import_p = sub.add_parser('import', parents=[common])
    import_p.set_defaults(func=lambda a: import_threads(Path(a.sync_dir).expanduser().resolve(), detect_codex_dir(a.codex_dir), a.project_path))

    repair_p = sub.add_parser('repair-paths')
    repair_p.add_argument('--codex-dir', default=None)
    repair_p.add_argument('--old-project-path', required=True)
    repair_p.add_argument('--new-project-path', required=True)
    repair_p.set_defaults(func=lambda a: repair_paths(detect_codex_dir(a.codex_dir), a.old_project_path, a.new_project_path))

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
