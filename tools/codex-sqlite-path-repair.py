from __future__ import annotations

import argparse
import datetime as dt
import pathlib
import sqlite3


OLD_PATH = r"E:\GitHub Projects\FillDoc"
NEW_PATH = r"C:\Projects\FillDoc"


def variants(path: str) -> list[str]:
    json_escaped = path.replace("\\", "\\\\")
    double_escaped = json_escaped.replace("\\", "\\\\")
    return [path, json_escaped, double_escaped]


def quote_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def iter_text_columns(con: sqlite3.Connection):
    tables = con.execute(
        "select name from sqlite_master where type = 'table' and name not like 'sqlite_%'"
    ).fetchall()
    for (table_name,) in tables:
        try:
            columns = con.execute(f"pragma table_info({quote_ident(table_name)})").fetchall()
        except sqlite3.DatabaseError:
            continue
        for column in columns:
            column_name = column[1]
            column_type = (column[2] or "").upper()
            if column_type == "" or "TEXT" in column_type or "CHAR" in column_type or "CLOB" in column_type:
                yield table_name, column_name


def backup_database(source: pathlib.Path) -> pathlib.Path:
    timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = source.with_name(f"{source.stem}.backup-filldoc-path-repair-{timestamp}{source.suffix}")
    source_con = sqlite3.connect(str(source), timeout=30)
    try:
        backup_con = sqlite3.connect(str(backup), timeout=30)
        try:
            source_con.backup(backup_con)
        finally:
            backup_con.close()
    finally:
        source_con.close()
    return backup


def find_matches(con: sqlite3.Connection):
    matches = []
    for table_name, column_name in iter_text_columns(con):
        table = quote_ident(table_name)
        column = quote_ident(column_name)
        for old, new in zip(variants(OLD_PATH), variants(NEW_PATH)):
            count_sql = (
                f"select count(*) from {table} "
                f"where typeof({column}) = 'text' and instr({column}, ?) > 0"
            )
            count = con.execute(count_sql, (old,)).fetchone()[0]
            if count:
                matches.append((table_name, column_name, old, new, count))
    return matches


def apply_matches(path: pathlib.Path, matches) -> int:
    con = sqlite3.connect(str(path), timeout=30)
    try:
        con.execute("pragma busy_timeout = 30000")
        before = con.total_changes
        for table_name, column_name, old, new, _count in matches:
            table = quote_ident(table_name)
            column = quote_ident(column_name)
            sql = (
                f"update {table} set {column} = replace({column}, ?, ?) "
                f"where typeof({column}) = 'text' and instr({column}, ?) > 0"
            )
            con.execute(sql, (old, new, old))
        con.commit()
        try:
            con.execute("pragma wal_checkpoint(full)")
        except sqlite3.DatabaseError:
            pass
        return con.total_changes - before
    finally:
        con.close()


def repair_database(path: pathlib.Path, dry_run: bool) -> int:
    if not path.exists():
        raise FileNotFoundError(path)

    con = sqlite3.connect(str(path), timeout=30)
    try:
        con.execute("pragma busy_timeout = 30000")
        matches = find_matches(con)
    finally:
        con.close()

    if dry_run:
        for table_name, column_name, _old, _new, count in matches:
            print(f"would repair {count} rows in {table_name}.{column_name}")
        return 0

    if not matches:
        return 0

    backup = backup_database(path)
    changed = apply_matches(path, matches)
    if backup:
        print(f"backup: {backup}")
    return changed


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--db",
        default=str(pathlib.Path.home() / ".codex" / "state_5.sqlite"),
        help="Path to Codex state_5.sqlite",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    changed = repair_database(pathlib.Path(args.db), args.dry_run)
    print(f"sqlite rows changed: {changed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
