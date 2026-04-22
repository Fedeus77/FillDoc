"""Тесты ротации резервных копий Excel."""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from filldoc.excel.excel_store import ExcelProjectStore


class TestRotateBackups:
    def _make_store(self, excel_path: str) -> ExcelProjectStore:
        return ExcelProjectStore(excel_path)

    def _touch_backup(self, backup_dir: Path, stem: str, suffix: str, age_days: float = 0) -> Path:
        ts = (dt.datetime.now() - dt.timedelta(days=age_days)).strftime("%Y%m%d_%H%M%S")
        p = backup_dir / f"{stem}__backup__{ts}{suffix}"
        p.write_text("x")
        # Сдвигаем mtime для имитации возраста
        mtime = (dt.datetime.now() - dt.timedelta(days=age_days)).timestamp()
        import os
        os.utime(p, (mtime, mtime))
        return p

    def test_removes_old_backups(self, tmp_path: Path) -> None:
        """Бэкапы старше 15 дней должны удаляться."""
        backup_dir = tmp_path / "_filldoc_backups"
        backup_dir.mkdir()
        store = self._make_store(str(tmp_path / "fake.xlsx"))

        old = self._touch_backup(backup_dir, "fake", ".xlsx", age_days=20)
        fresh = self._touch_backup(backup_dir, "fake", ".xlsx", age_days=1)

        store._rotate_backups(backup_dir, "fake", ".xlsx")

        assert not old.exists(), "Старый бэкап должен быть удалён"
        assert fresh.exists(), "Свежий бэкап должен остаться"

    def test_keeps_at_most_max_count(self, tmp_path: Path) -> None:
        """После ротации остаётся не больше _BACKUP_MAX_COUNT штук."""
        backup_dir = tmp_path / "_filldoc_backups"
        backup_dir.mkdir()
        store = self._make_store(str(tmp_path / "fake.xlsx"))

        import time
        for i in range(25):
            time.sleep(0.01)  # разные mtime
            self._touch_backup(backup_dir, "fake", ".xlsx", age_days=0)

        store._rotate_backups(backup_dir, "fake", ".xlsx")

        remaining = list(backup_dir.glob("fake__backup__*.xlsx"))
        assert len(remaining) <= ExcelProjectStore._BACKUP_MAX_COUNT
