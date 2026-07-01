from __future__ import annotations

import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from time import time

from bdr_wizard.storage.repository import SessionRepository


RETENTION_DAYS_ENV = "BDR_RETENTION_DAYS"


@dataclass(frozen=True)
class CleanupResult:
    deleted_paths: int
    deleted_sessions: int


class CleanupService:
    def __init__(
        self,
        repository: SessionRepository,
        roots: tuple[Path, ...] = (
            Path("data/uploads"),
            Path("data/outputs"),
            Path("data/reports"),
            Path("data/cache/workbooks"),
        ),
    ) -> None:
        self.repository = repository
        self.roots = roots

    def run(self, retention_days: int | None = None) -> CleanupResult:
        days = retention_days or int(os.getenv(RETENTION_DAYS_ENV, "30"))
        cutoff_timestamp = time() - days * 24 * 60 * 60
        deleted_paths = 0
        for root in self.roots:
            deleted_paths += self._delete_old_children(root, cutoff_timestamp)
        deleted_sessions = self.repository.delete_older_than(days)
        return CleanupResult(deleted_paths=deleted_paths, deleted_sessions=deleted_sessions)

    def _delete_old_children(self, root: Path, cutoff_timestamp: float) -> int:
        if not root.exists():
            return 0
        deleted = 0
        for child in root.iterdir():
            if child.name == ".gitkeep":
                continue
            try:
                if child.stat().st_mtime >= cutoff_timestamp:
                    continue
                if child.is_dir():
                    shutil.rmtree(child)
                else:
                    child.unlink()
                deleted += 1
            except FileNotFoundError:
                continue
        return deleted
