from __future__ import annotations

import json
from pathlib import Path

from bdr_wizard.models import ImportSession
from bdr_wizard.storage.encryption import EncryptedFileStore


class ReportingService:
    def __init__(
        self,
        report_root: Path = Path("data/reports"),
        encrypted_store: EncryptedFileStore | None = None,
    ) -> None:
        self.report_root = report_root
        self.encrypted_store = encrypted_store or EncryptedFileStore()
        self.report_root.mkdir(parents=True, exist_ok=True)

    def write_audit_report(self, session: ImportSession) -> Path:
        report_path = self.report_root / session.id / "audit_report.json"
        report_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "идентификатор_сессии": session.id,
            "создано": session.created_at.isoformat(),
            "файлы": [file.model_dump(mode="json") for file in session.files],
            "профили_workbook": [profile.model_dump(mode="json") for profile in session.profiles],
            "правила_маппинга": [rule.model_dump(mode="json") for rule in session.mapping_rules],
            "решения": [decision.model_dump(mode="json") for decision in session.decisions],
            "производительность": [
                entry.model_dump(mode="json") for entry in session.performance
            ],
            "путь_к_итоговому_файлу": str(session.output_path) if session.output_path else None,
        }
        session.report_path = self.encrypted_store.write_encrypted(
            report_path,
            json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
        )
        return session.report_path
