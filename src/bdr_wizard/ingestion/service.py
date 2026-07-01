from __future__ import annotations

import shutil
from pathlib import Path
from uuid import uuid4

from bdr_wizard.models import ImportSession, UploadedFile
from bdr_wizard.security import (
    WorkbookSecurityError,
    validate_file_size,
    validate_session_upload_size,
)
from bdr_wizard.storage.encryption import EncryptedFileStore


ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


class UnsupportedFileError(ValueError):
    pass


class FileTooLargeError(UnsupportedFileError):
    pass


class IngestionService:
    def __init__(
        self,
        upload_root: Path = Path("data/uploads"),
        encrypted_store: EncryptedFileStore | None = None,
    ) -> None:
        self.upload_root = upload_root
        self.encrypted_store = encrypted_store or EncryptedFileStore()
        self.upload_root.mkdir(parents=True, exist_ok=True)

    def create_session(self) -> ImportSession:
        return ImportSession()

    def discard_session_files(self, session: ImportSession) -> None:
        shutil.rmtree(self.upload_root / session.id, ignore_errors=True)

    def add_file(
        self,
        session: ImportSession,
        source_path: Path,
        original_name: str | None = None,
    ) -> UploadedFile:
        extension = source_path.suffix.lower()
        if extension not in ALLOWED_EXTENSIONS:
            raise UnsupportedFileError(
                "Поддерживаются только файлы Excel в формате .xlsx или .xlsm."
            )
        try:
            size_bytes = source_path.stat().st_size
            validate_file_size(size_bytes)
            validate_session_upload_size(self._session_size(session), size_bytes)
        except WorkbookSecurityError as exc:
            raise FileTooLargeError(str(exc)) from exc

        session_dir = self.upload_root / session.id
        session_dir.mkdir(parents=True, exist_ok=True)
        file_id = uuid4().hex
        safe_name = original_name or source_path.name
        stored_path = self.encrypted_store.write_encrypted(
            session_dir / f"{file_id}{extension}",
            source_path.read_bytes(),
        )

        uploaded = UploadedFile(
            id=file_id,
            original_name=safe_name,
            stored_path=stored_path,
            size_bytes=stored_path.stat().st_size,
            extension=extension,
        )
        session.files.append(uploaded)
        return uploaded

    def add_uploaded_bytes(
        self,
        session: ImportSession,
        original_name: str,
        content: bytes,
    ) -> UploadedFile:
        extension = Path(original_name).suffix.lower()
        if extension not in ALLOWED_EXTENSIONS:
            raise UnsupportedFileError(
                "Поддерживаются только файлы Excel в формате .xlsx или .xlsm."
            )
        try:
            validate_file_size(len(content))
            validate_session_upload_size(self._session_size(session), len(content))
        except WorkbookSecurityError as exc:
            raise FileTooLargeError(str(exc)) from exc
        session_dir = self.upload_root / session.id
        session_dir.mkdir(parents=True, exist_ok=True)
        file_id = uuid4().hex
        stored_path = self.encrypted_store.write_encrypted(
            session_dir / f"{file_id}{extension}",
            content,
        )
        uploaded = UploadedFile(
            id=file_id,
            original_name=original_name,
            stored_path=stored_path,
            size_bytes=len(content),
            extension=extension,
        )
        session.files.append(uploaded)
        return uploaded

    def _session_size(self, session: ImportSession) -> int:
        return sum(uploaded_file.size_bytes for uploaded_file in session.files)
