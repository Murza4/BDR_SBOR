from __future__ import annotations

import os
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

from cryptography.fernet import Fernet


ENCRYPTION_KEY_ENV = "BDR_ENCRYPTION_KEY"
DEFAULT_KEY_PATH = Path("data/bdr_encryption.key")
ENCRYPTED_SUFFIX = ".enc"


class EncryptedFileStore:
    def __init__(self, key_path: Path = DEFAULT_KEY_PATH) -> None:
        self.key_path = key_path
        self.fernet = Fernet(self._load_or_create_key())

    def encrypted_path(self, path: Path) -> Path:
        if path.suffix == ENCRYPTED_SUFFIX:
            return path
        return path.with_name(f"{path.name}{ENCRYPTED_SUFFIX}")

    def write_encrypted(self, path: Path, content: bytes) -> Path:
        encrypted_path = self.encrypted_path(path)
        encrypted_path.parent.mkdir(parents=True, exist_ok=True)
        encrypted_path.write_bytes(self.fernet.encrypt(content))
        return encrypted_path

    def read_bytes(self, path: Path) -> bytes:
        if path.suffix == ENCRYPTED_SUFFIX:
            return self.fernet.decrypt(path.read_bytes())
        return path.read_bytes()

    @contextmanager
    def materialized_file(self, path: Path, suffix: str = "") -> Iterator[Path]:
        if path.suffix != ENCRYPTED_SUFFIX:
            yield path
            return
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as handle:
            temporary_path = Path(handle.name)
            handle.write(self.read_bytes(path))
        try:
            yield temporary_path
        finally:
            temporary_path.unlink(missing_ok=True)

    def _load_or_create_key(self) -> bytes:
        env_key = os.getenv(ENCRYPTION_KEY_ENV)
        if env_key:
            return env_key.encode("utf-8")
        self.key_path.parent.mkdir(parents=True, exist_ok=True)
        if self.key_path.exists():
            return self.key_path.read_bytes().strip()
        key = Fernet.generate_key()
        self.key_path.write_bytes(key)
        return key
