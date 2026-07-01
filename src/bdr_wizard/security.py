from __future__ import annotations

from pathlib import Path
from zipfile import BadZipFile, ZipFile


MIN_UPLOAD_FILES = 1
MAX_UPLOAD_FILES = 7
MAX_UPLOAD_SIZE_BYTES = 50 * 1024 * 1024
MAX_SESSION_UPLOAD_SIZE_BYTES = 200 * 1024 * 1024
MAX_WORKBOOK_ARCHIVE_ENTRIES = 10_000
MAX_WORKBOOK_UNCOMPRESSED_BYTES = 250 * 1024 * 1024
MAX_WORKBOOK_COMPRESSION_RATIO = 100
MAX_WORKBOOK_SHEETS = 100
MAX_BUILD_ROWS = 100_000

DANGEROUS_EXCEL_PREFIXES = ("=", "+", "-", "@")


class WorkbookSecurityError(ValueError):
    pass


def format_bytes(size: int) -> str:
    if size >= 1024 * 1024:
        return f"{size / (1024 * 1024):.1f} МБ"
    if size >= 1024:
        return f"{size / 1024:.1f} КБ"
    return f"{size} байт"


def validate_file_size(size_bytes: int) -> None:
    if size_bytes > MAX_UPLOAD_SIZE_BYTES:
        raise WorkbookSecurityError(
            "Файл слишком большой: "
            f"{format_bytes(size_bytes)}. Максимальный размер: {format_bytes(MAX_UPLOAD_SIZE_BYTES)}."
        )


def validate_session_upload_size(current_size: int, next_size: int) -> None:
    total_size = current_size + next_size
    if total_size > MAX_SESSION_UPLOAD_SIZE_BYTES:
        raise WorkbookSecurityError(
            "Суммарный размер файлов в сессии слишком большой: "
            f"{format_bytes(total_size)}. Максимум: {format_bytes(MAX_SESSION_UPLOAD_SIZE_BYTES)}."
        )


def validate_upload_file_count(file_count: int) -> str | None:
    if file_count < MIN_UPLOAD_FILES:
        return "Выберите хотя бы один Excel-файл."
    if file_count > MAX_UPLOAD_FILES:
        return f"Можно загрузить не больше {MAX_UPLOAD_FILES} Excel-файлов за одну сессию."
    return None


def validate_excel_archive(path: Path) -> None:
    validate_file_size(path.stat().st_size)
    try:
        with ZipFile(path) as archive:
            entries = archive.infolist()
            if len(entries) > MAX_WORKBOOK_ARCHIVE_ENTRIES:
                raise WorkbookSecurityError(
                    "Excel-файл содержит слишком много внутренних элементов. "
                    "Возможна поврежденная или небезопасная книга."
                )

            compressed_size = sum(entry.compress_size for entry in entries)
            uncompressed_size = sum(entry.file_size for entry in entries)
            if uncompressed_size > MAX_WORKBOOK_UNCOMPRESSED_BYTES:
                raise WorkbookSecurityError(
                    "Excel-файл слишком велик после распаковки: "
                    f"{format_bytes(uncompressed_size)}. Максимум: "
                    f"{format_bytes(MAX_WORKBOOK_UNCOMPRESSED_BYTES)}."
                )
            if compressed_size and uncompressed_size / compressed_size > MAX_WORKBOOK_COMPRESSION_RATIO:
                raise WorkbookSecurityError(
                    "Excel-файл имеет подозрительно высокий коэффициент сжатия. "
                    "Обработка остановлена для защиты от перегрузки."
                )
    except BadZipFile as exc:
        raise WorkbookSecurityError(
            "Файл не похож на корректную Excel-книгу .xlsx или .xlsm."
        ) from exc


def neutralize_excel_formula(value: object) -> object:
    if not isinstance(value, str):
        return value
    if not value:
        return value
    first_character = value[0]
    if first_character in DANGEROUS_EXCEL_PREFIXES or first_character in {"\t", "\r", "\n"}:
        return f"'{value}"
    return value
