from __future__ import annotations

from datetime import UTC, datetime
from enum import StrEnum
from pathlib import Path
from typing import Any
from uuid import uuid4

from pydantic import BaseModel, Field


class FileRole(StrEnum):
    MAIN = "основной"
    MAPPING = "маппинг"
    DIRECTORY = "справочник"
    TEMPLATE = "шаблон"
    OTHER = "прочее"


class Severity(StrEnum):
    ERROR = "ошибка"
    WARNING = "предупреждение"
    INFO = "информация"


class ImportDecision(BaseModel):
    severity: Severity
    code: str
    message: str
    context: dict[str, Any] = Field(default_factory=dict)


class PerformanceEntry(BaseModel):
    stage: str
    duration_seconds: float
    file_id: str | None = None
    file_name: str | None = None
    sheets_processed: int = 0
    rows_read: int = 0
    cells_read: int = 0
    cache_status: str | None = None
    parallel_tasks: int = 1
    details: dict[str, Any] = Field(default_factory=dict)


class UploadedFile(BaseModel):
    id: str = Field(default_factory=lambda: uuid4().hex)
    original_name: str
    stored_path: Path
    size_bytes: int
    extension: str
    role: FileRole | None = None
    uploaded_at: datetime = Field(default_factory=lambda: datetime.now(UTC))


class SheetProfile(BaseModel):
    name: str
    max_row: int
    max_column: int
    non_empty_cells: int
    formula_cells: int
    data_density: float
    scanned_rows: int = 0
    scanned_cells: int = 0
    scan_truncated: bool = False
    hidden: bool = False
    empty: bool = False
    candidate_for_processing: bool = False
    header_row: int | None = None
    headers: list[str] = Field(default_factory=list)
    formula_patterns: list[str] = Field(default_factory=list)
    table_ranges: list[str] = Field(default_factory=list)
    repeated_blocks: list[str] = Field(default_factory=list)


class FormulaSignature(BaseModel):
    sheet_name: str
    pattern: str
    count: int = 1


class SheetMeta(BaseModel):
    name: str
    max_row: int
    max_column: int
    hidden: bool = False
    empty: bool = False
    headers: list[str] = Field(default_factory=list)
    formula_signatures: list[FormulaSignature] = Field(default_factory=list)


class WorkbookFingerprint(BaseModel):
    sheet_count: int
    total_non_empty_cells: int
    total_formula_cells: int
    header_tokens: list[str] = Field(default_factory=list)
    formula_tokens: list[str] = Field(default_factory=list)
    structure_signature: str


class WorkbookMeta(BaseModel):
    file_id: str
    original_name: str
    path: Path
    sheets: list[SheetMeta]
    fingerprint: WorkbookFingerprint


class WorkbookProfile(BaseModel):
    file_id: str
    original_name: str
    path: Path
    sheets: list[SheetProfile]
    fingerprint: WorkbookFingerprint
    detected_role: FileRole
    role_confidence: float
    role_reasons: list[str] = Field(default_factory=list)
    decisions: list[ImportDecision] = Field(default_factory=list)
    performance: list[PerformanceEntry] = Field(default_factory=list)


class CompareReport(BaseModel):
    decisions: list[ImportDecision] = Field(default_factory=list)
    similar_pairs: list[dict[str, Any]] = Field(default_factory=list)
    sheet_differences: list[dict[str, Any]] = Field(default_factory=list)


class ColumnCandidate(BaseModel):
    source_file_id: str
    source_sheet: str
    source_header: str
    target_field: str
    score: float
    reason: str


class MappingRule(BaseModel):
    target_field: str
    source_file_id: str | None = None
    source_sheet: str | None = None
    source_header: str | None = None
    source_signature: str | None = None
    confirmed: bool = False


class ImportSession(BaseModel):
    id: str = Field(default_factory=lambda: uuid4().hex)
    created_at: datetime = Field(default_factory=lambda: datetime.now(UTC))
    files: list[UploadedFile] = Field(default_factory=list)
    profiles: list[WorkbookProfile] = Field(default_factory=list)
    mapping_rules: list[MappingRule] = Field(default_factory=list)
    decisions: list[ImportDecision] = Field(default_factory=list)
    performance: list[PerformanceEntry] = Field(default_factory=list)
    output_path: Path | None = None
    report_path: Path | None = None
