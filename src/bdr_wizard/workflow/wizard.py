from __future__ import annotations

from concurrent.futures import ProcessPoolExecutor
from time import perf_counter

from bdr_wizard.excel.analyzer import WorkbookAnalyzer
from bdr_wizard.excel.compare import CompareEngine
from bdr_wizard.ingestion import IngestionService
from bdr_wizard.mapper import MappingEngine
from bdr_wizard.models import ImportSession, PerformanceEntry, UploadedFile, WorkbookProfile
from bdr_wizard.storage import SessionRepository


def analyze_uploaded_file(uploaded_file: UploadedFile) -> tuple[WorkbookProfile | None, str | None]:
    try:
        return WorkbookAnalyzer().analyze(uploaded_file), None
    except Exception as exc:
        return None, str(exc)


class WizardWorkflow:
    def __init__(
        self,
        ingestion: IngestionService,
        mapper: MappingEngine,
        repository: SessionRepository,
        compare_engine: CompareEngine | None = None,
    ) -> None:
        self.ingestion = ingestion
        self.mapper = mapper
        self.repository = repository
        self.compare_engine = compare_engine or CompareEngine()

    def create_upload_session(self, files: list[tuple[str, bytes]]) -> ImportSession:
        session = self.ingestion.create_session()
        for filename, content in files:
            self.ingestion.add_uploaded_bytes(session, filename, content)
        return session

    def recognize(self, session: ImportSession) -> ImportSession:
        analysis_started_at = perf_counter()
        parallel_tasks = min(len(session.files), 4) or 1
        with ProcessPoolExecutor(max_workers=parallel_tasks) as executor:
            analysis_results = list(executor.map(analyze_uploaded_file, session.files))

        analysis_errors = [error for _, error in analysis_results if error]
        if analysis_errors:
            raise RuntimeError(" ".join(analysis_errors))

        profiles = [profile for profile, _ in analysis_results if profile is not None]
        for profile in profiles:
            for entry in profile.performance:
                entry.parallel_tasks = parallel_tasks
            session.profiles.append(profile)
            session.decisions.extend(profile.decisions)
            session.performance.extend(profile.performance)

        session.performance.append(
            PerformanceEntry(
                stage="parallel_file_analysis",
                duration_seconds=round(perf_counter() - analysis_started_at, 6),
                sheets_processed=sum(len(profile.sheets) for profile in profiles),
                rows_read=sum(entry.rows_read for profile in profiles for entry in profile.performance),
                cells_read=sum(entry.cells_read for profile in profiles for entry in profile.performance),
                parallel_tasks=parallel_tasks,
            )
        )
        compare_report = self.compare_engine.compare(session.profiles)
        session.decisions.extend(compare_report.decisions)
        suggested_rules, _ = self.mapper.suggest(session.profiles)
        session.mapping_rules = suggested_rules
        self.repository.save(session)
        return session
