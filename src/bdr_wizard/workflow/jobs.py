from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from threading import Lock

from bdr_wizard.ingestion import IngestionService
from bdr_wizard.models import ImportSession
from bdr_wizard.workflow.wizard import WizardWorkflow


JOB_WORKERS_ENV = "BDR_JOB_WORKERS"


@dataclass
class JobStatus:
    state: str
    error: str | None = None


class WizardJobQueue:
    def __init__(
        self,
        workflow: WizardWorkflow,
        ingestion: IngestionService,
        max_workers: int | None = None,
    ) -> None:
        self.workflow = workflow
        self.ingestion = ingestion
        workers = max_workers or int(os.getenv(JOB_WORKERS_ENV, "2"))
        self.executor = ThreadPoolExecutor(max_workers=max(workers, 1), thread_name_prefix="bdr-job")
        self.statuses: dict[str, JobStatus] = {}
        self.lock = Lock()

    def enqueue_recognition(self, session: ImportSession) -> None:
        with self.lock:
            self.statuses[session.id] = JobStatus(state="в обработке")
        self.executor.submit(self._run_recognition, session)

    def status(self, session: ImportSession) -> JobStatus:
        with self.lock:
            status = self.statuses.get(session.id)
        if status is not None:
            return status
        if session.profiles:
            return JobStatus(state="готово")
        return JobStatus(state="в обработке")

    def _run_recognition(self, session: ImportSession) -> None:
        try:
            self.workflow.recognize(session)
        except Exception as exc:
            self.ingestion.discard_session_files(session)
            with self.lock:
                self.statuses[session.id] = JobStatus(state="ошибка", error=str(exc))
            return
        with self.lock:
            self.statuses[session.id] = JobStatus(state="готово")
