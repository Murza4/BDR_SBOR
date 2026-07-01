from __future__ import annotations

import os
from pathlib import Path
from secrets import token_urlsafe

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, PlainTextResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.background import BackgroundTask
from starlette.exceptions import HTTPException as StarletteHTTPException

from bdr_wizard.auth import (
    AUTH_COOKIE_NAME,
    AUTH_MAX_AGE_SECONDS,
    create_auth_token,
    load_auth_settings,
    verify_auth_token,
    verify_credentials,
)
from bdr_wizard.builder import BdrBuilder
from bdr_wizard.excel.loader import WorkbookLoader
from bdr_wizard.ingestion import IngestionService, UnsupportedFileError
from bdr_wizard.localization import ROLE_LABELS, TARGET_FIELDS
from bdr_wizard.mapper import MappingEngine
from bdr_wizard.models import FileRole, ImportDecision, MappingRule, Severity
from bdr_wizard.reporting import ReportingService
from bdr_wizard.security import MAX_UPLOAD_SIZE_BYTES, format_bytes, validate_upload_file_count
from bdr_wizard.storage import CleanupService, SessionRepository
from bdr_wizard.storage.encryption import ENCRYPTED_SUFFIX, EncryptedFileStore
from bdr_wizard.workflow import WizardJobQueue, WizardWorkflow


BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

app = FastAPI(title="Мастер сборки БДР", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
UPLOAD_CHUNK_SIZE = 1024 * 1024
CSRF_COOKIE_NAME = "bdr_csrf_token"
LOCAL_CLIENT_HOSTS = {"127.0.0.1", "::1", "localhost", "testclient"}

repository = SessionRepository()
cleanup = CleanupService(repository)
encrypted_store = EncryptedFileStore()
ingestion = IngestionService(encrypted_store=encrypted_store)
mapping_engine = MappingEngine()
builder = BdrBuilder(loader=WorkbookLoader(encrypted_store=encrypted_store))
reporting = ReportingService(encrypted_store=encrypted_store)
wizard = WizardWorkflow(ingestion, mapping_engine, repository)
job_queue = WizardJobQueue(wizard, ingestion)
auth_settings = load_auth_settings()
PUBLIC_PATHS = {"/login"}


@app.middleware("http")
async def local_only_middleware(request: Request, call_next):
    if os.getenv("BDR_ALLOW_REMOTE") == "1":
        return await call_next(request)
    client_host = request.client.host if request.client else ""
    if client_host and client_host not in LOCAL_CLIENT_HOSTS:
        return PlainTextResponse(
            "Доступ разрешен только с локального компьютера.",
            status_code=403,
        )
    return await call_next(request)


@app.on_event("startup")
def cleanup_old_artifacts() -> None:
    cleanup.run()


@app.middleware("http")
async def csrf_cookie_middleware(request: Request, call_next):
    token = request.cookies.get(CSRF_COOKIE_NAME) or token_urlsafe(32)
    request.state.csrf_token = token
    response = await call_next(request)
    if request.method == "GET" and request.cookies.get(CSRF_COOKIE_NAME) is None:
        response.set_cookie(
            CSRF_COOKIE_NAME,
            token,
            httponly=True,
            samesite="lax",
        )
    return response


@app.middleware("http")
async def authentication_middleware(request: Request, call_next):
    if auth_settings.disabled:
        request.state.authenticated = True
        return await call_next(request)
    path = request.url.path
    if path.startswith("/static") or path in PUBLIC_PATHS:
        return await call_next(request)
    if verify_auth_token(request.cookies.get(AUTH_COOKIE_NAME), auth_settings):
        request.state.authenticated = True
        return await call_next(request)
    return RedirectResponse("/login", status_code=303)


@app.exception_handler(StarletteHTTPException)
async def http_error_page(request: Request, exc: StarletteHTTPException):
    message = str(exc.detail) if exc.detail else "Страница не найдена."
    return templates.TemplateResponse(
        request,
        "error.html",
        {"title": "Ошибка", "message": message},
        status_code=exc.status_code,
    )


@app.exception_handler(RuntimeError)
async def runtime_error_page(request: Request, exc: RuntimeError):
    return templates.TemplateResponse(
        request,
        "error.html",
        {"title": "Ошибка", "message": str(exc)},
        status_code=400,
    )


@app.exception_handler(RequestValidationError)
async def validation_error_page(request: Request, exc: RequestValidationError):
    if request.url.path == "/sessions":
        return templates.TemplateResponse(
            request,
            "upload.html",
            {
                "title": "Загрузка Excel-файлов",
                "error": "Выберите от 1 до 7 Excel-файлов перед началом распознавания.",
                "csrf_token": _csrf_token(request),
            },
            status_code=400,
        )
    return templates.TemplateResponse(
        request,
        "error.html",
        {
            "title": "Ошибка",
            "message": "Проверьте заполнение формы и повторите действие.",
        },
        status_code=400,
    )


@app.get("/login")
def login_page(request: Request):
    if not auth_settings.disabled and verify_auth_token(request.cookies.get(AUTH_COOKIE_NAME), auth_settings):
        return RedirectResponse("/", status_code=303)
    return templates.TemplateResponse(
        request,
        "login.html",
        {
            "title": "Вход",
            "error": None,
            "csrf_token": _csrf_token(request),
        },
    )


@app.post("/login")
async def login(request: Request, username: str = Form(...), password: str = Form(...), csrf_token: str = Form(...)):
    _require_csrf(request, csrf_token)
    if not verify_credentials(username, password, auth_settings):
        return templates.TemplateResponse(
            request,
            "login.html",
            {
                "title": "Вход",
                "error": "Неверный логин или пароль.",
                "csrf_token": _csrf_token(request),
            },
            status_code=400,
        )
    response = RedirectResponse("/", status_code=303)
    response.set_cookie(
        AUTH_COOKIE_NAME,
        create_auth_token(username, auth_settings),
        max_age=AUTH_MAX_AGE_SECONDS,
        httponly=True,
        samesite="lax",
    )
    return response


@app.post("/logout")
async def logout(request: Request):
    form = await request.form()
    _require_csrf(request, str(form.get("csrf_token") or ""))
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(AUTH_COOKIE_NAME)
    return response


@app.get("/")
def index(request: Request):
    return templates.TemplateResponse(
        request,
        "upload.html",
        {
            "title": "Загрузка Excel-файлов",
            "error": None,
            "csrf_token": _csrf_token(request),
        },
    )


@app.post("/sessions")
async def create_session(
    request: Request,
    csrf_token: str = Form(...),
    files: list[UploadFile] = File(...),
):
    _require_csrf(request, csrf_token)
    file_count_error = validate_upload_file_count(len(files))
    if file_count_error:
        return templates.TemplateResponse(
            request,
            "upload.html",
            {
                "title": "Загрузка Excel-файлов",
                "error": file_count_error,
                "csrf_token": _csrf_token(request),
            },
            status_code=400,
        )

    session = ingestion.create_session()
    try:
        for uploaded in files:
            content = await _read_upload_with_limit(uploaded)
            ingestion.add_uploaded_bytes(session, uploaded.filename or "файл.xlsx", content)
    except UnsupportedFileError as exc:
        ingestion.discard_session_files(session)
        return templates.TemplateResponse(
            request,
            "upload.html",
            {
                "title": "Загрузка Excel-файлов",
                "error": str(exc),
                "csrf_token": _csrf_token(request),
            },
            status_code=400,
        )

    repository.save(session)
    job_queue.enqueue_recognition(session)
    return RedirectResponse(f"/sessions/{session.id}/processing", status_code=303)


@app.get("/sessions/{session_id}/processing")
def processing(request: Request, session_id: str):
    session = _require_session(session_id)
    status = job_queue.status(session)
    if status.state == "готово":
        return RedirectResponse(f"/sessions/{session.id}/preview", status_code=303)
    return templates.TemplateResponse(
        request,
        "processing.html",
        {
            "title": "Обработка файлов",
            "session": session,
            "status": status,
        },
        status_code=400 if status.state == "ошибка" else 200,
    )


@app.get("/sessions/{session_id}/preview")
def preview(request: Request, session_id: str):
    session = _require_session(session_id)
    if not session.profiles:
        return RedirectResponse(f"/sessions/{session.id}/processing", status_code=303)
    return templates.TemplateResponse(
        request,
        "preview.html",
        {
            "title": "Предпросмотр распознавания",
            "session": session,
            "role_labels": ROLE_LABELS,
            "roles": list(FileRole),
            "csrf_token": _csrf_token(request),
        },
    )


@app.post("/sessions/{session_id}/preview")
async def update_roles(request: Request, session_id: str):
    session = _require_session(session_id)
    form = await request.form()
    _require_csrf(request, str(form.get("csrf_token") or ""))
    for profile in session.profiles:
        selected = form.get(f"role_{profile.file_id}")
        if selected:
            try:
                role = FileRole(str(selected))
            except ValueError:
                raise RuntimeError("Выбрана неизвестная роль файла. Обновите страницу и повторите действие.") from None
            profile.detected_role = role
            for uploaded_file in session.files:
                if uploaded_file.id == profile.file_id:
                    uploaded_file.role = role
    repository.save(session)
    return RedirectResponse(f"/sessions/{session.id}/mapping", status_code=303)


@app.get("/sessions/{session_id}/mapping")
def mapping(request: Request, session_id: str):
    session = _require_session(session_id)
    suggested_rules, candidates = mapping_engine.suggest(session.profiles)
    if not session.mapping_rules:
        session.mapping_rules = suggested_rules
        repository.save(session)
    choices_by_field = {
        field: [candidate for candidate in candidates if candidate.target_field == field]
        for field in TARGET_FIELDS
    }
    return templates.TemplateResponse(
        request,
        "mapping.html",
        {
            "title": "Сопоставление полей",
            "session": session,
            "target_fields": list(TARGET_FIELDS.keys()),
            "choices_by_field": choices_by_field,
            "profiles_by_id": {profile.file_id: profile for profile in session.profiles},
            "csrf_token": _csrf_token(request),
        },
    )


@app.post("/sessions/{session_id}/mapping")
async def update_mapping(request: Request, session_id: str):
    session = _require_session(session_id)
    form = await request.form()
    _require_csrf(request, str(form.get("csrf_token") or ""))
    _, candidates = mapping_engine.suggest(session.profiles)
    signatures_by_choice = {
        f"{candidate.source_file_id}|||{candidate.source_sheet}|||{candidate.source_header}": mapping_engine.source_signature(
            session.profiles,
            candidate.source_file_id,
            candidate.source_sheet,
            candidate.source_header,
        )
        for candidate in candidates
    }
    allowed_choices = set(signatures_by_choice)
    rules: list[MappingRule] = []
    for index, target_field in enumerate(TARGET_FIELDS.keys()):
        selected = str(form.get(f"mapping_{index}") or "")
        if not selected:
            rules.append(MappingRule(target_field=target_field, confirmed=False))
            continue
        if selected not in allowed_choices:
            raise RuntimeError(
                "Форма сопоставления содержит неизвестный источник. "
                "Обновите страницу и выберите значение из списка."
            )
        source_file_id, source_sheet, source_header = selected.split("|||", 2)
        rules.append(
            MappingRule(
                target_field=target_field,
                source_file_id=source_file_id,
                source_sheet=source_sheet,
                source_header=source_header,
                source_signature=signatures_by_choice.get(selected),
                confirmed=True,
            )
        )
    session.mapping_rules = rules
    mapping_engine.rule_store.save(rules)
    repository.save(session)
    return RedirectResponse(f"/sessions/{session.id}/build", status_code=303)


@app.get("/sessions/{session_id}/build")
def build_result(request: Request, session_id: str):
    session = _require_session(session_id)
    if session.output_path is None or not session.output_path.exists():
        if not any(rule.confirmed for rule in session.mapping_rules):
            session.decisions.append(
                ImportDecision(
                    severity=Severity.WARNING,
                    code="маппинг_не_подтвержден",
                    message="Пользователь не подтвердил ни одного правила сопоставления.",
                )
            )
        builder.build(session)
        reporting.write_audit_report(session)
        repository.save(session)
    return templates.TemplateResponse(
        request,
        "result.html",
        {
            "title": "Результат сборки",
            "session": session,
        },
    )


@app.get("/sessions/{session_id}/download/result")
def download_result(session_id: str):
    session = _require_session(session_id)
    if session.output_path is None:
        raise RuntimeError("Итоговый файл еще не создан.")
    return _download_file(
        session.output_path,
        filename="itogovyj_bdr.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/sessions/{session_id}/download/report")
def download_report(session_id: str):
    session = _require_session(session_id)
    if session.report_path is None:
        raise RuntimeError("Отчет еще не создан.")
    return _download_file(
        session.report_path,
        filename="otchet_ob_obrabotke.json",
        media_type="application/json",
    )


def _require_session(session_id: str):
    session = repository.get(session_id)
    if session is None:
        raise RuntimeError("Сессия импорта не найдена.")
    return session


def _download_file(path: Path, filename: str, media_type: str) -> FileResponse:
    if path.suffix != ENCRYPTED_SUFFIX:
        return FileResponse(path, filename=filename, media_type=media_type)
    suffix = Path(filename).suffix
    with encrypted_store.materialized_file(path, suffix=suffix) as temporary_path:
        response_path = Path(f"{temporary_path}.download")
        temporary_path.replace(response_path)
    return FileResponse(
        response_path,
        filename=filename,
        media_type=media_type,
        background=BackgroundTask(response_path.unlink, missing_ok=True),
    )


def _csrf_token(request: Request) -> str:
    return str(getattr(request.state, "csrf_token", ""))


def _require_csrf(request: Request, submitted_token: str) -> None:
    cookie_token = request.cookies.get(CSRF_COOKIE_NAME)
    if not cookie_token or not submitted_token or submitted_token != cookie_token:
        raise RuntimeError("Срок действия формы истек. Обновите страницу и повторите действие.")


async def _read_upload_with_limit(uploaded: UploadFile) -> bytes:
    total_size = 0
    chunks: list[bytes] = []
    while True:
        chunk = await uploaded.read(UPLOAD_CHUNK_SIZE)
        if not chunk:
            break
        total_size += len(chunk)
        if total_size > MAX_UPLOAD_SIZE_BYTES:
            raise UnsupportedFileError(
                "Файл слишком большой: "
                f"максимальный размер {format_bytes(MAX_UPLOAD_SIZE_BYTES)}."
            )
        chunks.append(chunk)
    return b"".join(chunks)
