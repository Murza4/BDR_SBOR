from __future__ import annotations

import os
import socket
import threading

import pytest
import uvicorn


pytest.importorskip("playwright.sync_api")
from playwright.sync_api import Error as PlaywrightError, sync_playwright  # noqa: E402


def test_upload_screen_opens_in_browser() -> None:
    os.environ["BDR_AUTH_DISABLED"] = "1"
    port = _free_port()
    server = uvicorn.Server(
        uvicorn.Config(
            "main:app",
            host="127.0.0.1",
            port=port,
            log_level="error",
        )
    )
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()

    try:
        with sync_playwright() as playwright:
            try:
                browser = playwright.chromium.launch()
            except PlaywrightError as exc:
                pytest.skip(f"Браузер Playwright не установлен: {exc}")
            page = browser.new_page()
            page.goto(f"http://127.0.0.1:{port}", wait_until="networkidle")
            assert "Загрузите до 7 Excel-файлов" in page.text_content("body")
            browser.close()
    finally:
        server.should_exit = True
        thread.join(timeout=5)


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])
