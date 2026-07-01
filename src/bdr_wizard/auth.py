from __future__ import annotations

import base64
import hashlib
import hmac
import os
from dataclasses import dataclass
from secrets import compare_digest, token_urlsafe
from time import time


AUTH_COOKIE_NAME = "bdr_auth"
AUTH_MAX_AGE_SECONDS = 12 * 60 * 60
AUTH_DISABLED_ENV = "BDR_AUTH_DISABLED"
ADMIN_USER_ENV = "BDR_ADMIN_USER"
ADMIN_PASSWORD_ENV = "BDR_ADMIN_PASSWORD"
AUTH_SECRET_ENV = "BDR_AUTH_SECRET"


@dataclass(frozen=True)
class AuthSettings:
    disabled: bool
    username: str
    password: str
    secret: str


def load_auth_settings() -> AuthSettings:
    username = os.getenv(ADMIN_USER_ENV, "admin")
    password = os.getenv(ADMIN_PASSWORD_ENV, "admin")
    secret = os.getenv(AUTH_SECRET_ENV) or password or token_urlsafe(32)
    return AuthSettings(
        disabled=os.getenv(AUTH_DISABLED_ENV) == "1",
        username=username,
        password=password,
        secret=secret,
    )


def verify_credentials(username: str, password: str, settings: AuthSettings | None = None) -> bool:
    settings = settings or load_auth_settings()
    return compare_digest(username, settings.username) and compare_digest(password, settings.password)


def create_auth_token(username: str, settings: AuthSettings | None = None) -> str:
    settings = settings or load_auth_settings()
    issued_at = str(int(time()))
    payload = f"{username}:{issued_at}"
    signature = _sign(payload, settings.secret)
    token = f"{payload}:{signature}".encode("utf-8")
    return base64.urlsafe_b64encode(token).decode("ascii")


def verify_auth_token(token: str | None, settings: AuthSettings | None = None) -> bool:
    if not token:
        return False
    settings = settings or load_auth_settings()
    try:
        decoded = base64.urlsafe_b64decode(token.encode("ascii")).decode("utf-8")
        username, issued_at_text, signature = decoded.rsplit(":", 2)
        issued_at = int(issued_at_text)
    except (ValueError, UnicodeDecodeError):
        return False
    if username != settings.username:
        return False
    if int(time()) - issued_at > AUTH_MAX_AGE_SECONDS:
        return False
    expected_signature = _sign(f"{username}:{issued_at_text}", settings.secret)
    return compare_digest(signature, expected_signature)


def _sign(payload: str, secret: str) -> str:
    digest = hmac.new(secret.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).digest()
    return base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")
