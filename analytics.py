from __future__ import annotations

import hashlib
import json
import logging
import os
import secrets
from datetime import datetime, timezone
from typing import Any

from dotenv import load_dotenv
from flask import Request, has_request_context, request
from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    Integer,
    MetaData,
    String,
    Table,
    Text,
    create_engine,
    func,
    select,
)
from sqlalchemy.engine import Engine
from sqlalchemy.exc import OperationalError


load_dotenv()

logger = logging.getLogger(__name__)

metadata = MetaData()

analytics_events = Table(
    "analytics_events",
    metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("created_at", DateTime(timezone=True), nullable=False, default=lambda: datetime.now(timezone.utc)),
    Column("event_name", String(80), nullable=False, index=True),
    Column("session_id", String(80), nullable=True, index=True),
    Column("success", Boolean, nullable=True, index=True),
    Column("page_path", String(500), nullable=True),
    Column("card_count", Integer, nullable=True),
    Column("row_count", Integer, nullable=True),
    Column("error_message", String(500), nullable=True),
    Column("user_agent", String(500), nullable=True),
    Column("ip_hash", String(64), nullable=True),
    Column("meta_json", Text, nullable=True),
)

_engine: Engine | None = None


def _database_url() -> str:
    url = os.environ.get("DATABASE_URL") or "sqlite:///analytics_local.db"
    if url.startswith("postgres://"):
        return "postgresql://" + url[len("postgres://") :]
    return url


def init_analytics() -> Engine | None:
    """Initialize analytics storage. Failures are logged and never raised."""
    global _engine
    if _engine is not None:
        return _engine

    try:
        _engine = create_engine(_database_url(), pool_pre_ping=True, future=True)
        metadata.create_all(_engine)
        return _engine
    except OperationalError as exc:
        if "already exists" in str(exc).lower():
            return _engine
        logger.warning("analytics initialization failed: %s", type(exc).__name__)
        _engine = None
        return None
    except Exception as exc:
        logger.warning("analytics initialization failed: %s", type(exc).__name__)
        _engine = None
        return None


def get_engine() -> Engine | None:
    return _engine or init_analytics()


def get_session_id(req: Request | None = None, payload: dict[str, Any] | None = None) -> str:
    req = req or (request if has_request_context() else None)
    if payload:
        value = payload.get("session_id") or payload.get("sessionId")
        if value:
            return str(value)[:80]
    if req is not None:
        value = req.cookies.get("analytics_session_id") or req.headers.get("X-Analytics-Session-Id")
        if value:
            return str(value)[:80]
    return secrets.token_urlsafe(24)


def hash_ip(ip: str | None) -> str | None:
    if not ip:
        return None
    salt = os.environ.get("ANALYTICS_IP_SALT", "")
    return hashlib.sha256(f"{salt}:{ip}".encode("utf-8")).hexdigest()


def _client_ip(req: Request | None) -> str | None:
    if req is None:
        return None
    forwarded_for = req.headers.get("X-Forwarded-For", "")
    if forwarded_for:
        return forwarded_for.split(",", 1)[0].strip()
    return req.remote_addr


def _clean_meta(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return value[:200]
    if isinstance(value, list):
        return [_clean_meta(item) for item in value[:20]]
    if isinstance(value, dict):
        clean: dict[str, Any] = {}
        blocked = {"name", "姓名", "phone", "手机号", "id_card", "身份证", "photo", "照片", "file", "excel"}
        for key, item in value.items():
            key_text = str(key)[:80]
            if any(word in key_text.lower() for word in blocked):
                continue
            clean[key_text] = _clean_meta(item)
        return clean
    return str(value)[:200]


def track_event(
    event_name: str,
    *,
    success: bool | None = None,
    page_path: str | None = None,
    card_count: int | None = None,
    row_count: int | None = None,
    error_message: str | None = None,
    meta_json: dict[str, Any] | None = None,
    session_id: str | None = None,
    req: Request | None = None,
) -> str | None:
    """Persist one analytics event. Never raises to callers."""
    try:
        engine = get_engine()
        if engine is None:
            return session_id

        req = req or (request if has_request_context() else None)
        payload = None
        if req is not None and req.is_json:
            payload = req.get_json(silent=True) or {}
        session_id = session_id or get_session_id(req, payload)
        resolved_path = page_path or (req.path if req is not None else None)
        user_agent = str(req.user_agent)[:500] if req is not None else None
        ip_hash = hash_ip(_client_ip(req))
        err = str(error_message)[:500] if error_message else None
        meta_text = json.dumps(_clean_meta(meta_json), ensure_ascii=False) if meta_json else None

        with engine.begin() as conn:
            conn.execute(
                analytics_events.insert().values(
                    created_at=datetime.now(timezone.utc),
                    event_name=event_name[:80],
                    session_id=session_id,
                    success=success,
                    page_path=resolved_path[:500] if resolved_path else None,
                    card_count=card_count,
                    row_count=row_count,
                    error_message=err,
                    user_agent=user_agent,
                    ip_hash=ip_hash,
                    meta_json=meta_text,
                )
            )
        return session_id
    except Exception as exc:
        logger.warning("analytics event write failed: %s", type(exc).__name__)
        return session_id


def count_events_since(start: datetime, event_name: str, success: bool | None = None) -> int:
    engine = get_engine()
    if engine is None:
        return 0
    stmt = select(func.count()).select_from(analytics_events).where(
        analytics_events.c.created_at >= start,
        analytics_events.c.event_name == event_name,
    )
    if success is not None:
        stmt = stmt.where(analytics_events.c.success.is_(success))
    with engine.begin() as conn:
        return int(conn.execute(stmt).scalar() or 0)
