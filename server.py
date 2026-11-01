#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AFE TestFlow local backend.

Standard library only (openpyxl is optional, used solely for XLSX export).
Run:  python server.py [--host 127.0.0.1] [--port 3080]

Storage layout (all under this folder):
    data/afe-testflow.db      SQLite database (WAL)
    data/uploads/             attachment files, opaque generated names
    public/                   static frontend (owned by frontend agent)
    seed.json                 initial tests/summary/instructions (optional)
"""

from __future__ import annotations

import argparse
import io
import json
import mimetypes
import os
import posixpath
import re
import shutil
import socket
import sqlite3
import sys
import threading
import traceback
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlsplit

APP_DIR = Path(__file__).resolve().parent

import workbook_import  # noqa: E402  (local module next to this file)

MAX_FILE_BYTES = 50 * 1024 * 1024          # per file, 50 MiB
MAX_REQUEST_BYTES = 100 * 1024 * 1024      # whole multipart request
MAX_JSON_BYTES = 1024 * 1024               # JSON bodies
MAX_FILENAME_CHARS = 120
MAX_PROJECT_BYTES = 20 * 1024 * 1024       # uploaded project workbook, 20 MiB
MAX_TEST_TEXT = 4000                       # editable guidance field length
TEST_ID_RE = re.compile(r"^[A-Za-z0-9._-]{1,32}$")

DEFAULT_PROJECT_ID = "afe"
DEFAULT_PROJECT_NAME = "AFE 推荐测试项目"

DATA_EXTS = {".xlsx", ".xls", ".xlsm"}
IMAGE_EXTS = {".jpg", ".jpeg"}
MODES = {"simulation", "measurement"}
STATUSES = {"", "Pass", "Fail"}

TEST_FIELDS = (
    "id", "priority", "module", "title", "purpose", "stimulus", "metrics",
    "conditions", "calculation", "criteria", "instruments", "source_result",
    "source_unit", "source_status", "source_date", "source_notes",
)

# Editable per-test guidance (the workbook columns the user can maintain in the UI).
GUIDANCE_FIELDS = (
    "purpose", "stimulus", "metrics", "conditions", "calculation", "criteria",
    "instruments",
)

RECORD_FIELDS = (
    "id", "test_id", "mode", "result", "unit", "status", "test_date",
    "notes", "created_at", "updated_at",
)

LOCAL_HOSTS = {"localhost", "127.0.0.1", "::1"}

# Host names accepted in the Host / Origin headers.  Starts with the loopback
# names and is extended at startup with every address this machine answers on
# plus whatever AFE_ALLOWED_HOSTS lists (comma/space/semicolon separated).
ALLOWED_HOSTS = {name.lower() for name in LOCAL_HOSTS}


def _normalize_host(name: str) -> str:
    name = name.strip().lower()
    if name.startswith("["):
        end = name.find("]")
        return name[1:end] if end > 0 else ""
    if name.count(":") == 1:          # host:port
        name = name.split(":", 1)[0]
    return name


def _local_ipv4_addresses() -> set[str]:
    """Best-effort list of this machine's own IPv4 addresses."""
    found: set[str] = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            found.add(info[4][0])
    except OSError:
        pass
    for probe in ("8.8.8.8", "1.1.1.1"):
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            sock.connect((probe, 53))       # UDP connect sends no packet
            found.add(sock.getsockname()[0])
        except OSError:
            pass
        finally:
            sock.close()
    return found


def configure_allowed_hosts(host: str | None = None) -> set[str]:
    """Allow loopback, LAN addresses of this box, and AFE_ALLOWED_HOSTS."""
    candidates = set(_local_ipv4_addresses())
    if host and host not in ("0.0.0.0", "::", ""):
        candidates.add(host)
    for part in re.split(r"[,\s;]+", os.environ.get("AFE_ALLOWED_HOSTS", "")):
        if part.strip():
            candidates.add(part)
    for name in candidates:
        name = _normalize_host(name)
        if name:
            ALLOWED_HOSTS.add(name)
    return ALLOWED_HOSTS


class Config:
    """Runtime configuration; replaceable in tests."""

    def __init__(self, base_dir=None, host="127.0.0.1", port=3080,
                 data_dir=None, seed_path=None):
        self.base_dir = Path(base_dir).resolve() if base_dir else APP_DIR
        self.host = host
        self.port = port
        self.data_dir = Path(data_dir).resolve() if data_dir else self.base_dir / "data"
        self.uploads_dir = self.data_dir / "uploads"
        self.public_dir = self.base_dir / "public"
        self.db_path = self.data_dir / "afe-testflow.db"
        self.seed_path = Path(seed_path).resolve() if seed_path \
            else self.base_dir / "seed.json"

    def set_data_dir(self, data_dir) -> None:
        self.data_dir = Path(data_dir).resolve()
        self.uploads_dir = self.data_dir / "uploads"
        self.db_path = self.data_dir / "afe-testflow.db"

    def set_seed_path(self, seed_path) -> None:
        self.seed_path = Path(seed_path).resolve()


def _env_port(default=3080):
    try:
        return int(os.environ.get("AFE_PORT", default))
    except (TypeError, ValueError):
        return default


CONFIG = Config(host=os.environ.get("AFE_HOST", "127.0.0.1"), port=_env_port())

_QUIET = os.environ.get("AFE_QUIET", "").strip() not in ("", "0", "false", "no")


class ApiError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status
        self.message = message


# --------------------------------------------------------------------------
# storage / database
# --------------------------------------------------------------------------

def init_storage() -> None:
    CONFIG.data_dir.mkdir(parents=True, exist_ok=True)
    CONFIG.uploads_dir.mkdir(parents=True, exist_ok=True)


def connect_db() -> sqlite3.Connection:
    conn = sqlite3.connect(str(CONFIG.db_path), timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 15000")
    return conn


SCHEMA = """
CREATE TABLE IF NOT EXISTS projects (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    description TEXT NOT NULL DEFAULT '',
    source TEXT NOT NULL DEFAULT '',
    summary TEXT NOT NULL DEFAULT '[]',
    instructions TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    position INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS tests (
    project_id TEXT NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    id TEXT NOT NULL,
    priority TEXT NOT NULL DEFAULT '',
    module TEXT NOT NULL DEFAULT '',
    title TEXT NOT NULL DEFAULT '',
    purpose TEXT NOT NULL DEFAULT '',
    stimulus TEXT NOT NULL DEFAULT '',
    metrics TEXT NOT NULL DEFAULT '',
    conditions TEXT NOT NULL DEFAULT '',
    calculation TEXT NOT NULL DEFAULT '',
    criteria TEXT NOT NULL DEFAULT '',
    instruments TEXT NOT NULL DEFAULT '',
    source_result TEXT NOT NULL DEFAULT '',
    source_unit TEXT NOT NULL DEFAULT '',
    source_status TEXT NOT NULL DEFAULT '',
    source_date TEXT NOT NULL DEFAULT '',
    source_notes TEXT NOT NULL DEFAULT '',
    position INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (project_id, id)
);
CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    test_id TEXT NOT NULL,
    mode TEXT NOT NULL,
    result TEXT NOT NULL DEFAULT '',
    unit TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    test_date TEXT NOT NULL DEFAULT '',
    notes TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (project_id, test_id) REFERENCES tests(project_id, id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_records_test ON records(project_id, test_id, id);
CREATE TABLE IF NOT EXISTS clipboard (
    position INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id TEXT NOT NULL,
    test_id TEXT NOT NULL,
    op TEXT NOT NULL,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS attachments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    record_id INTEGER NOT NULL REFERENCES records(id) ON DELETE CASCADE,
    kind TEXT NOT NULL,
    name TEXT NOT NULL,
    size INTEGER NOT NULL,
    stored_name TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_attachments_record ON attachments(record_id, id);
"""


def _table_exists(conn, name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (name,)
    ).fetchone()
    return row is not None


def _table_columns(conn, name: str) -> set:
    return {row[1] for row in conn.execute(f'PRAGMA table_info("{name}")')}


def _migrate_single_project(conn) -> None:
    """Upgrade the pre-multi-project schema (tests/records/attachments) in place."""
    print("[migrate] upgrading database to the multi-project schema", flush=True)
    meta = {}
    if _table_exists(conn, "meta"):
        for row in conn.execute("SELECT key, value FROM meta"):
            meta[row["key"]] = row["value"]

    def as_list(key):
        raw = meta.get(key)
        if not raw:
            return []
        try:
            value = json.loads(raw)
        except (TypeError, ValueError):
            return []
        return value if isinstance(value, list) else []

    summary = as_list("summary")
    instructions = as_list("instructions")
    source = meta.get("source", "") or ""
    conn.execute("PRAGMA foreign_keys = OFF")
    try:
        conn.executescript(
            "DROP INDEX IF EXISTS idx_records_test;"
            "DROP INDEX IF EXISTS idx_attachments_record;"
        )
        conn.execute("ALTER TABLE tests RENAME TO tests_v1")
        conn.execute("ALTER TABLE records RENAME TO records_v1")
        conn.execute("ALTER TABLE attachments RENAME TO attachments_v1")
        conn.executescript(SCHEMA)
        timestamp = now_iso()
        conn.execute(
            "INSERT INTO projects (id, name, description, source, summary, instructions,"
            " created_at, updated_at, position) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)",
            (DEFAULT_PROJECT_ID, DEFAULT_PROJECT_NAME, "由旧版单项目数据自动迁移", source,
             json.dumps(summary, ensure_ascii=False),
             json.dumps(instructions, ensure_ascii=False), timestamp, timestamp),
        )
        columns = ", ".join(TEST_FIELDS)
        conn.execute(
            f"INSERT INTO tests (project_id, {columns}, position)"
            f" SELECT ?, {columns}, position FROM tests_v1",
            (DEFAULT_PROJECT_ID,),
        )
        conn.execute(
            "INSERT INTO records (id, project_id, test_id, mode, result, unit, status,"
            " test_date, notes, created_at, updated_at)"
            " SELECT id, ?, test_id, mode, result, unit, status, test_date, notes,"
            " created_at, updated_at FROM records_v1",
            (DEFAULT_PROJECT_ID,),
        )
        conn.execute(
            "INSERT INTO attachments (id, record_id, kind, name, size, stored_name, created_at)"
            " SELECT id, record_id, kind, name, size, stored_name, created_at FROM attachments_v1"
        )
        conn.executescript(
            "DROP TABLE tests_v1; DROP TABLE records_v1; DROP TABLE attachments_v1;"
        )
        conn.commit()
    finally:
        conn.execute("PRAGMA foreign_keys = ON")


def init_db() -> None:
    conn = connect_db()
    try:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA foreign_keys = ON")
        if _table_exists(conn, "tests") and "project_id" not in _table_columns(conn, "tests"):
            _migrate_single_project(conn)
        conn.executescript(SCHEMA)
        conn.commit()
    finally:
        conn.close()


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _as_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float, bool)):
        return str(value)
    return json.dumps(value, ensure_ascii=False)


def seed_default_project() -> None:
    """Create the built-in project from seed.json when the database has none."""
    conn = connect_db()
    try:
        if conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]:
            return
        path = CONFIG.seed_path
        raw = {}
        if path.is_file():
            try:
                loaded = json.loads(path.read_text(encoding="utf-8-sig"))
                if isinstance(loaded, dict):
                    raw = loaded
                else:
                    print("[seed] seed root must be an object; ignored", flush=True)
            except Exception as exc:  # noqa: BLE001 - startup diagnostics
                print(f"[seed] invalid seed.json: {exc}", flush=True)
        else:
            print(f"[seed] not found: {path} (starting with no project)", flush=True)
            return

        tests = raw.get("tests") if isinstance(raw.get("tests"), list) else []
        summary = raw.get("summary") if isinstance(raw.get("summary"), list) else []
        instructions = raw.get("instructions") if isinstance(raw.get("instructions"), list) else []
        name = _as_text(raw.get("project_name")).strip() or DEFAULT_PROJECT_NAME
        source = _as_text(raw.get("source")) or (path.name if path.is_file() else "")
        timestamp = now_iso()
        with conn:
            conn.execute(
                "INSERT INTO projects (id, name, description, source, summary, instructions,"
                " created_at, updated_at, position) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)",
                (DEFAULT_PROJECT_ID, name, "", source,
                 json.dumps(summary, ensure_ascii=False),
                 json.dumps(instructions, ensure_ascii=False), timestamp, timestamp),
            )
            for index, item in enumerate(tests):
                if not isinstance(item, dict):
                    continue
                test_id = _as_text(item.get("id")).strip()
                if not test_id:
                    continue
                values = [DEFAULT_PROJECT_ID]
                values.extend(test_id if f == "id" else _as_text(item.get(f)) for f in TEST_FIELDS)
                values.append(index)
                placeholders = ", ".join(["?"] * (len(TEST_FIELDS) + 2))
                conn.execute(
                    f"INSERT INTO tests (project_id, {', '.join(TEST_FIELDS)}, position)"
                    f" VALUES ({placeholders})",
                    values,
                )
        print(f"[seed] created project {DEFAULT_PROJECT_ID!r} with {len(tests)} tests", flush=True)
    finally:
        conn.close()


# --------------------------------------------------------------------------
# serialization helpers
# --------------------------------------------------------------------------

def _attachment_dict(row) -> dict:
    return {
        "id": row["id"],
        "record_id": row["record_id"],
        "kind": row["kind"],
        "name": row["name"],
        "size": row["size"],
        "url": f"/api/attachments/{row['id']}/file",
    }


def _record_dict(conn, row, attachments=None) -> dict:
    if attachments is None:
        attachments = [
            _attachment_dict(a)
            for a in conn.execute(
                "SELECT * FROM attachments WHERE record_id = ? ORDER BY id", (row["id"],)
            )
        ]
    data = {field: row[field] for field in RECORD_FIELDS}
    data["attachments"] = attachments
    return data


def _json_list(raw, default=None):
    default = [] if default is None else default
    try:
        value = json.loads(raw) if raw else default
    except (TypeError, ValueError):
        return list(default)
    return value if isinstance(value, list) else list(default)


def project_row(conn, project_id: str):
    return conn.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()


def project_dict(conn, row) -> dict:
    counts = conn.execute(
        "SELECT (SELECT COUNT(*) FROM tests WHERE project_id = ?) AS tests,"
        " (SELECT COUNT(*) FROM records WHERE project_id = ?) AS records",
        (row["id"], row["id"]),
    ).fetchone()
    return {
        "id": row["id"],
        "name": row["name"],
        "description": row["description"],
        "source": row["source"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
        "test_count": counts["tests"],
        "record_count": counts["records"],
    }


def project_bootstrap(conn, project_id: str) -> dict:
    row = project_row(conn, project_id)
    if row is None:
        raise ApiError(404, "项目不存在")
    tests = [
        {field: item[field] for field in TEST_FIELDS}
        for item in conn.execute(
            "SELECT * FROM tests WHERE project_id = ? ORDER BY position, id", (project_id,)
        )
    ]
    return {
        "project": project_dict(conn, row),
        "tests": tests,
        "summary": _json_list(row["summary"]),
        "instructions": _json_list(row["instructions"]),
        "source": row["source"],
    }


def _project_slug(name: str) -> str:
    base = re.sub(r"[^a-z0-9]+", "-", (name or "").lower()).strip("-")[:40]
    return base or "project"


def _next_test_id(conn, project_id: str) -> str:
    """Suggest the next free id, keeping the project's prefix/width (P01 -> P35, else T01)."""
    ids = [row["id"] for row in conn.execute(
        "SELECT id FROM tests WHERE project_id = ?", (project_id,)
    )]
    groups = {}
    for value in ids:
        match = re.fullmatch(r"([A-Za-z][A-Za-z0-9_-]*?)(\d+)", value)
        if match:
            groups.setdefault(match.group(1), []).append((int(match.group(2)), len(match.group(2))))
    if groups:
        prefix = max(groups, key=lambda key: len(groups[key]))
        numbers = groups[prefix]
        number = max(item[0] for item in numbers) + 1
        width = max(item[1] for item in numbers)
        candidate = f"{prefix}{number:0{width}d}"
    else:
        candidate = "T01"
    while any(candidate == value for value in ids):
        match = re.fullmatch(r"([A-Za-z][A-Za-z0-9_-]*?)(\d+)", candidate)
        if not match:
            candidate = candidate + "-2"
            continue
        candidate = f"{match.group(1)}{int(match.group(2)) + 1:0{len(match.group(2))}d}"
    return candidate


def _unique_project_id(conn, name: str) -> str:
    base = _project_slug(name)
    candidate = base
    counter = 2
    while project_row(conn, candidate) is not None:
        candidate = f"{base}-{counter}"
        counter += 1
    return candidate


# --------------------------------------------------------------------------
# input validation
# --------------------------------------------------------------------------

def _text_field(body: dict, key: str, *, maxlen: int, fallback: str = "") -> str:
    value = body.get(key, fallback)
    if value is None:
        value = ""
    if not isinstance(value, str):
        raise ApiError(400, f"字段 {key} 必须是字符串")
    if len(value) > maxlen:
        raise ApiError(400, f"字段 {key} 超长（最多 {maxlen} 字符）")
    return value.strip()


def _validate_payload(body: dict, existing=None) -> dict:
    if not isinstance(body, dict):
        raise ApiError(400, "请求体必须是 JSON 对象")
    mode = body.get("mode", existing["mode"] if existing else None)
    if mode not in MODES:
        raise ApiError(400, "mode 必须为 simulation 或 measurement")
    status = _text_field(body, "status", maxlen=20,
                         fallback=existing["status"] if existing else "")
    if status not in STATUSES:
        raise ApiError(400, "status 必须为 Pass、Fail 或空")
    test_date = _text_field(body, "test_date", maxlen=10,
                            fallback=existing["test_date"] if existing else "")
    if test_date:
        try:
            parsed = datetime.strptime(test_date, "%Y-%m-%d")
        except ValueError:
            raise ApiError(400, "test_date 必须为 YYYY-MM-DD 格式的有效日期") from None
        if parsed.strftime("%Y-%m-%d") != test_date:
            raise ApiError(400, "test_date 必须为 YYYY-MM-DD 格式的有效日期")
    return {
        "mode": mode,
        "result": _text_field(body, "result", maxlen=500,
                              fallback=existing["result"] if existing else ""),
        "unit": _text_field(body, "unit", maxlen=100,
                            fallback=existing["unit"] if existing else ""),
        "status": status,
        "test_date": test_date,
        "notes": _text_field(body, "notes", maxlen=20000,
                             fallback=existing["notes"] if existing else ""),
    }


def _sanitize_filename(name: str, fallback: str) -> str:
    name = (name or "").replace("\\", "/").split("/")[-1]
    name = "".join(ch for ch in name if ch >= " " and ch != "\x7f").strip().strip(".")
    if len(name) > MAX_FILENAME_CHARS:
        suffix = Path(name).suffix[:16]
        name = name[: MAX_FILENAME_CHARS - len(suffix)] + suffix
    return name or fallback


def _check_file_content(ext: str, data: bytes) -> str | None:
    if ext in {".xlsx", ".xlsm"}:
        if not data.startswith(b"PK\x03\x04"):
            return "文件内容不是有效的 XLSX/XLSM（ZIP）格式"
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = set(zf.namelist())
        except (zipfile.BadZipFile, OSError):
            return "XLSX/XLSM 文件已损坏"
        if "[Content_Types].xml" not in names or not any(n.startswith("xl/") for n in names):
            return "XLSX/XLSM 缺少工作簿结构"
        return None
    if ext == ".xls":
        if not data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
            return "文件内容不是有效的 XLS（OLE2）格式"
        return None
    if not data.startswith(b"\xff\xd8\xff"):
        return "文件内容不是有效的 JPG 图片"
    return None


def _validate_upload(kind: str, filename: str, data: bytes) -> tuple[str, str]:
    """Return (display_name, extension) or raise ApiError."""
    if kind not in ("data", "image"):
        raise ApiError(400, "kind 必须为 data 或 image")
    ext = Path(filename).suffix.lower()
    if kind == "data":
        if ext not in DATA_EXTS:
            raise ApiError(400, "数据附件仅支持 xlsx、xls、xlsm 格式")
    else:
        if ext not in IMAGE_EXTS:
            raise ApiError(400, "图片附件仅支持 jpg、jpeg 格式")
    if not data:
        raise ApiError(400, "附件不能为空")
    if len(data) > MAX_FILE_BYTES:
        raise ApiError(400, "单个附件不得超过 50 MiB")
    problem = _check_file_content(ext, data)
    if problem:
        raise ApiError(400, problem)
    safe_name = _sanitize_filename(filename, f"attachment{ext}")
    if not Path(safe_name).suffix:
        safe_name += ext
    return safe_name, ext


# --------------------------------------------------------------------------
# multipart parsing
# --------------------------------------------------------------------------

def parse_multipart(headers, body: bytes) -> tuple[dict, list]:
    ctype = headers.get("Content-Type", "")
    if not ctype.lower().startswith("multipart/form-data"):
        raise ApiError(415, "内容类型必须为 multipart/form-data")
    try:
        message = BytesParser(policy=policy.default).parsebytes(
            b"Content-Type: " + ctype.encode("latin-1", "replace")
            + b"\r\nMIME-Version: 1.0\r\n\r\n" + body
        )
    except Exception as exc:  # noqa: BLE001
        raise ApiError(400, "无法解析 multipart 请求") from exc
    if not message.is_multipart():
        raise ApiError(400, "multipart 请求缺少边界")
    fields: dict = {}
    files: list = []
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition") or ""
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        if filename is None:
            try:
                fields[name] = payload.decode("utf-8")
            except UnicodeDecodeError:
                fields[name] = payload.decode("utf-8", "replace")
        else:
            files.append((name, str(filename), payload))
    return fields, files


# --------------------------------------------------------------------------
# HTTP helpers
# --------------------------------------------------------------------------

def _content_disposition(name: str, inline: bool = False) -> str:
    fallback = re.sub(r"[^\x20-\x7e]", "_", name).replace('"', "_").replace("\\", "_")
    fallback = fallback or "download"
    kind = "inline" if inline else "attachment"
    return f"{kind}; filename=\"{fallback}\"; filename*=UTF-8''{quote(name, safe='')}"


def _json_bytes(obj) -> bytes:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")


STATIC_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".mjs": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".ico": "image/x-icon",
    ".woff": "font/woff",
    ".woff2": "font/woff2",
    ".txt": "text/plain; charset=utf-8",
    ".map": "application/json; charset=utf-8",
}


class Handler(BaseHTTPRequestHandler):
    server_version = "AFETestFlow/1.0"
    protocol_version = "HTTP/1.1"

    # ---- logging ---------------------------------------------------------
    def log_message(self, fmt, *args):  # noqa: D401
        if not _QUIET:
            sys.stderr.write("[http] %s %s\n" % (self.address_string(), fmt % args))

    # ---- security --------------------------------------------------------
    def _host_value(self) -> str:
        return self.headers.get("Host", "").strip()

    def _host_ok(self) -> bool:
        raw = self._host_value()
        if not raw or "@" in raw or "/" in raw or any(c in raw for c in " \t\r\n"):
            return False
        if raw.startswith("["):
            end = raw.find("]")
            if end < 0:
                return False
            hostname = raw[1:end]
            rest = raw[end + 1:]
            port = rest[1:] if rest.startswith(":") else None
        elif raw.count(":") == 1:
            hostname, port = raw.split(":", 1)
        elif ":" in raw:
            return False
        else:
            hostname, port = raw, None
        if hostname.lower() not in ALLOWED_HOSTS:
            return False
        actual_port = self.server.server_address[1]
        return port in (None, "", str(actual_port))

    def _origin_ok(self) -> bool:
        origin = self.headers.get("Origin")
        if origin is None:
            return True
        if not origin or origin == "null":
            return False
        try:
            parts = urlsplit(origin)
        except ValueError:
            return False
        if parts.scheme not in ("http", "https") or "@" in parts.netloc:
            return False
        host = parts.hostname or ""
        if host.lower() not in ALLOWED_HOSTS:
            return False
        actual_port = self.server.server_address[1]
        try:
            port = parts.port
        except ValueError:
            return False
        if port is None:
            port = 80 if parts.scheme == "http" else 443
        return port == actual_port

    # ---- responses -------------------------------------------------------
    def _base_headers(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("X-Frame-Options", "DENY")

    def _send(self, status: int, body: bytes, content_type: str,
              extra_headers=None, head_only: bool = False) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self._base_headers()
        for key, value in (extra_headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        if head_only or self.command == "HEAD":
            return
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError, OSError):
            self.close_connection = True

    def _send_json(self, status: int, obj) -> None:
        self._send(status, _json_bytes(obj), "application/json; charset=utf-8",
                   {"Cache-Control": "no-store"})

    def _error(self, status: int, message: str) -> None:
        extra = {"Cache-Control": "no-store"}
        if status in (411, 413):
            extra["Connection"] = "close"
            self.close_connection = True
        self._send(status, _json_bytes({"error": message}),
                   "application/json; charset=utf-8", extra)

    # ---- request body ----------------------------------------------------
    def _read_body(self, limit: int) -> bytes:
        raw = self.headers.get("Content-Length")
        if raw is None:
            raise ApiError(411, "缺少 Content-Length")
        try:
            length = int(raw)
        except (TypeError, ValueError):
            raise ApiError(400, "Content-Length 无效") from None
        if length < 0:
            raise ApiError(400, "Content-Length 无效")
        if length > limit:
            limit_mib = limit // (1024 * 1024)
            raise ApiError(413, f"请求体超过限制（{limit_mib} MiB）")
        data = self.rfile.read(length)
        if len(data) != length:
            raise ApiError(400, "请求体不完整")
        self._body_read = True
        return data

    def _discard_body(self, limit: int = MAX_JSON_BYTES) -> None:
        if getattr(self, "_body_read", False):
            return
        raw = self.headers.get("Content-Length")
        if raw is None:
            return
        try:
            length = int(raw)
        except (TypeError, ValueError):
            return
        if 0 < length <= limit:
            try:
                self.rfile.read(length)
            except OSError:
                self.close_connection = True
            else:
                self._body_read = True

    def _read_json(self) -> dict:
        ctype = self.headers.get("Content-Type", "").split(";")[0].strip().lower()
        if ctype != "application/json":
            raise ApiError(415, "Content-Type 必须为 application/json")
        body = self._read_body(MAX_JSON_BYTES)
        try:
            data = json.loads(body.decode("utf-8")) if body else {}
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ApiError(400, "JSON 格式无效") from None
        if not isinstance(data, dict):
            raise ApiError(400, "请求体必须是 JSON 对象")
        return data

    # ---- routing ---------------------------------------------------------
    def do_GET(self):
        self._dispatch("GET")

    def do_HEAD(self):
        self._dispatch("HEAD")

    def do_POST(self):
        self._dispatch("POST")

    def do_PUT(self):
        self._dispatch("PUT")

    def do_DELETE(self):
        self._dispatch("DELETE")

    def _dispatch(self, method: str):
        self._body_read = False
        try:
            if not self._host_ok():
                raise ApiError(403, "仅允许通过 localhost 访问")
            if method in ("POST", "PUT", "DELETE") and not self._origin_ok():
                raise ApiError(403, "请求来源不被允许")
            path = urlsplit(self.path).path or "/"
            if path == "/api" or path.startswith("/api/"):
                self._api(method, path)
            elif method in ("GET", "HEAD"):
                self._static(path)
            else:
                raise ApiError(405, "方法不允许")
        except ApiError as exc:
            # Drain small rejected bodies first: closing a socket with unread
            # inbound data makes Windows send RST and can discard our error
            # response. Oversized bodies are still rejected by closing.
            if method in ("POST", "PUT", "DELETE") and exc.status not in (411, 413):
                self._discard_body()
            self._error(exc.status, exc.message)
        except (BrokenPipeError, ConnectionResetError):
            self.close_connection = True
        except Exception:  # noqa: BLE001
            if not _QUIET:
                traceback.print_exc()
            self._error(500, "服务器内部错误")
        finally:
            if method in ("POST", "PUT", "DELETE"):
                raw = self.headers.get("Content-Length")
                # If a body was not consumed the connection must be dropped.
                try:
                    if raw is not None and int(raw) > 0 and not getattr(self, "_body_read", False):
                        self.close_connection = True
                except (TypeError, ValueError):
                    self.close_connection = True

    def _api(self, method: str, path: str):
        if path == "/api/health":
            if method in ("GET", "HEAD"):
                return self._send_json(200, {"ok": True})
            raise ApiError(405, "方法不允许")

        if path == "/api/projects":
            if method in ("GET", "HEAD"):
                return self._list_projects()
            if method == "POST":
                return self._create_project()
            raise ApiError(405, "方法不允许")

        if path == "/api/bootstrap":
            if method in ("GET", "HEAD"):
                return self._default_bootstrap()
            raise ApiError(405, "方法不允许")

        if path == "/api/export.xlsx":
            if method in ("GET", "HEAD"):
                return self._export_xlsx(None)
            raise ApiError(405, "方法不允许")

        if path == "/api/backup.zip":
            if method in ("GET", "HEAD"):
                return self._backup_zip()
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/projects/([A-Za-z0-9_-]{1,64})", path)
        if match:
            project_id = match.group(1)
            if method in ("GET", "HEAD"):
                return self._get_project(project_id)
            if method == "DELETE":
                return self._delete_project(project_id)
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/projects/([A-Za-z0-9_-]{1,64})/bootstrap", path)
        if match:
            if method in ("GET", "HEAD"):
                conn = connect_db()
                try:
                    return self._send_json(200, project_bootstrap(conn, match.group(1)))
                finally:
                    conn.close()
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/projects/([A-Za-z0-9_-]{1,64})/export\.xlsx", path)
        if match:
            if method in ("GET", "HEAD"):
                return self._export_xlsx(match.group(1))
            raise ApiError(405, "方法不允许")

        if path == "/api/clipboard":
            if method in ("GET", "HEAD"):
                return self._get_clipboard()
            if method == "POST":
                return self._set_clipboard()
            if method == "DELETE":
                return self._clear_clipboard()
            raise ApiError(405, "方法不允许")

        if path == "/api/clipboard/paste":
            if method == "POST":
                return self._paste_clipboard()
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/projects/([A-Za-z0-9_-]{1,64})/tests", path)
        if match:
            if method == "POST":
                return self._create_test(match.group(1))
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/projects/([A-Za-z0-9_-]{1,64})/tests/([^/]+)", path)
        if match:
            project_id = match.group(1)
            test_id = unquote(match.group(2))
            if method == "PUT":
                return self._update_test(project_id, test_id)
            if method == "DELETE":
                return self._delete_test(project_id, test_id)
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/projects/([A-Za-z0-9_-]{1,64})/tests/([^/]+)/records", path)
        if match:
            project_id = match.group(1)
            test_id = unquote(match.group(2))
            if method in ("GET", "HEAD"):
                return self._list_records(project_id, test_id)
            if method == "POST":
                return self._create_record(project_id, test_id)
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/records/(\d+)", path)
        if match:
            record_id = int(match.group(1))
            if method == "PUT":
                return self._update_record(record_id)
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/records/(\d+)/attachments", path)
        if match:
            record_id = int(match.group(1))
            if method == "POST":
                return self._upload_attachments(record_id)
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/attachments/(\d+)", path)
        if match:
            if method == "DELETE":
                return self._delete_attachment(int(match.group(1)))
            raise ApiError(405, "方法不允许")

        match = re.fullmatch(r"/api/attachments/(\d+)/file", path)
        if match:
            if method in ("GET", "HEAD"):
                return self._download_attachment(int(match.group(1)))
            raise ApiError(405, "方法不允许")

        raise ApiError(404, "接口不存在")

    # ---- project endpoints -----------------------------------------------
    def _list_projects(self):
        conn = connect_db()
        try:
            rows = conn.execute(
                "SELECT * FROM projects ORDER BY position, created_at, id"
            ).fetchall()
            self._send_json(200, {"projects": [project_dict(conn, row) for row in rows]})
        finally:
            conn.close()

    def _default_bootstrap(self):
        conn = connect_db()
        try:
            row = conn.execute(
                "SELECT id FROM projects ORDER BY position, created_at, id LIMIT 1"
            ).fetchone()
            if row is None:
                raise ApiError(404, "还没有任何项目")
            self._send_json(200, project_bootstrap(conn, row["id"]))
        finally:
            conn.close()

    def _get_project(self, project_id: str):
        conn = connect_db()
        try:
            self._send_json(200, project_bootstrap(conn, project_id))
        finally:
            conn.close()

    def _create_project(self):
        ctype = self.headers.get("Content-Type", "")
        if not ctype.lower().startswith("multipart/form-data"):
            raise ApiError(415, "内容类型必须为 multipart/form-data")
        body = self._read_body(MAX_PROJECT_BYTES)
        self._body_read = True
        fields, files = parse_multipart(self.headers, body)
        if not files:
            raise ApiError(400, "请选择要上传的测试项目表格（.xlsx / .xlsm）")
        if len(files) > 1:
            raise ApiError(400, "一次只能上传一个表格文件")
        _field, filename, payload = files[0]
        ext = Path(filename).suffix.lower()
        if ext == ".xls":
            raise ApiError(400, "暂不支持 .xls，请在 Excel 中另存为 .xlsx 后再上传")
        if ext not in DATA_EXTS:
            raise ApiError(400, "测试项目仅支持 .xlsx 或 .xlsm 表格")
        if not payload:
            raise ApiError(400, "表格文件为空")
        if len(payload) > MAX_PROJECT_BYTES:
            raise ApiError(400, "表格文件不得超过 20 MiB")
        problem = _check_file_content(ext, payload)
        if problem:
            raise ApiError(400, problem)
        try:
            parsed = workbook_import.parse_workbook(payload, filename)
        except workbook_import.WorkbookError as exc:
            raise ApiError(400, str(exc)) from None

        name = _as_text(fields.get("name")).strip() or Path(filename).stem.strip() or "未命名项目"
        name = name[:80]
        conn = connect_db()
        try:
            project_id = _unique_project_id(conn, name)
            timestamp = now_iso()
            position = conn.execute(
                "SELECT COALESCE(MAX(position), -1) + 1 FROM projects"
            ).fetchone()[0]
            with conn:
                conn.execute(
                    "INSERT INTO projects (id, name, description, source, summary,"
                    " instructions, created_at, updated_at, position)"
                    " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (project_id, name, "", parsed["source"],
                     json.dumps(parsed["summary"], ensure_ascii=False),
                     json.dumps(parsed["instructions"], ensure_ascii=False),
                     timestamp, timestamp, position),
                )
                placeholders = ", ".join(["?"] * (len(TEST_FIELDS) + 2))
                for index, item in enumerate(parsed["tests"]):
                    values = [project_id]
                    values.extend(item.get(field, "") for field in TEST_FIELDS)
                    values.append(index)
                    conn.execute(
                        f"INSERT INTO tests (project_id, {', '.join(TEST_FIELDS)}, position)"
                        f" VALUES ({placeholders})",
                        values,
                    )
            row = project_row(conn, project_id)
            self._send_json(201, {
                "project": project_dict(conn, row),
                "tests": len(parsed["tests"]),
                "summary": len(parsed["summary"]),
                "instructions": len(parsed["instructions"]),
            })
        finally:
            conn.close()

    def _delete_project(self, project_id: str):
        conn = connect_db()
        try:
            if project_row(conn, project_id) is None:
                raise ApiError(404, "项目不存在")
            total = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
            if total <= 1:
                raise ApiError(400, "至少需要保留一个项目")
            stored = [
                row["stored_name"]
                for row in conn.execute(
                    "SELECT a.stored_name FROM attachments a"
                    " JOIN records r ON r.id = a.record_id WHERE r.project_id = ?",
                    (project_id,),
                )
            ]
            with conn:
                conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))
            for name in stored:
                if re.fullmatch(r"[0-9a-f]{32}\.(xlsx|xls|xlsm|jpg|jpeg)", name or ""):
                    try:
                        (CONFIG.uploads_dir / name).unlink()
                    except OSError:
                        pass
            self._send_json(200, {"ok": True, "id": project_id})
        finally:
            conn.close()

    # ---- clipboard (copy / cut / paste test items) -----------------------
    def _clipboard_payload(self, conn) -> dict:
        rows = conn.execute(
            "SELECT c.*, p.name AS project_name FROM clipboard c"
            " LEFT JOIN projects p ON p.id = c.project_id ORDER BY c.position"
        ).fetchall()
        items = []
        for row in rows:
            test = conn.execute(
                "SELECT * FROM tests WHERE project_id = ? AND id = ?",
                (row["project_id"], row["test_id"]),
            ).fetchone()
            items.append({
                "project_id": row["project_id"],
                "project_name": row["project_name"] or row["project_id"],
                "test_id": row["test_id"],
                "op": row["op"],
                "missing": test is None,
                "test": {field: test[field] for field in TEST_FIELDS} if test is not None else None,
            })
        return {"items": items, "count": len(items)}

    def _get_clipboard(self):
        conn = connect_db()
        try:
            self._send_json(200, self._clipboard_payload(conn))
        finally:
            conn.close()

    def _set_clipboard(self):
        body = self._read_json()
        self._body_read = True
        if not isinstance(body, dict):
            raise ApiError(400, "请求体必须是 JSON 对象")
        project_id = _as_text(body.get("project_id")).strip()
        op = _as_text(body.get("op")).strip()
        if op not in ("copy", "cut"):
            raise ApiError(400, "op 必须为 copy 或 cut")
        raw_ids = body.get("test_ids")
        if not isinstance(raw_ids, list) or not raw_ids:
            raise ApiError(400, "test_ids 必须是非空数组")
        if not isinstance(project_id, str) or not TEST_ID_RE.match(project_id):
            raise ApiError(400, "项目 ID 无效")
        test_ids = []
        for value in raw_ids[:50]:
            test_id = _as_text(value).strip()
            if test_id and test_id not in test_ids:
                test_ids.append(test_id)
        if not test_ids:
            raise ApiError(400, "test_ids 必须是非空数组")
        conn = connect_db()
        try:
            if project_row(conn, project_id) is None:
                raise ApiError(404, "项目不存在")
            found = []
            for test_id in test_ids:
                row = conn.execute(
                    "SELECT id FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
                ).fetchone()
                if row is None:
                    raise ApiError(404, f"测试项 {test_id} 不存在")
                found.append(test_id)
            timestamp = now_iso()
            with conn:
                conn.execute("DELETE FROM clipboard")
                for test_id in found:
                    conn.execute(
                        "INSERT INTO clipboard (project_id, test_id, op, created_at)"
                        " VALUES (?, ?, ?, ?)",
                        (project_id, test_id, op, timestamp),
                    )
            self._send_json(200, self._clipboard_payload(conn))
        finally:
            conn.close()

    def _clear_clipboard(self):
        conn = connect_db()
        try:
            with conn:
                conn.execute("DELETE FROM clipboard")
            self._send_json(200, {"ok": True, "count": 0})
        finally:
            conn.close()

    def _paste_clipboard(self):
        body = self._read_json()
        self._body_read = True
        if not isinstance(body, dict):
            raise ApiError(400, "请求体必须是 JSON 对象")
        target_project = _as_text(body.get("project_id")).strip()
        target_module = _as_text(body.get("module")).strip()[:60]
        conn = connect_db()
        try:
            items = conn.execute("SELECT * FROM clipboard ORDER BY position").fetchall()
            if not items:
                raise ApiError(400, "剪贴板为空")
            if target_project and project_row(conn, target_project) is None:
                raise ApiError(404, "目标项目不存在")
            pasted = []
            skipped = []
            moved_any = False
            with conn:
                for item in items:
                    source_project = item["project_id"]
                    source_test = item["test_id"]
                    row = conn.execute(
                        "SELECT * FROM tests WHERE project_id = ? AND id = ?",
                        (source_project, source_test),
                    ).fetchone()
                    if row is None:
                        skipped.append({"project_id": source_project, "test_id": source_test,
                                        "reason": "测试项已不存在"})
                        continue
                    dest_project = target_project or source_project
                    dest_module = target_module or row["module"]

                    # cut inside the same project is just a module change: no id / FK churn
                    if item["op"] == "cut" and dest_project == source_project:
                        conn.execute(
                            "UPDATE tests SET module = ? WHERE project_id = ? AND id = ?",
                            (dest_module, source_project, source_test),
                        )
                        pasted.append({
                            "source_project_id": source_project,
                            "source_test_id": source_test,
                            "project_id": dest_project,
                            "test_id": source_test,
                            "mode": "move-module",
                            "records": 0,
                        })
                        moved_any = True
                        continue

                    new_id = self._free_test_id(conn, dest_project, source_test)
                    position = conn.execute(
                        "SELECT COALESCE(MAX(position), -1) + 1 FROM tests WHERE project_id = ?",
                        (dest_project,),
                    ).fetchone()[0]
                    # insert the copy/moved row first, then re-point records, then drop the source
                    columns = ["project_id", "id", "priority", "module", "title"] + list(GUIDANCE_FIELDS) + [
                        "source_result", "source_unit", "source_status", "source_date",
                        "source_notes", "position",
                    ]
                    values = [dest_project, new_id, row["priority"], dest_module, row["title"]] + [
                        row[field] for field in GUIDANCE_FIELDS
                    ] + [row["source_result"], row["source_unit"], row["source_status"],
                         row["source_date"], row["source_notes"], position]
                    conn.execute(
                        f"INSERT INTO tests ({', '.join(columns)})"
                        f" VALUES ({', '.join(['?'] * len(columns))})",
                        values,
                    )
                    if item["op"] == "cut":
                        cursor = conn.execute(
                            "UPDATE records SET project_id = ?, test_id = ?"
                            " WHERE project_id = ? AND test_id = ?",
                            (dest_project, new_id, source_project, source_test),
                        )
                        conn.execute(
                            "DELETE FROM tests WHERE project_id = ? AND id = ?",
                            (source_project, source_test),
                        )
                        moved_any = True
                        pasted.append({
                            "source_project_id": source_project,
                            "source_test_id": source_test,
                            "project_id": dest_project,
                            "test_id": new_id,
                            "mode": "move",
                            "records": cursor.rowcount,
                        })
                    else:
                        pasted.append({
                            "source_project_id": source_project,
                            "source_test_id": source_test,
                            "project_id": dest_project,
                            "test_id": new_id,
                            "mode": "copy",
                            "records": 0,
                        })
                if moved_any:
                    conn.execute("DELETE FROM clipboard WHERE op = 'cut'")
            self._send_json(200, {
                "pasted": pasted,
                "skipped": skipped,
                "clipboard": self._clipboard_payload(conn),
            })
        finally:
            conn.close()

    def _free_test_id(self, conn, project_id: str, preferred: str) -> str:
        exists = conn.execute(
            "SELECT 1 FROM tests WHERE project_id = ? AND id = ?", (project_id, preferred)
        ).fetchone()
        if exists is None and TEST_ID_RE.match(preferred or ""):
            return preferred
        return _next_test_id(conn, project_id)

    def _create_test(self, project_id: str):
        body = self._read_json()
        self._body_read = True
        if not isinstance(body, dict):
            raise ApiError(400, "请求体必须是 JSON 对象")
        conn = connect_db()
        try:
            if project_row(conn, project_id) is None:
                raise ApiError(404, "项目不存在")
            title = _text_field(body, "title", maxlen=200)
            if not title:
                raise ApiError(400, "测试项目名称不能为空")
            test_id = _as_text(body.get("id")).strip()
            if not test_id:
                test_id = _next_test_id(conn, project_id)
            if not TEST_ID_RE.match(test_id):
                raise ApiError(400, "测试编号只能包含字母、数字、点、下划线和短横线（最多 32 字符）")
            if conn.execute(
                "SELECT 1 FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
            ).fetchone():
                raise ApiError(409, f"测试编号 {test_id} 已存在")
            priority = _text_field(body, "priority", maxlen=20, fallback="P0").upper() or "P0"
            module = _text_field(body, "module", maxlen=60, fallback="未分组") or "未分组"
            guidance = {
                field: _text_field(body, field, maxlen=MAX_TEST_TEXT) for field in GUIDANCE_FIELDS
            }
            position = conn.execute(
                "SELECT COALESCE(MAX(position), -1) + 1 FROM tests WHERE project_id = ?",
                (project_id,),
            ).fetchone()[0]
            columns = ["project_id", "id", "priority", "module", "title"] + list(GUIDANCE_FIELDS) + ["position"]
            values = [project_id, test_id, priority, module, title] + [
                guidance[field] for field in GUIDANCE_FIELDS
            ] + [position]
            with conn:
                conn.execute(
                    f"INSERT INTO tests ({', '.join(columns)})"
                    f" VALUES ({', '.join(['?'] * len(columns))})",
                    values,
                )
            row = conn.execute(
                "SELECT * FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
            ).fetchone()
            self._send_json(201, {"test": {field: row[field] for field in TEST_FIELDS}})
        finally:
            conn.close()

    def _delete_test(self, project_id: str, test_id: str):
        conn = connect_db()
        try:
            if project_row(conn, project_id) is None:
                raise ApiError(404, "项目不存在")
            row = conn.execute(
                "SELECT id FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
            ).fetchone()
            if row is None:
                raise ApiError(404, "测试项不存在")
            stored = [
                item["stored_name"]
                for item in conn.execute(
                    "SELECT a.stored_name FROM attachments a"
                    " JOIN records r ON r.id = a.record_id"
                    " WHERE r.project_id = ? AND r.test_id = ?",
                    (project_id, test_id),
                )
            ]
            with conn:
                conn.execute(
                    "DELETE FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
                )
            for name in stored:
                if re.fullmatch(r"[0-9a-f]{32}\.(xlsx|xls|xlsm|jpg|jpeg)", name or ""):
                    try:
                        (CONFIG.uploads_dir / name).unlink()
                    except OSError:
                        pass
            self._send_json(200, {"ok": True, "id": test_id, "attachments": len(stored)})
        finally:
            conn.close()

    def _update_test(self, project_id: str, test_id: str):
        body = self._read_json()
        self._body_read = True
        if not isinstance(body, dict):
            raise ApiError(400, "请求体必须是 JSON 对象")
        updates = {}
        for field in GUIDANCE_FIELDS:
            if field not in body:
                continue
            value = body[field]
            if value is None:
                value = ""
            if not isinstance(value, str):
                raise ApiError(400, f"字段 {field} 必须是字符串")
            value = value.strip()
            if len(value) > MAX_TEST_TEXT:
                raise ApiError(400, f"字段 {field} 超长（最多 {MAX_TEST_TEXT} 字符）")
            updates[field] = value
        if not updates:
            raise ApiError(400, "没有可更新的字段")
        conn = connect_db()
        try:
            row = conn.execute(
                "SELECT * FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
            ).fetchone()
            if row is None:
                raise ApiError(404, "测试项目不存在")
            assignments = ", ".join(f"{field} = ?" for field in updates)
            with conn:
                conn.execute(
                    f"UPDATE tests SET {assignments} WHERE project_id = ? AND id = ?",
                    list(updates.values()) + [project_id, test_id],
                )
            row = conn.execute(
                "SELECT * FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
            ).fetchone()
            self._send_json(200, {"test": {field: row[field] for field in TEST_FIELDS}})
        finally:
            conn.close()

    # ---- record endpoints ------------------------------------------------
    def _list_records(self, project_id: str, test_id: str):
        conn = connect_db()
        try:
            test = conn.execute(
                "SELECT id FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
            ).fetchone()
            if test is None:
                raise ApiError(404, "测试项目不存在")
            rows = conn.execute(
                "SELECT * FROM records WHERE project_id = ? AND test_id = ? ORDER BY id",
                (project_id, test_id),
            ).fetchall()
            records = [_record_dict(conn, row) for row in rows]
            self._send_json(200, {"records": records})
        finally:
            conn.close()

    def _create_record(self, project_id: str, test_id: str):
        body = self._read_json()
        self._body_read = True
        payload = _validate_payload(body)
        timestamp = now_iso()
        conn = connect_db()
        try:
            test = conn.execute(
                "SELECT id FROM tests WHERE project_id = ? AND id = ?", (project_id, test_id)
            ).fetchone()
            if test is None:
                raise ApiError(404, "测试项目不存在")
            with conn:
                cursor = conn.execute(
                    "INSERT INTO records (project_id, test_id, mode, result, unit, status,"
                    " test_date, notes, created_at, updated_at)"
                    " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (project_id, test_id, payload["mode"], payload["result"], payload["unit"],
                     payload["status"], payload["test_date"], payload["notes"],
                     timestamp, timestamp),
                )
                record_id = cursor.lastrowid
            row = conn.execute("SELECT * FROM records WHERE id = ?", (record_id,)).fetchone()
            self._send_json(201, {"record": _record_dict(conn, row, [])})
        finally:
            conn.close()

    def _update_record(self, record_id: int):
        body = self._read_json()
        self._body_read = True
        conn = connect_db()
        try:
            row = conn.execute("SELECT * FROM records WHERE id = ?", (record_id,)).fetchone()
            if row is None:
                raise ApiError(404, "记录不存在")
            payload = _validate_payload(body, existing=row)
            with conn:
                conn.execute(
                    "UPDATE records SET mode = ?, result = ?, unit = ?, status = ?,"
                    " test_date = ?, notes = ?, updated_at = ? WHERE id = ?",
                    (payload["mode"], payload["result"], payload["unit"],
                     payload["status"], payload["test_date"], payload["notes"],
                     now_iso(), record_id),
                )
            row = conn.execute("SELECT * FROM records WHERE id = ?", (record_id,)).fetchone()
            self._send_json(200, {"record": _record_dict(conn, row)})
        finally:
            conn.close()

    # ---- attachment endpoints -------------------------------------------
    def _upload_attachments(self, record_id: int):
        conn = connect_db()
        written: list[Path] = []
        try:
            record = conn.execute("SELECT id FROM records WHERE id = ?", (record_id,)).fetchone()
            if record is None:
                raise ApiError(404, "记录不存在")
            ctype = self.headers.get("Content-Type", "")
            if not ctype.lower().startswith("multipart/form-data"):
                raise ApiError(415, "内容类型必须为 multipart/form-data")
            body = self._read_body(MAX_REQUEST_BYTES)
            self._body_read = True
            fields, files = parse_multipart(self.headers, body)
            if not files:
                raise ApiError(400, "请选择至少一个附件文件")
            kind = (fields.get("kind") or "data").strip()
            if kind not in ("data", "image"):
                raise ApiError(400, "kind 必须为 data 或 image")
            prepared = []
            for _field, filename, payload in files:
                display_name, ext = _validate_upload(kind, filename, payload)
                prepared.append((display_name, ext, payload))

            timestamp = now_iso()
            inserted_ids = []
            for display_name, ext, payload in prepared:
                stored_name = uuid.uuid4().hex + ext
                target = CONFIG.uploads_dir / stored_name
                written.append(target)
                target.write_bytes(payload)
                cursor = conn.execute(
                    "INSERT INTO attachments (record_id, kind, name, size, stored_name, created_at)"
                    " VALUES (?, ?, ?, ?, ?, ?)",
                    (record_id, kind, display_name, len(payload), stored_name, timestamp),
                )
                inserted_ids.append(cursor.lastrowid)
            conn.execute("UPDATE records SET updated_at = ? WHERE id = ?", (timestamp, record_id))
            conn.commit()
            rows = [
                conn.execute("SELECT * FROM attachments WHERE id = ?", (aid,)).fetchone()
                for aid in inserted_ids
            ]
            self._send_json(200, {"attachments": [_attachment_dict(r) for r in rows]})
        except Exception:
            conn.rollback()
            for path in written:
                try:
                    path.unlink()
                except OSError:
                    pass
            raise
        finally:
            conn.close()

    def _download_attachment(self, attachment_id: int):
        conn = connect_db()
        try:
            row = conn.execute("SELECT * FROM attachments WHERE id = ?", (attachment_id,)).fetchone()
            if row is None:
                raise ApiError(404, "附件不存在")
            stored = row["stored_name"]
            if not re.fullmatch(r"[0-9a-f]{32}\.(xlsx|xls|xlsm|jpg|jpeg)", stored or ""):
                raise ApiError(404, "附件不存在")
            path = CONFIG.uploads_dir / stored
            try:
                resolved = path.resolve(strict=True)
            except (OSError, RuntimeError):
                raise ApiError(404, "附件文件缺失") from None
            if not resolved.is_relative_to(CONFIG.uploads_dir.resolve()) or not resolved.is_file():
                raise ApiError(404, "附件文件缺失")
            size = resolved.stat().st_size
            if row["kind"] == "image":
                content_type = "image/jpeg"
                inline = True
            else:
                content_type = "application/octet-stream"
                inline = False
            headers = {
                "Content-Disposition": _content_disposition(row["name"], inline=inline),
                "Content-Security-Policy": "default-src 'none'; sandbox",
                "Cache-Control": "private, max-age=0, no-store",
                "Content-Length": str(size),
            }
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            for key, value in headers.items():
                self.send_header(key, value)
            self._base_headers()
            self.end_headers()
            if self.command == "HEAD":
                return
            try:
                with resolved.open("rb") as fh:
                    shutil.copyfileobj(fh, self.wfile, 256 * 1024)
            except (BrokenPipeError, ConnectionResetError, OSError):
                self.close_connection = True
        finally:
            conn.close()

    def _delete_attachment(self, attachment_id: int):
        conn = connect_db()
        try:
            row = conn.execute("SELECT * FROM attachments WHERE id = ?", (attachment_id,)).fetchone()
            if row is None:
                raise ApiError(404, "附件不存在")
            with conn:
                conn.execute("DELETE FROM attachments WHERE id = ?", (attachment_id,))
                conn.execute("UPDATE records SET updated_at = ? WHERE id = ?",
                             (now_iso(), row["record_id"]))
            stored = row["stored_name"]
            if re.fullmatch(r"[0-9a-f]{32}\.(xlsx|xls|xlsm|jpg|jpeg)", stored or ""):
                try:
                    (CONFIG.uploads_dir / stored).unlink()
                except OSError:
                    pass
            self._send_json(200, {"ok": True})
        finally:
            conn.close()

    # ---- export / backup -------------------------------------------------
    def _export_xlsx(self, project_id: str | None = None):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Font, PatternFill
            from openpyxl.utils import get_column_letter
        except ImportError:
            raise ApiError(503, "服务器未安装 openpyxl，无法导出 XLSX") from None

        conn = connect_db()
        try:
            project_name = ""
            if project_id is not None:
                project = project_row(conn, project_id)
                if project is None:
                    raise ApiError(404, "项目不存在")
                project_name = project["name"]
            base_sql = (
                "SELECT r.*, p.name AS project_name, t.priority, t.module, t.title,"
                " (SELECT COUNT(*) FROM attachments a WHERE a.record_id = r.id) AS attachment_count"
                " FROM records r"
                " JOIN tests t ON t.project_id = r.project_id AND t.id = r.test_id"
                " JOIN projects p ON p.id = r.project_id"
            )
            if project_id is None:
                rows = conn.execute(base_sql + " ORDER BY p.position, p.id, r.id").fetchall()
            else:
                rows = conn.execute(base_sql + " WHERE r.project_id = ? ORDER BY r.id",
                                    (project_id,)).fetchall()
        finally:
            conn.close()

        wb = Workbook()
        ws = wb.active
        ws.title = (project_name or "测试记录")[:31]
        headers = ["记录ID", "测试编号", "优先级", "模块", "测试项目", "模式", "结果", "单位",
                   "状态", "测试日期", "备注", "附件数", "创建时间", "更新时间"]
        if project_id is None:
            headers.insert(1, "项目")
        ws.append(headers)
        fill = PatternFill("solid", fgColor="1E3A5F")
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for record in rows:
            values = [
                record["id"], record["test_id"], record["priority"], record["module"],
                record["title"],
                "仿真" if record["mode"] == "simulation" else "实测",
                record["result"], record["unit"], record["status"], record["test_date"],
                record["notes"], record["attachment_count"],
                record["created_at"], record["updated_at"],
            ]
            if project_id is None:
                values.insert(1, record["project_name"])
            ws.append(values)
        widths = [8, 14, 8, 18, 42, 8, 16, 10, 8, 12, 40, 8, 22, 22]
        if project_id is None:
            widths.insert(1, 20)
        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        buffer = io.BytesIO()
        wb.save(buffer)
        slug = _project_slug(project_name) if project_name else "all"
        filename = f"testflow-{slug}-records-{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx"
        self._send(200, buffer.getvalue(),
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   {"Content-Disposition": _content_disposition(filename),
                    "Cache-Control": "no-store"})

    def _backup_zip(self):
        # NOTE: tempfile.mkdtemp() uses mode 0o700, which this Windows host maps
        # to an ACL that denies access. Use a normally-created work directory.
        temp_dir = CONFIG.data_dir / f".backup-{uuid.uuid4().hex}"
        temp_dir.mkdir()
        temp_zip = temp_dir / "backup.zip"
        try:
            snapshot = temp_dir / "afe-testflow.db"
            source = connect_db()
            try:
                target = sqlite3.connect(str(snapshot))
                try:
                    source.backup(target)
                finally:
                    target.close()
            finally:
                source.close()

            conn = connect_db()
            try:
                project_count = conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
                record_count = conn.execute("SELECT COUNT(*) FROM records").fetchone()[0]
            finally:
                conn.close()
            info = {
                "created_at": now_iso(),
                "database": "data/afe-testflow.db",
                "uploads": "data/uploads/",
                "seed": "seed.json" if CONFIG.seed_path.is_file() else None,
                "projects": project_count,
                "records": record_count,
            }
            with zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(snapshot, "data/afe-testflow.db")
                if CONFIG.seed_path.is_file():
                    zf.write(CONFIG.seed_path, "seed.json")
                uploads_root = CONFIG.uploads_dir.resolve()
                if uploads_root.is_dir():
                    for item in sorted(uploads_root.iterdir()):
                        if item.is_file():
                            try:
                                resolved = item.resolve(strict=True)
                            except OSError:
                                continue
                            if resolved.is_relative_to(uploads_root):
                                zf.write(resolved, f"data/uploads/{resolved.name}")
                zf.writestr("backup-info.json", json.dumps(info, ensure_ascii=False, indent=2))

            filename = f"testflow-backup-{datetime.now().strftime('%Y%m%d-%H%M%S')}.zip"
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Length", str(temp_zip.stat().st_size))
            self.send_header("Content-Disposition", _content_disposition(filename))
            self.send_header("Cache-Control", "no-store")
            self._base_headers()
            self.end_headers()
            if self.command == "HEAD":
                return
            try:
                with temp_zip.open("rb") as fh:
                    shutil.copyfileobj(fh, self.wfile, 256 * 1024)
            except (BrokenPipeError, ConnectionResetError, OSError):
                self.close_connection = True
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    # ---- static files ----------------------------------------------------
    def _static(self, path: str):
        public_root = CONFIG.public_dir.resolve()
        if path in ("", "/"):
            rel = "index.html"
        else:
            try:
                decoded = unquote(path, errors="strict")
            except UnicodeDecodeError:
                raise ApiError(400, "路径无效") from None
            if "\x00" in decoded:
                raise ApiError(400, "路径无效")
            rel = decoded.lstrip("/")
        normalized = posixpath.normpath(rel)
        if normalized in ("", ".") or normalized.startswith("../") or normalized == ".." \
                or normalized.startswith("/"):
            raise ApiError(404, "资源不存在")
        candidate = (public_root / normalized)
        try:
            resolved = candidate.resolve(strict=True)
        except (OSError, RuntimeError):
            if "." not in Path(normalized).name:
                index = public_root / "index.html"
                if index.is_file():
                    return self._send_static_file(index)
            raise ApiError(404, "资源不存在") from None
        if not resolved.is_relative_to(public_root) or not resolved.is_file():
            raise ApiError(404, "资源不存在")
        return self._send_static_file(resolved)

    def _send_static_file(self, path: Path):
        try:
            body = path.read_bytes()
        except OSError:
            raise ApiError(404, "资源不存在") from None
        suffix = path.suffix.lower()
        content_type = STATIC_TYPES.get(suffix) or mimetypes.guess_type(path.name)[0] \
            or "application/octet-stream"
        headers = {"Cache-Control": "no-cache"}
        if suffix == ".html":
            headers["Content-Security-Policy"] = (
                "default-src 'self'; img-src 'self' data: blob:; "
                "style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; "
                "connect-src 'self'; object-src 'none'; base-uri 'none'"
            )
        self._send(200, body, content_type, headers)

    # ---- unsupported methods --------------------------------------------
    def _method_not_allowed(self):
        self._error(405, "方法不允许")

    do_PATCH = _method_not_allowed
    do_OPTIONS = _method_not_allowed


class TestFlowServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def serve(host: str | None = None, port: int | None = None) -> None:
    init_storage()
    init_db()
    seed_default_project()
    host = host or CONFIG.host
    port = CONFIG.port if port is None else port
    configure_allowed_hosts(host)
    httpd = TestFlowServer((host, port), Handler)
    display_host = "127.0.0.1" if host in ("0.0.0.0", "") else host
    print(f"TestFlow backend listening on http://{display_host}:{port}", flush=True)
    if host in ("0.0.0.0", "::", ""):
        lan = sorted(h for h in ALLOWED_HOSTS if h not in LOCAL_HOSTS)
        print("  reachable as: " + ", ".join(f"http://{h}:{port}" for h in lan)
              if lan else "  reachable as: (loopback only)", flush=True)
    print("  allowed Host headers: " + ", ".join(sorted(ALLOWED_HOSTS)), flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nstopping...", flush=True)
    finally:
        httpd.shutdown()
        httpd.server_close()


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="AFE TestFlow local backend")
    parser.add_argument("--host", default=CONFIG.host)
    parser.add_argument("--port", type=int, default=CONFIG.port)
    parser.add_argument("--data-dir", default=None,
                        help="override data directory (SQLite + uploads); useful for tests")
    parser.add_argument("--seed", default=None, help="override seed.json path")
    args = parser.parse_args(argv)
    if args.data_dir:
        CONFIG.set_data_dir(args.data_dir)
    if args.seed:
        CONFIG.set_seed_path(args.seed)
    CONFIG.host = args.host
    CONFIG.port = args.port
    try:
        serve(args.host, args.port)
    except OSError as exc:
        print(f"无法启动服务：{exc}", file=sys.stderr, flush=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
