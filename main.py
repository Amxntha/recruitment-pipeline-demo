# ═══════════════════════════════════════════════════════════════════════════════
# HELP AT HAND SUPPORT - Recruitment Pipeline API
# ═══════════════════════════════════════════════════════════════════════════════
# FastAPI app for managing recruitment candidates and receiving JotForm webhooks.
# In this pure SPA design, FastAPI saves submissions locally while the browser
# handles Microsoft sign-in and calls Microsoft Graph directly.

# ───────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ───────────────────────────────────────────────────────────────────────────────

import ast
import base64
import json
import mimetypes
import os
import re
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Optional
from urllib.parse import quote, unquote, urlparse

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

load_dotenv()

# ───────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ───────────────────────────────────────────────────────────────────────────────

DB_PATH = os.getenv("DB_PATH", "hahs.db")
STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./local_storage")

# Pure SPA / MSAL.js public-client settings. These values are sent to the
# browser because the frontend performs Microsoft login and Graph API calls.
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", os.getenv("CLIENT_ID", "")).strip()
MS_TENANT_ID = os.getenv("MS_TENANT_ID", os.getenv("TENANT_ID", "")).strip()
GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Default SharePoint folder suggestion only. The actual SharePoint site and
# document library are selected by the user in the Integration tab and stored
# in SQLite, not hardcoded in .env.
DEFAULT_BASE_FOLDER = os.getenv("DEFAULT_BASE_FOLDER", "HR Demo Candidate Log").strip().strip("/")

# Optional JotForm API fallback. This is used only when the webhook file URL
# returns HTML instead of real file bytes. The API key stays server-side.
JOTFORM_API_KEY = os.getenv("JOTFORM_API_KEY", "").strip()
JOTFORM_API_BASE = os.getenv("JOTFORM_API_BASE", "https://api.jotform.com").rstrip("/")

app = FastAPI(title="HelpAtHandSupport API")

# ───────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ───────────────────────────────────────────────────────────────────────────────

VALID_ROLES = {"Support Worker", "Client"}
VALID_STAGES = {
    "Shortlisted",
    "Screening Call",
    "Documents Requested",
    "Documents Received",
    "Interviewed",
    "Hired",
    "Declined",
}

CANDIDATE_FIELDS = [
    "name",
    "role",
    "skills",
    "stage",
    "date",
    "mobile_number",
    "email",
    "state",
    "car_registration",
    "ndis_worker_check",
    "police_check",
    "working_with_children",
    "id_100_points",
    "first_aid_cpr",
    "ndis_orientation",
    "covid_training",
    "car_insurance",
    "car_rego_proof",
    "face_id_picture",
    "certificates_study",
    "signature",
    "captcha_passed",
    "confirmation_agreed",
    "reference_1",
    "reference_2",
    "local_folder_path",
    "sharepoint_sync_status",
    "sharepoint_folder_id",
    "sharepoint_folder_url",
    "sharepoint_folder_path",
]

REQUIRED_CANDIDATE_FIELDS = {"name", "role", "stage"}
CANDIDATE_DB_COLUMNS = {
    field: "TEXT NOT NULL" if field in REQUIRED_CANDIDATE_FIELDS else "TEXT"
    for field in CANDIDATE_FIELDS
}

DOCUMENT_UPLOAD_FIELDS = {
    "ndis_worker_check": "fileUpload",
    "police_check": "policeCheck",
    "working_with_children": "workingWith",
    "id_100_points": "100Points",
    "first_aid_cpr": "firstAidcpr",
    "ndis_orientation": "ndisWorker",
    "covid_training": "covid19Training",
    "car_insurance": "evidenceOf",
    "car_rego_proof": "evidenceOf16",
    "face_id_picture": "pictureOf",
    "certificates_study": "certificatesOf",
    "signature": "q8_iConfirm",
}

DOCUMENT_DISPLAY_LABELS = {
    "ndis_worker_check": "NDIS worker check",
    "police_check": "Police check",
    "working_with_children": "Working with children",
    "id_100_points": "100 points ID",
    "first_aid_cpr": "First aid / CPR",
    "ndis_orientation": "NDIS orientation",
    "covid_training": "COVID training",
    "car_insurance": "Car insurance",
    "car_rego_proof": "Car registration proof",
    "face_id_picture": "Face ID picture",
    "certificates_study": "Certificates / study",
    "signature": "Signature",
}

PREVIEWABLE_CONTENT_TYPES = {
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/webp",
    "image/svg+xml",
    "application/pdf",
    "text/plain",
}


class FileDownloadError(ValueError):
    """Raised when a source URL does not return usable file bytes."""

    def __init__(self, message: str, *, returned_html: bool = False):
        super().__init__(message)
        self.returned_html = returned_html


INTEGRATION_ID = "default"

# ───────────────────────────────────────────────────────────────────────────────
# DATABASE HELPERS
# ───────────────────────────────────────────────────────────────────────────────

@contextmanager
def get_conn():
    """Context manager for SQLite connections."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def row_to_dict(row: sqlite3.Row) -> dict:
    """Convert SQLite Row object to dictionary."""
    return dict(row)


def get_existing_columns(conn: sqlite3.Connection, table_name: str) -> set[str]:
    """Return existing column names for a table."""
    return {
        row["name"]
        for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    }


def init_db() -> None:
    """Create database tables and add missing columns."""
    candidate_columns_sql = ",\n                ".join(
        f"{column} {column_type}"
        for column, column_type in CANDIDATE_DB_COLUMNS.items()
    )

    with get_conn() as conn:

        conn.execute(f"""
            CREATE TABLE IF NOT EXISTS candidates (
                id TEXT PRIMARY KEY,
                {candidate_columns_sql},
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)

        existing_cols = get_existing_columns(conn, "candidates")
        for column, column_type in CANDIDATE_DB_COLUMNS.items():
            if column not in existing_cols:
                nullable_type = "TEXT" if "NOT NULL" in column_type else column_type
                conn.execute(f"ALTER TABLE candidates ADD COLUMN {column} {nullable_type}")

        conn.execute("""
            CREATE TABLE IF NOT EXISTS document_requests (
                id TEXT PRIMARY KEY,
                first_name TEXT NOT NULL,
                last_name TEXT NOT NULL,
                mobile_number TEXT,
                email TEXT,
                state TEXT,
                car_registration TEXT,
                ndis_worker_check TEXT,
                police_check TEXT,
                working_with_children TEXT,
                id_100_points TEXT,
                first_aid_cpr TEXT,
                ndis_orientation TEXT,
                covid_training TEXT,
                car_insurance TEXT,
                car_rego_proof TEXT,
                face_id_picture TEXT,
                certificates_study TEXT,
                signature TEXT,
                captcha_passed TEXT,
                confirmation_agreed TEXT,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS document_references (
                id TEXT PRIMARY KEY,
                document_request_id TEXT NOT NULL,
                first_name TEXT,
                last_name TEXT,
                phone_number TEXT,
                email TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(document_request_id) REFERENCES document_requests(id)
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS sharepoint_destination (
                id TEXT PRIMARY KEY,
                site_id TEXT NOT NULL,
                site_name TEXT,
                site_url TEXT,
                drive_id TEXT NOT NULL,
                drive_name TEXT,
                drive_url TEXT,
                base_folder TEXT,
                selected_at TEXT DEFAULT (datetime('now'))
            )
        """)

# ───────────────────────────────────────────────────────────────────────────────
# PYDANTIC MODELS
# ───────────────────────────────────────────────────────────────────────────────

class CandidateIn(BaseModel):
    """Candidate input model for creating and updating candidates."""
    name: str
    role: str
    skills: Optional[str] = ""
    stage: str
    date: Optional[str] = None
    mobile_number: Optional[str] = None
    email: Optional[str] = None
    state: Optional[str] = None
    car_registration: Optional[str] = None
    ndis_worker_check: Optional[str] = None
    police_check: Optional[str] = None
    working_with_children: Optional[str] = None
    id_100_points: Optional[str] = None
    first_aid_cpr: Optional[str] = None
    ndis_orientation: Optional[str] = None
    covid_training: Optional[str] = None
    car_insurance: Optional[str] = None
    car_rego_proof: Optional[str] = None
    face_id_picture: Optional[str] = None
    certificates_study: Optional[str] = None
    signature: Optional[str] = None
    captcha_passed: Optional[str] = None
    confirmation_agreed: Optional[str] = None
    reference_1: Optional[str] = None
    reference_2: Optional[str] = None
    local_folder_path: Optional[str] = None
    sharepoint_sync_status: Optional[str] = None
    sharepoint_folder_id: Optional[str] = None
    sharepoint_folder_url: Optional[str] = None
    sharepoint_folder_path: Optional[str] = None


class CandidateOut(CandidateIn):
    """Candidate output model including database metadata."""
    id: str
    created_at: str


class SharePointSyncIn(BaseModel):
    """Payload sent by the browser after it syncs local files to SharePoint."""
    folder_id: Optional[str] = None
    folder_url: Optional[str] = None
    folder_path: Optional[str] = None
    status: Optional[str] = "Synced"


class SharePointDestinationIn(BaseModel):
    """SharePoint destination selected by the user from the Integration tab."""
    site_id: str
    site_name: Optional[str] = ""
    site_url: Optional[str] = ""
    drive_id: str
    drive_name: Optional[str] = ""
    drive_url: Optional[str] = ""
    base_folder: Optional[str] = ""

# ───────────────────────────────────────────────────────────────────────────────
# PURE SPA / MICROSOFT GRAPH CONFIGURATION
# ───────────────────────────────────────────────────────────────────────────────

def get_sharepoint_destination() -> Optional[dict]:
    """Return the saved SharePoint site/library/folder destination, if selected."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM sharepoint_destination WHERE id = ?",
            (INTEGRATION_ID,),
        ).fetchone()
    return row_to_dict(row) if row else None


def destination_display_path(destination: Optional[dict]) -> str:
    """Return a user-friendly path without exposing Graph IDs in the UI."""
    if not destination:
        return ""
    parts = [
        destination.get("site_name") or "Selected site",
        destination.get("drive_name") or "Selected library",
    ]
    base_folder = (destination.get("base_folder") or "").strip().strip("/")
    if base_folder:
        parts.append(base_folder)
    return " / ".join(parts)


@app.get("/api/spa-config")
def spa_config():
    """Return public Microsoft config for the browser SPA.

    There is intentionally no client secret here. The frontend uses MSAL.js with
    PKCE and calls Microsoft Graph directly from the browser. The SharePoint
    drive is selected from the Integration tab and stored separately.
    """
    destination = get_sharepoint_destination()
    return {
        "mode": "pure_spa_pkce",
        "auth": {
            "client_id": MS_CLIENT_ID,
            "tenant_id": MS_TENANT_ID,
            "authority": f"https://login.microsoftonline.com/{MS_TENANT_ID}" if MS_TENANT_ID else "",
        },
        "graph": {
            "base_url": GRAPH_BASE_URL,
            "scopes": GRAPH_SCOPES,
            "default_base_folder": DEFAULT_BASE_FOLDER,
        },
        "destination": {
            "configured": bool(destination),
            "site_name": destination.get("site_name", "") if destination else "",
            "site_url": destination.get("site_url", "") if destination else "",
            "drive_name": destination.get("drive_name", "") if destination else "",
            "drive_url": destination.get("drive_url", "") if destination else "",
            "base_folder": destination.get("base_folder", "") if destination else "",
            "display_path": destination_display_path(destination),
        },
        "configured": bool(MS_CLIENT_ID and MS_TENANT_ID),
    }


@app.get("/api/sharepoint-destination")
def read_sharepoint_destination():
    """Return the saved SharePoint destination. IDs are returned for browser Graph calls,
    but the frontend displays only the friendly site/library/folder path.
    """
    
    destination = get_sharepoint_destination()
    if not destination:
    
        return {
            "configured": False,
            "site_id": "",
            "site_name": "",
            "site_url": "",
            "drive_id": "",
            "drive_name": "",
            "drive_url": "",
            "base_folder": DEFAULT_BASE_FOLDER,
            "display_path": "",
        }
    
    
    return {
        "configured": True,
        "site_id": destination.get("site_id") or "",
        "site_name": destination.get("site_name") or "",
        "site_url": destination.get("site_url") or "",
        "drive_id": destination.get("drive_id") or "",
        "drive_name": destination.get("drive_name") or "",
        "drive_url": destination.get("drive_url") or "",
        "base_folder": destination.get("base_folder") or "",
        "display_path": destination_display_path(destination),
        "selected_at": destination.get("selected_at") or "",
    }


@app.post("/api/sharepoint-destination")
def save_sharepoint_destination(data: SharePointDestinationIn):
    """Save the SharePoint site, document library and base folder selected by the user."""
    site_id = data.site_id.strip()
    drive_id = data.drive_id.strip()
    base_folder = (data.base_folder or "").strip().strip("/")

    if not site_id:
        raise HTTPException(400, "SharePoint site is required")
    if not drive_id:
        raise HTTPException(400, "SharePoint document library is required")

    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO sharepoint_destination (
                id, site_id, site_name, site_url, drive_id, drive_name, drive_url, base_folder, selected_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(id) DO UPDATE SET
                site_id = excluded.site_id,
                site_name = excluded.site_name,
                site_url = excluded.site_url,
                drive_id = excluded.drive_id,
                drive_name = excluded.drive_name,
                drive_url = excluded.drive_url,
                base_folder = excluded.base_folder,
                selected_at = datetime('now')
            """,
            (
                INTEGRATION_ID,
                site_id,
                (data.site_name or "").strip(),
                (data.site_url or "").strip(),
                drive_id,
                (data.drive_name or "").strip(),
                (data.drive_url or "").strip(),
                base_folder,
            ),
        )

    destination = get_sharepoint_destination()
  
    return {
        "message": "SharePoint destination saved.",
        "configured": True,
        "site_name": destination.get("site_name") or "",
        "site_url": destination.get("site_url") or "",
        "drive_name": destination.get("drive_name") or "",
        "drive_url": destination.get("drive_url") or "",
        "base_folder": destination.get("base_folder") or "",
        "display_path": destination_display_path(destination),
    }


@app.get("/api/integration/microsoft/status")
def microsoft_integration_status():
    """Compatibility endpoint used by the Integration tab.

    In pure SPA mode, connection status is checked in the browser because the
    browser owns the Microsoft account session and token cache.
    """
    destination = get_sharepoint_destination()
    return {
        "mode": "pure_spa_pkce",
        "configured": bool(MS_CLIENT_ID and MS_TENANT_ID),
        "destination_configured": bool(destination),
        "connected": False,
        "client_id": MS_CLIENT_ID,
        "tenant_id": MS_TENANT_ID,
        "destination": {
            "site_name": destination.get("site_name", "") if destination else "",
            "site_url": destination.get("site_url", "") if destination else "",
            "drive_name": destination.get("drive_name", "") if destination else "",
            "drive_url": destination.get("drive_url", "") if destination else "",
            "base_folder": destination.get("base_folder", "") if destination else "",
            "display_path": destination_display_path(destination),
        },
        "message": "Microsoft sign-in and Graph API calls are handled by the browser using MSAL.js and PKCE.",
    }


@app.get("/api/integration/microsoft/connect")
def connect_microsoft_deprecated():
    """This backend endpoint is not used in pure SPA mode."""
    raise HTTPException(
        410,
        "Pure SPA mode uses the frontend Connect button with MSAL.js; the backend does not perform Microsoft sign-in.",
    )

# ───────────────────────────────────────────────────────────────────────────────
# CANDIDATE HELPERS
# ───────────────────────────────────────────────────────────────────────────────

def validate_candidate(data: CandidateIn) -> None:
    """Validate candidate input data."""
    if not data.name or not data.name.strip():
        raise HTTPException(400, "Name is required")
    if data.role not in VALID_ROLES:
        raise HTTPException(400, f"Role must be one of: {sorted(VALID_ROLES)}")
    if data.stage not in VALID_STAGES:
        raise HTTPException(400, f"Stage must be one of: {sorted(VALID_STAGES)}")


def normalise_candidate_payload(data: CandidateIn) -> dict:
    """Return a clean dictionary containing only database candidate fields."""
    payload = data.dict()
    payload["name"] = payload.get("name", "").strip()
    return {field: payload.get(field) or "" for field in CANDIDATE_FIELDS}


def insert_candidate_record(conn: sqlite3.Connection, data: CandidateIn) -> dict:
    """Insert a candidate and return the created row as a dictionary."""
    validate_candidate(data)
    cid = str(uuid.uuid4())
    payload = normalise_candidate_payload(data)

    columns = ["id", *CANDIDATE_FIELDS]
    placeholders = ",".join("?" for _ in columns)
    values = [cid, *(payload[field] for field in CANDIDATE_FIELDS)]

    conn.execute(
        f"INSERT INTO candidates ({','.join(columns)}) VALUES ({placeholders})",
        values,
    )
    row = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()
    return row_to_dict(row)


def update_candidate_record(
    conn: sqlite3.Connection,
    cid: str,
    data: CandidateIn,
    *,
    preserve_existing_values: bool,
) -> dict:
    """Update a candidate using one shared merge/update path."""
    validate_candidate(data)
    existing_row = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()
    if not existing_row:
        raise HTTPException(404, "Candidate not found")

    existing = row_to_dict(existing_row)
    incoming = normalise_candidate_payload(data)

    merged = {}
    for field in CANDIDATE_FIELDS:
        if preserve_existing_values:
            merged[field] = existing.get(field) or incoming.get(field, "")
        else:
            merged[field] = incoming.get(field) or existing.get(field, "")

    assignments = ", ".join(f"{field} = ?" for field in CANDIDATE_FIELDS)
    values = [merged[field] for field in CANDIDATE_FIELDS] + [cid]

    conn.execute(f"UPDATE candidates SET {assignments} WHERE id = ?", values)
    row = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()
    return row_to_dict(row)


def find_existing_candidate(
    conn: sqlite3.Connection,
    email: str,
    full_name: str,
) -> Optional[sqlite3.Row]:
    """Find an existing candidate by email first, then by full name."""
    if email:
        candidate = conn.execute("SELECT id FROM candidates WHERE email = ?", (email,)).fetchone()
        if candidate:
            return candidate
    if full_name:
        return conn.execute("SELECT id FROM candidates WHERE name = ?", (full_name,)).fetchone()
    return None

# ───────────────────────────────────────────────────────────────────────────────
# JOTFORM PARSING HELPERS
# ───────────────────────────────────────────────────────────────────────────────

def parse_raw_request(data: dict) -> dict:
    """Parse the rawRequest field from a JotForm webhook payload."""
    if not data.get("rawRequest"):
        return {}
    try:
        raw = json.loads(data["rawRequest"])
        print(f"[WEBHOOK DEBUG] rawRequest parsed, keys={list(raw.keys())}")
        return raw
    except (json.JSONDecodeError, TypeError, ValueError) as exc:
        print(f"[WEBHOOK DEBUG] rawRequest parse error: {exc}")
        return {}


def parse_jotform_image_url(raw_value: Any) -> Optional[str]:
    """Extract image/file URLs from common JotForm structures."""
    if not raw_value:
        return None

    if isinstance(raw_value, dict):
        for key in ("url", "href", "link", "src", "file"):
            if raw_value.get(key):
                return raw_value[key]
        return None

    if isinstance(raw_value, list):
        return parse_jotform_image_url(raw_value[0]) if raw_value else None

    if isinstance(raw_value, str):
        text = raw_value.strip()
        if not text:
            return None
        if text.startswith("http") or text.startswith("data:image/"):
            return text
        try:
            parsed = ast.literal_eval(text)
            return parse_jotform_image_url(parsed)
        except (ValueError, SyntaxError):
            match = re.search(r"https?://[^\s'\"\[\]]+", text)
            return match.group(0) if match else None

    return None


def extract_name(raw: dict) -> tuple[str, str, str]:
    """Extract first, last, and full name from JotForm raw data."""
    name_obj = raw.get("q1_name", {})
    first_name = name_obj.get("first", "") if isinstance(name_obj, dict) else ""
    last_name = name_obj.get("last", "") if isinstance(name_obj, dict) else ""
    full_name = f"{first_name} {last_name}".strip()
    return first_name, last_name, full_name


def extract_phone(raw: dict, key: str) -> str:
    """Extract a JotForm phone field."""
    phone_obj = raw.get(key, {})
    return phone_obj.get("full", "") if isinstance(phone_obj, dict) else ""


def parse_submit_date(raw: dict) -> Optional[str]:
    """Parse JotForm submitDate timestamp into YYYY-MM-DD."""
    if not raw.get("submitDate"):
        return None
    try:
        timestamp_ms = int(raw["submitDate"])
        return datetime.fromtimestamp(timestamp_ms / 1000).strftime("%Y-%m-%d")
    except (ValueError, TypeError) as exc:
        print(f"[WEBHOOK DEBUG] submitDate parse error: {exc}")
        return None


def build_reference(
    raw: dict,
    name_key: str,
    phone_key: str,
    email_key: str,
) -> Optional[str]:
    """Convert JotForm reference fields into a compact JSON string."""
    name_obj = raw.get(name_key, {})
    first = name_obj.get("first", "") if isinstance(name_obj, dict) else ""
    last = name_obj.get("last", "") if isinstance(name_obj, dict) else ""
    phone = extract_phone(raw, phone_key)
    email = raw.get(email_key, "")

    reference = {}
    if first or last:
        reference["name"] = f"{first} {last}".strip()
    if phone:
        reference["phone"] = phone
    if email:
        reference["email"] = email
    return json.dumps(reference) if reference else None


def build_candidate_from_jotform(raw: dict) -> tuple[CandidateIn, dict]:
    """Build CandidateIn and document URL dictionary from JotForm raw data."""
    first_name, last_name, full_name = extract_name(raw)
    if not first_name or not last_name:
        raise HTTPException(400, "Missing required fields: first name and last name")

    document_urls = {
        candidate_field: parse_jotform_image_url(raw.get(jotform_key, ""))
        for candidate_field, jotform_key in DOCUMENT_UPLOAD_FIELDS.items()
    }

    candidate = CandidateIn(
        name=full_name,
        role="Support Worker",
        skills="N/A",
        stage="Documents Received",
        date=parse_submit_date(raw),
        mobile_number=extract_phone(raw, "q26_mobileNumber"),
        email=raw.get("q2_email", ""),
        state=raw.get("q46_state", ""),
        car_registration=raw.get("q18_carRegistration", ""),
        confirmation_agreed=raw.get("q13_iConfirm", ""),
        reference_1=build_reference(raw, "q38_name38", "q39_phoneNumber", "q44_email44"),
        reference_2=build_reference(raw, "q41_name41", "q40_phoneNumber40", "q45_email45"),
        **{field: str(url) if url else None for field, url in document_urls.items()},
    )

    return candidate, document_urls

# ───────────────────────────────────────────────────────────────────────────────
# STORAGE / SHAREPOINT HELPERS
# ───────────────────────────────────────────────────────────────────────────────

def safe_json(response: requests.Response) -> Any:
    """Return JSON if available, otherwise response text."""
    try:
        return response.json()
    except ValueError:
        return response.text


def sanitize_filename(name: str) -> str:
    """Create a filesystem-safe name from a user-provided value."""
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._-")
    return cleaned or "submission"



def jotform_api_get(path: str, params: Optional[dict] = None) -> dict:
    """Call the JotForm API using the server-side API key.

    The API is used as a fallback source of submission/file metadata when the
    direct file URL from the webhook returns HTML instead of file bytes.
    """
    if not JOTFORM_API_KEY:
        raise FileDownloadError(
            "JotForm API fallback is unavailable because JOTFORM_API_KEY is not configured."
        )

    api_path = path if path.startswith("/") else f"/{path}"
    response = requests.get(
        f"{JOTFORM_API_BASE}{api_path}",
        params=params or {},
        timeout=90,
        headers={
            "APIKEY": JOTFORM_API_KEY,
            "Accept": "application/json",
            "User-Agent": "HelpAtHandSupport-RecruitmentPipeline/1.0",
        },
    )
    response.raise_for_status()
    return response.json()


def extract_urls_from_any(value: Any) -> list[str]:
    """Recursively collect HTTP/HTTPS URLs from strings, lists and dicts."""
    urls: list[str] = []

    if value is None:
        return urls

    if isinstance(value, str):
        urls.extend(re.findall(r"https?://[^\s'\"<>\]]+", value))
        return urls

    if isinstance(value, list):
        for item in value:
            urls.extend(extract_urls_from_any(item))
        return urls

    if isinstance(value, dict):
        for item in value.values():
            urls.extend(extract_urls_from_any(item))
        return urls

    return urls


def normalise_jotform_field_name(name: str) -> str:
    """Normalise q-numbered JotForm keys for safer comparison."""
    value = str(name or "").strip().lower()
    value = re.sub(r"^q\d+_", "", value)
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def unique_urls(urls: list[str]) -> list[str]:
    """Return URLs in original order without duplicates."""
    seen = set()
    cleaned = []
    for url in urls:
        candidate = (url or "").strip().rstrip(',.;)\"]}')
        if candidate and candidate not in seen:
            seen.add(candidate)
            cleaned.append(candidate)
    return cleaned


def get_jotform_submission_file_urls(submission_id: Optional[str], document_key: str) -> list[str]:
    """Find field-specific uploaded file URLs from the JotForm submission API."""
    if not submission_id:
        return []

    payload = jotform_api_get(f"/submission/{submission_id}")
    content = payload.get("content", payload) if isinstance(payload, dict) else {}
    answers = content.get("answers", {}) if isinstance(content, dict) else {}

    if not isinstance(answers, dict):
        return []

    expected_field = DOCUMENT_UPLOAD_FIELDS.get(document_key, "")
    expected_norm = normalise_jotform_field_name(expected_field)
    matches: list[str] = []

    for answer in answers.values():
        if not isinstance(answer, dict):
            continue

        answer_name = answer.get("name", "")
        answer_text = answer.get("text", "")
        answer_norm = normalise_jotform_field_name(answer_name)
        text_norm = normalise_jotform_field_name(answer_text)

        # Prefer exact field-name matches so the fallback does not attach the
        # wrong applicant document to the wrong candidate field.
        if expected_norm and expected_norm in {answer_norm, text_norm}:
            matches.extend(extract_urls_from_any(answer))

    return unique_urls(matches)


def document_match_tokens(document_key: str) -> set[str]:
    """Build a small token set for cautious matching in /form/{id}/files fallback."""
    label = DOCUMENT_DISPLAY_LABELS.get(document_key, document_key)
    raw = f"{document_key} {DOCUMENT_UPLOAD_FIELDS.get(document_key, '')} {label}".lower()
    tokens = set(re.findall(r"[a-z0-9]+", raw))
    stop_words = {"file", "upload", "proof", "document", "documents", "of", "and", "the", "id"}
    return {token for token in tokens if len(token) >= 3 and token not in stop_words}


def get_jotform_form_file_urls(form_id: Optional[str], document_key: str) -> list[str]:
    """Fallback to /form/{id}/files and cautiously keep URLs that look relevant."""
    if not form_id:
        return []

    payload = jotform_api_get(f"/form/{form_id}/files")
    content = payload.get("content", payload) if isinstance(payload, dict) else payload
    urls = unique_urls(extract_urls_from_any(content))
    if not urls:
        return []

    tokens = document_match_tokens(document_key)
    if not tokens:
        return []

    matched: list[str] = []
    for url in urls:
        lower_url = unquote(urlparse(url).path + " " + urlparse(url).query).lower()
        if any(token in lower_url for token in tokens):
            matched.append(url)

    return unique_urls(matched)


def get_jotform_api_fallback_urls(
    document_key: str,
    *,
    form_id: Optional[str],
    submission_id: Optional[str],
    primary_url: Optional[str] = None,
) -> list[str]:
    """Return possible replacement URLs from JotForm API metadata."""
    urls: list[str] = []

    # Best fallback: field-specific URL from this exact submission.
    urls.extend(get_jotform_submission_file_urls(submission_id, document_key))

    # Secondary fallback: form-level file list with cautious filename matching.
    # This is intentionally conservative to avoid attaching the wrong document.
    urls.extend(get_jotform_form_file_urls(form_id, document_key))

    cleaned = unique_urls(urls)
    if primary_url:
        cleaned = [url for url in cleaned if url != primary_url]
    return cleaned


def download_document_with_jotform_fallback(
    *,
    document_key: str,
    primary_url: str,
    label: str,
    jotform_context: Optional[dict] = None,
) -> dict:
    """Download from webhook URL first, then try JotForm API metadata on HTML.

    Flow:
    1. Try the URL provided in the webhook/rawRequest.
    2. If that URL returns HTML, ask JotForm API for field-specific file URLs.
    3. If API URLs also return HTML/invalid bytes, surface a clear error.
    """
    context = jotform_context or {}

    try:
        downloaded = download_validated_file_from_url(primary_url, label)
        downloaded["download_source"] = "webhook_url"
        downloaded["source_url"] = primary_url
        return downloaded
    except FileDownloadError as primary_error:
        if not primary_error.returned_html:
            raise

        form_id = context.get("form_id")
        submission_id = context.get("submission_id")
        fallback_attempt_errors = [f"Webhook URL: {primary_error}"]

        if not JOTFORM_API_KEY:
            raise FileDownloadError(
                f"{label} webhook URL returned HTML. JotForm API fallback cannot run because "
                "JOTFORM_API_KEY is not configured in the backend .env file.",
                returned_html=True,
            )

        fallback_urls = get_jotform_api_fallback_urls(
            document_key,
            form_id=form_id,
            submission_id=submission_id,
            primary_url=primary_url,
        )

        if not fallback_urls:
            raise FileDownloadError(
                f"{label} webhook URL returned HTML. JotForm API fallback found no matching file URL "
                f"for this field. Check that JOTFORM_API_KEY is configured and the upload field name matches.",
                returned_html=True,
            )

        for index, fallback_url in enumerate(fallback_urls, start=1):
            try:
                downloaded = download_validated_file_from_url(
                    fallback_url,
                    f"{label} JotForm API fallback {index}",
                )
                downloaded["download_source"] = "jotform_api_fallback"
                downloaded["source_url"] = fallback_url
                downloaded["primary_error"] = str(primary_error)
                return downloaded
            except Exception as fallback_error:
                fallback_attempt_errors.append(f"Fallback URL {index}: {fallback_error}")

        raise FileDownloadError(
            f"{label} webhook URL returned HTML. JotForm API fallback was tried, "
            f"but no valid file bytes were downloaded. Attempts: "
            + " | ".join(fallback_attempt_errors),
            returned_html=True,
        )

def is_html_bytes(file_bytes: bytes) -> bool:
    """Return True when downloaded bytes look like an HTML page, not a file."""
    sample = file_bytes[:512].lstrip().lower()
    return sample.startswith(b"<!doctype") or sample.startswith(b"<html") or b"<html" in sample[:128]


def detect_file_type(file_bytes: bytes, header_content_type: str = "") -> dict:
    """Detect the true file type from bytes, falling back to trusted content type.

    File extensions from URLs are intentionally not trusted because JotForm can
    return preview pages, redirects or blocked HTML while the URL still appears
    to end in an image extension.
    """
    header_content_type = (header_content_type or "").split(";")[0].strip().lower()

    if file_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
        return {"extension": ".png", "content_type": "image/png", "detected_type": "PNG image", "previewable": True}
    if file_bytes.startswith(b"\xff\xd8\xff"):
        return {"extension": ".jpg", "content_type": "image/jpeg", "detected_type": "JPEG image", "previewable": True}
    if file_bytes.startswith(b"GIF87a") or file_bytes.startswith(b"GIF89a"):
        return {"extension": ".gif", "content_type": "image/gif", "detected_type": "GIF image", "previewable": True}
    if file_bytes.startswith(b"RIFF") and b"WEBP" in file_bytes[:20]:
        return {"extension": ".webp", "content_type": "image/webp", "detected_type": "WEBP image", "previewable": True}
    if file_bytes.startswith(b"%PDF"):
        return {"extension": ".pdf", "content_type": "application/pdf", "detected_type": "PDF document", "previewable": True}

    stripped = file_bytes[:256].lstrip()
    if stripped.startswith(b"<?xml") or stripped.startswith(b"<svg"):
        return {"extension": ".svg", "content_type": "image/svg+xml", "detected_type": "SVG image", "previewable": True}

    if header_content_type in PREVIEWABLE_CONTENT_TYPES:
        extension = mimetypes.guess_extension(header_content_type) or ".bin"
        if extension == ".jpe":
            extension = ".jpg"
        return {
            "extension": extension,
            "content_type": header_content_type,
            "detected_type": f"{header_content_type} file",
            "previewable": header_content_type.startswith("image/") or header_content_type in {"application/pdf", "text/plain"},
        }

    return {"extension": ".bin", "content_type": header_content_type or "application/octet-stream", "detected_type": "Unknown binary file", "previewable": False}


def download_validated_file_from_url(url: str, document_label: str = "file") -> dict:
    """Download a JotForm file and validate the response before saving it.

    Returns bytes plus detected metadata. Raises ValueError when the URL returns
    HTML, empty content, or a clearly invalid response.
    """
    if not url:
        raise FileDownloadError("No URL supplied")

    if url.startswith("data:"):
        header, encoded = url.split(",", 1)
        encoded += "=" * ((-len(encoded)) % 4)
        file_bytes = base64.b64decode(encoded, validate=False)
        header_content_type = header.split(";")[0].replace("data:", "").strip().lower()

        if not file_bytes:
            raise FileDownloadError(f"{document_label} is empty")
        if is_html_bytes(file_bytes):
            raise FileDownloadError(f"{document_label} decoded to HTML instead of a file", returned_html=True)

        detected = detect_file_type(file_bytes, header_content_type)
        return {"bytes": file_bytes, **detected}

    response = requests.get(
        url,
        timeout=90,
        allow_redirects=True,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "image/avif,image/webp,image/apng,image/*,application/pdf,*/*;q=0.8",
        },
    )

    if response.status_code != 200:
        raise FileDownloadError(f"{document_label} download returned HTTP {response.status_code}")

    file_bytes = response.content
    content_type = response.headers.get("Content-Type", "").split(";")[0].strip().lower()

    if not file_bytes:
        raise FileDownloadError(f"{document_label} download returned an empty file")

    if "text/html" in content_type or is_html_bytes(file_bytes):
        raise FileDownloadError(
            f"{document_label} URL returned HTML instead of file bytes. "
            "This usually means the JotForm link is a preview page, blocked, expired, or requires access.",
            returned_html=True,
        )

    detected = detect_file_type(file_bytes, content_type)
    if detected["extension"] == ".bin" and not content_type.startswith("application/octet-stream"):
        print(f"[STORAGE DEBUG] Unknown file type for {document_label}: content_type={content_type}")

    return {"bytes": file_bytes, **detected}


def get_media_type_for_file(path: str) -> str:
    """Return a useful media type for a saved local file."""
    try:
        with open(path, "rb") as handle:
            file_bytes = handle.read(512)
        return detect_file_type(file_bytes, mimetypes.guess_type(path)[0] or "")["content_type"]
    except Exception:
        return mimetypes.guess_type(path)[0] or "application/octet-stream"

def build_submission_summary(full_name: str, fields: dict, document_urls: dict) -> str:
    """Create the submission_details.txt content using expected saved documents and references."""

    def clean_value(value):
        if value is None:
            return "N/A"

        if isinstance(value, list):
            cleaned = [str(v).strip() for v in value if str(v).strip()]
            return ", ".join(cleaned) if cleaned else "N/A"

        if isinstance(value, dict):
            cleaned = [f"{k}: {v}" for k, v in value.items() if str(v).strip()]
            return ", ".join(cleaned) if cleaned else "N/A"

        value = str(value).strip()
        return value if value else "N/A"

    def get_value(*keys):
        for key in keys:
            value = clean_value(fields.get(key))
            if value != "N/A":
                return value
        return "N/A"

    def get_document_value(*keys):
        for key in keys:
            value = clean_value(document_urls.get(key))
            if value != "N/A":
                return value

        for key in keys:
            value = clean_value(fields.get(key))
            if value != "N/A":
                return value

        return "N/A"

    def get_signature_status():
        signature_value = get_document_value("signature")
        return "Provided" if signature_value != "N/A" else "N/A"

    def format_reference(reference_value):
        if not reference_value:
            return ["N/A"]

        try:
            reference = json.loads(reference_value) if isinstance(reference_value, str) else reference_value
        except Exception:
            return [clean_value(reference_value)]

        if not isinstance(reference, dict) or not reference:
            return ["N/A"]

        return [
            f"Name: {clean_value(reference.get('name'))}",
            f"Phone: {clean_value(reference.get('phone'))}",
            f"Email: {clean_value(reference.get('email'))}",
        ]

    summary_lines = [
        "Candidate Submission Summary",
        "============================",
        "",
        "Candidate Details:",
        f"Candidate name: {get_value('full_name') if get_value('full_name') != 'N/A' else clean_value(full_name)}",
        f"Candidate ID: {get_value('candidate_id')}",
        f"Email: {get_value('email')}",
        f"Mobile: {get_value('mobile_number', 'mobile', 'phone')}",
        f"State: {get_value('state')}",
        f"Car Registration: {get_value('car_registration')}",
        f"Submission Date: {get_value('submission_date')}",
        f"Form ID: {get_value('form_id')}",
        f"Submission ID: {get_value('submission_id')}",
        "",
        "Reference Information:",
        "Reference 1:",
    ]

    for line in format_reference(fields.get("reference_1")):
        summary_lines.append(f"- {line}")

    summary_lines.append("Reference 2:")

    for line in format_reference(fields.get("reference_2")):
        summary_lines.append(f"- {line}")

    summary_lines.extend([
        "",
        "Document Links:",
        f"NDIS worker check: {get_document_value('ndis_worker_check')}",
        f"Police check: {get_document_value('police_check')}",
        f"Working with children: {get_document_value('working_with_children')}",
        f"100 points ID: {get_document_value('id_100_points')}",
        f"First aid / CPR: {get_document_value('first_aid_cpr')}",
        f"NDIS orientation: {get_document_value('ndis_orientation')}",
        f"COVID training: {get_document_value('covid_training')}",
        f"Car insurance: {get_document_value('car_insurance')}",
        f"Car registration proof: {get_document_value('car_rego_proof')}",
        f"Face ID picture: {get_document_value('face_id_picture')}",
        f"Certificates / study: {get_document_value('certificates_study')}",
        f"Signature: {get_signature_status()}",
    ])

    return "\n".join(summary_lines) + "\n"


def candidate_folder_name(cid: str, full_name: Optional[str] = None) -> str:
    """Return the folder name used for both local storage and SharePoint."""
    safe_id = sanitize_filename(cid)
    safe_name = sanitize_filename(full_name or "")

    if safe_name:
        return f"{safe_name}_{safe_id}"

    return safe_id


def candidate_local_folder(cid: str, full_name: Optional[str] = None) -> str:
    """Return the local storage folder for a candidate."""
    return os.path.join(STORAGE_ROOT, candidate_folder_name(cid, full_name))


def save_local_submission_files(
    local_folder: str,
    summary_text: str,
    document_urls: dict,
    jotform_context: Optional[dict] = None,
) -> dict:
    """Save submission details and validated documents to local storage.

    The saved local files become the single reliable source used by the site
    preview and by browser-based SharePoint sync. JotForm URLs are treated only
    as temporary source links used during webhook processing.
    """
    os.makedirs(local_folder, exist_ok=True)
    info_path = os.path.join(local_folder, "submission_details.txt")

    with open(info_path, "w", encoding="utf-8") as handle:
        handle.write(summary_text)

    saved_files = ["submission_details.txt"]
    manifest = [
        {
            "document_key": "submission_details",
            "label": "Submission details",
            "filename": "submission_details.txt",
            "content_type": "text/plain",
            "detected_type": "Text file",
            "previewable": True,
            "size_bytes": os.path.getsize(info_path),
        }
    ]
    errors = []

    for document_key, url in document_urls.items():
        if not url:
            continue

        label = DOCUMENT_DISPLAY_LABELS.get(document_key, document_key.replace("_", " ").title())

        try:
            downloaded = download_document_with_jotform_fallback(
                document_key=document_key,
                primary_url=url,
                label=label,
                jotform_context=jotform_context,
            )
            extension = downloaded["extension"]
            filename = f"{sanitize_filename(document_key)}{extension}"
            file_path = os.path.join(local_folder, filename)

            with open(file_path, "wb") as handle:
                handle.write(downloaded["bytes"])

            file_info = {
                "document_key": document_key,
                "label": label,
                "filename": filename,
                "content_type": downloaded["content_type"],
                "detected_type": downloaded["detected_type"],
                "previewable": downloaded["previewable"],
                "size_bytes": os.path.getsize(file_path),
                "download_source": downloaded.get("download_source", "webhook_url"),
            }
            manifest.append(file_info)
            saved_files.append(filename)
            print(f"[STORAGE DEBUG] Saved {label} as {filename} ({downloaded['detected_type']})")

        except Exception as exc:
            error = {
                "document_key": document_key,
                "label": label,
                "source_url_present": True,
                "fallback_attempted": isinstance(exc, FileDownloadError) and exc.returned_html,
                "error": str(exc),
            }
            errors.append(error)
            print(f"[STORAGE DEBUG] Failed to save local file {label}: {exc}")

    manifest_path = os.path.join(local_folder, "documents_manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as handle:
        json.dump({"files": manifest, "errors": errors}, handle, indent=2)

    if errors:
        error_path = os.path.join(local_folder, "download_errors.txt")
        with open(error_path, "w", encoding="utf-8") as handle:
            handle.write("Some JotForm files could not be downloaded or validated.\n\n")
            for error in errors:
                handle.write(f"- {error['label']}: {error['error']}\n")
        saved_files.append("download_errors.txt")
        manifest.append({
            "document_key": "download_errors",
            "label": "Download errors",
            "filename": "download_errors.txt",
            "content_type": "text/plain",
            "detected_type": "Text file",
            "previewable": True,
            "size_bytes": os.path.getsize(error_path),
        })

        # Rewrite manifest so it includes the error file entry as well.
        with open(manifest_path, "w", encoding="utf-8") as handle:
            json.dump({"files": manifest, "errors": errors}, handle, indent=2)

    return {
        "storage_type": "local",
        "folder": local_folder,
        "files": saved_files,
        "details_file": info_path,
        "manifest_file": manifest_path,
        "errors": errors,
    }


def load_local_documents_manifest(local_folder: str) -> dict:
    """Load local document metadata generated during webhook storage."""
    manifest_path = os.path.join(local_folder, "documents_manifest.json")
    if not os.path.isfile(manifest_path):
        return {"files": [], "errors": []}
    try:
        with open(manifest_path, "r", encoding="utf-8") as handle:
            return json.load(handle)
    except Exception as exc:
        print(f"[STORAGE DEBUG] Could not read documents manifest: {exc}")
        return {"files": [], "errors": []}


def save_candidate_submission(
    cid: str,
    candidate: CandidateIn,
    document_urls: dict,
    jotform_context: Optional[dict] = None,
) -> None:
    """Save the JotForm submission locally for later browser-based SharePoint sync."""
    local_folder = candidate_local_folder(cid, candidate.name)

    summary_text = build_submission_summary(
        candidate.name,
        {
            "candidate_id": cid,
            "full_name": candidate.name,
            "email": candidate.email,
            "mobile_number": candidate.mobile_number,
            "state": candidate.state,
            "car_registration": candidate.car_registration,
            "submission_date": candidate.date,
            "reference_1": candidate.reference_1,
            "reference_2": candidate.reference_2,
            "form_id": (jotform_context or {}).get("form_id"),
            "submission_id": (jotform_context or {}).get("submission_id"),
        },
        document_urls,
    )

    storage_result = save_local_submission_files(local_folder, summary_text, document_urls, jotform_context)
    sync_status = "Pending browser sync"
    if storage_result.get("errors"):
        sync_status = "Pending browser sync - some files failed validation"

    with get_conn() as conn:
        conn.execute(
            """
            UPDATE candidates
            SET local_folder_path = ?,
                sharepoint_sync_status = CASE
                    WHEN sharepoint_folder_url IS NOT NULL AND sharepoint_folder_url != '' THEN sharepoint_sync_status
                    ELSE ?
                END
            WHERE id = ?
            """,
            (local_folder, sync_status, cid),
        )

    print("[STORAGE DEBUG] Submission files saved locally", storage_result)

# ───────────────────────────────────────────────────────────────────────────────
# API ROUTES - WEBHOOK
# ───────────────────────────────────────────────────────────────────────────────

@app.post("/jotform-webhook")
async def jotform_webhook(request: Request):
    """Handle JotForm webhook submissions."""
    form = await request.form()
    data = {key: value for key, value in form.items()}
    form_id = data.get("formID")
    form_title = data.get("formTitle")

    print(f"[WEBHOOK DEBUG] formID={form_id}, formTitle={form_title}, submissionID={data.get('submissionID')}")

    raw = parse_raw_request(data)
    if not raw:
        raise HTTPException(400, "Missing rawRequest payload")

    if form_id == "261460903084858" or (form_title and "document request" in form_title.lower()):
        print("[WEBHOOK DEBUG] routing to document request processor")
        jotform_context = {
            "form_id": form_id or raw.get("formID") or raw.get("form_id"),
            "submission_id": data.get("submissionID") or raw.get("submissionID") or raw.get("submission_id"),
        }
        return handle_document_request(raw, jotform_context)

    print("[WEBHOOK DEBUG] unsupported JotForm webhook payload")
    raise HTTPException(400, "Unsupported JotForm webhook payload")


def handle_document_request(raw: dict, jotform_context: Optional[dict] = None):
    """Create or update a candidate from a JotForm document request submission."""
    candidate, document_urls = build_candidate_from_jotform(raw)

    with get_conn() as conn:
        existing_candidate = find_existing_candidate(conn, candidate.email or "", candidate.name)

        if existing_candidate:
            cid = existing_candidate["id"]
            print(f"[WEBHOOK DEBUG] updating existing candidate {cid} with JotForm data")
            result = update_candidate_record(conn, cid, candidate, preserve_existing_values=True)
        else:
            print(f"[WEBHOOK DEBUG] creating new candidate for {candidate.name}")
            result = insert_candidate_record(conn, candidate)

    cid = result.get("id")
    print(f"[WEBHOOK DEBUG] candidate UUID: {cid}")
    save_candidate_submission(cid, candidate, document_urls, jotform_context)

    with get_conn() as conn:
        refreshed = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()
    return row_to_dict(refreshed)

# ───────────────────────────────────────────────────────────────────────────────
# API ROUTES - CANDIDATES
# ───────────────────────────────────────────────────────────────────────────────

@app.get("/api/candidates", response_model=list[CandidateOut])
def list_candidates(
    stage: Optional[str] = None,
    role: Optional[str] = None,
    q: Optional[str] = None,
):
    """List candidates with optional filtering by stage, role, or search query."""
    sql = "SELECT * FROM candidates WHERE 1=1"
    params = []

    if stage:
        sql += " AND stage = ?"
        params.append(stage)
    if role:
        sql += " AND role = ?"
        params.append(role)
    if q:
        sql += " AND (name LIKE ? OR skills LIKE ?)"
        params.extend([f"%{q}%", f"%{q}%"])

    sql += " ORDER BY created_at DESC"

    with get_conn() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [row_to_dict(row) for row in rows]


@app.post("/api/candidates", response_model=CandidateOut, status_code=201)
def create_candidate(data: CandidateIn):
    """Create a new candidate record."""
    with get_conn() as conn:
        created = insert_candidate_record(conn, data)
    print(f"[WEBHOOK DEBUG] candidate UUID: {created.get('id')}")
    return created


@app.get("/api/candidates/{cid}/local-files")
def list_candidate_local_files(cid: str):
    """List locally saved files for preview and browser SharePoint sync."""
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()

    if not row:
        raise HTTPException(404, "Candidate not found")
  

    candidate = row_to_dict(row)
    folder_name = candidate_folder_name(cid, candidate.get("name"))
    local_folder = candidate.get("local_folder_path") or candidate_local_folder(cid, candidate.get("name"))
    if not os.path.isdir(local_folder):
        return {
            "candidate_id": cid,
            "folder_name": os.path.basename(local_folder.rstrip("/\\")) or folder_name,
            "folder_exists": False,
            "files": [],
            "errors": [],
            "message": "No local files were found for this candidate.",
        }


    manifest = load_local_documents_manifest(local_folder)
    manifest_files = manifest.get("files") or []
    manifest_by_name = {item.get("filename"): item for item in manifest_files if item.get("filename")}

    files = []
    for filename in sorted(os.listdir(local_folder)):
        if filename == "documents_manifest.json":
            continue

        file_path = os.path.join(local_folder, filename)
        if not os.path.isfile(file_path):
            continue

        metadata = manifest_by_name.get(filename, {})
        content_type = metadata.get("content_type") or get_media_type_for_file(file_path)
        files.append({
            "filename": filename,
            "label": metadata.get("label") or filename,
            "document_key": metadata.get("document_key") or os.path.splitext(filename)[0],
            "content_type": content_type,
            "detected_type": metadata.get("detected_type") or detect_file_type(open(file_path, "rb").read(512), content_type)["detected_type"],
            "previewable": bool(metadata.get("previewable", content_type.startswith("image/") or content_type in {"application/pdf", "text/plain"})),
            "size_bytes": os.path.getsize(file_path),
            "download_url": f"/api/candidates/{cid}/files/{quote(filename, safe='')}",
        })

    return {
        "candidate_id": cid,
        "candidate_name": candidate.get("name") or "",
        "folder_name": os.path.basename(local_folder.rstrip("/\\")) or folder_name,
        "folder_exists": True,
        "sharepoint_destination": {
            "configured": bool(get_sharepoint_destination()),
            "display_path": destination_display_path(get_sharepoint_destination()),
        },
        "sharepoint_sync_status": candidate.get("sharepoint_sync_status") or "",
        "sharepoint_folder_url": candidate.get("sharepoint_folder_url") or "",
        "files": files,
        "errors": manifest.get("errors") or [],
    }


@app.get("/api/candidates/{cid}/files/{filename:path}")
def download_candidate_local_file(cid: str, filename: str):
    """Serve one locally saved candidate file to the browser."""
    safe_filename = os.path.basename(filename)
    if safe_filename != filename:
        raise HTTPException(400, "Invalid filename")

    with get_conn() as conn:
        row = conn.execute("SELECT name, local_folder_path FROM candidates WHERE id = ?", (cid,)).fetchone()

    if not row:
        raise HTTPException(404, "Candidate not found")

    local_folder = row["local_folder_path"] or candidate_local_folder(cid, row["name"])
    file_path = os.path.join(local_folder, safe_filename)

    if not os.path.isfile(file_path):
        raise HTTPException(404, "Local file not found")


    return FileResponse(
        file_path,
        filename=safe_filename,
        media_type=get_media_type_for_file(file_path),
        headers={"Content-Disposition": f"inline; filename=\"{safe_filename}\""},
    )


@app.post("/api/candidates/{cid}/sharepoint-sync", response_model=CandidateOut)
def update_candidate_sharepoint_sync(cid: str, data: SharePointSyncIn):
    """Store SharePoint metadata after the browser uploads local files using Graph."""
    with get_conn() as conn:
        existing = conn.execute("SELECT id FROM candidates WHERE id = ?", (cid,)).fetchone()
        if not existing:
            raise HTTPException(404, "Candidate not found")

        conn.execute(
            """
            UPDATE candidates
            SET sharepoint_folder_id = ?,
                sharepoint_folder_url = ?,
                sharepoint_folder_path = ?,
                sharepoint_sync_status = ?
            WHERE id = ?
            """,
            (
                data.folder_id or "",
                data.folder_url or "",
                data.folder_path or "",
                data.status or "Synced",
                cid,
            ),
        )
        row = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()
  
    return row_to_dict(row)


@app.get("/api/candidates/{cid}", response_model=CandidateOut)
def get_candidate(cid: str):
    """Retrieve a specific candidate by ID."""
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()
    if not row:
        raise HTTPException(404, "Candidate not found")
    return row_to_dict(row)


@app.put("/api/candidates/{cid}", response_model=CandidateOut)
def update_candidate(cid: str, data: CandidateIn):
    """Update a candidate, preserving existing values when submitted fields are empty."""
    with get_conn() as conn:
        return update_candidate_record(conn, cid, data, preserve_existing_values=False)


@app.delete("/api/candidates/{cid}", status_code=204)
def delete_candidate(cid: str):
    """Delete a candidate record by ID."""
    with get_conn() as conn:
        existing = conn.execute("SELECT id FROM candidates WHERE id = ?", (cid,)).fetchone()
        if not existing:
            raise HTTPException(404, "Candidate not found")
        conn.execute("DELETE FROM candidates WHERE id = ?", (cid,))

# ───────────────────────────────────────────────────────────────────────────────
# API ROUTES - STATISTICS / HEALTH
# ───────────────────────────────────────────────────────────────────────────────

@app.get("/api/health")
def health_check():
    return {"status": "ok"}


@app.get("/api/stats")
def stats():
    """Get recruitment pipeline statistics."""
    with get_conn() as conn:
        total = conn.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]
        hired = conn.execute("SELECT COUNT(*) FROM candidates WHERE stage = 'Hired'").fetchone()[0]
        interviewed = conn.execute("SELECT COUNT(*) FROM candidates WHERE stage IN ('Interviewed','Hired')").fetchone()[0]
        active = conn.execute("SELECT COUNT(*) FROM candidates WHERE stage NOT IN ('Hired','Declined')").fetchone()[0]
        by_stage = conn.execute("SELECT stage, COUNT(*) as n FROM candidates GROUP BY stage").fetchall()

    return {
        "total": total,
        "hired": hired,
        "interviewed": interviewed,
        "active": active,
        "by_stage": {row["stage"]: row["n"] for row in by_stage},
    }

@app.get("/api/debug/local-file-types/{cid}")
def debug_local_file_types(cid: str):
    """Inspect actual local file bytes and media types for a candidate."""
    with get_conn() as conn:
        row = conn.execute("SELECT name, local_folder_path FROM candidates WHERE id = ?", (cid,)).fetchone()

    if not row:
        raise HTTPException(404, "Candidate not found")

    local_folder = row["local_folder_path"] or candidate_local_folder(cid, row["name"])
    if not os.path.isdir(local_folder):
        raise HTTPException(404, "Local candidate folder not found")

    results = []
    for filename in sorted(os.listdir(local_folder)):
        path = os.path.join(local_folder, filename)
        if not os.path.isfile(path):
            continue

        with open(path, "rb") as handle:
            first_bytes = handle.read(512)

        detected = detect_file_type(first_bytes, mimetypes.guess_type(path)[0] or "")
        results.append({
            "filename": filename,
            "size_bytes": os.path.getsize(path),
            "content_type": detected["content_type"],
            "detected_type": "HTML page saved as file" if is_html_bytes(first_bytes) else detected["detected_type"],
            "previewable": detected["previewable"],
            "first_bytes_hex": first_bytes[:32].hex(),
        })

    return {"candidate_id": cid, "local_folder": local_folder, "files": results}


# ───────────────────────────────────────────────────────────────────────────────
# FRONTEND SERVING
# ───────────────────────────────────────────────────────────────────────────────

init_db()

if os.path.isdir("static"):
    app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/{full_path:path}", include_in_schema=False)
def serve_frontend(full_path: str):
    """Catch-all route to serve the single-page application frontend."""
    index_path = "static/index.html"
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return {"message": "API is running, but static/index.html was not found."}
