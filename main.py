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
import os
import re
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Optional
from urllib.parse import quote

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

# SharePoint destination. The browser uses these values when syncing files.
SITE_ID = os.getenv("SITE_ID", "").strip()
DRIVE_ID = os.getenv("DRIVE_ID", "").strip()
BASE_FOLDER = os.getenv("BASE_FOLDER", "HR Demo Candidate Log").strip().strip("/")

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

# ───────────────────────────────────────────────────────────────────────────────
# PURE SPA / MICROSOFT GRAPH CONFIGURATION
# ───────────────────────────────────────────────────────────────────────────────

@app.get("/api/spa-config")
def spa_config():
    """Return public Microsoft/SharePoint config for the browser SPA.

    There is intentionally no client secret here. The frontend uses MSAL.js with
    PKCE and calls Microsoft Graph directly from the browser.
    """
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
            "site_id": SITE_ID,
            "drive_id": DRIVE_ID,
            "base_folder": BASE_FOLDER,
        },
        "configured": bool(MS_CLIENT_ID and MS_TENANT_ID and DRIVE_ID),
    }


@app.get("/api/integration/microsoft/status")
def microsoft_integration_status():
    """Compatibility endpoint used by the Integration tab.

    In pure SPA mode, connection status is checked in the browser because the
    browser owns the Microsoft account session and token cache.
    """
    return {
        "mode": "pure_spa_pkce",
        "configured": bool(MS_CLIENT_ID and MS_TENANT_ID and DRIVE_ID),
        "connected": False,
        "client_id": MS_CLIENT_ID,
        "tenant_id": MS_TENANT_ID,
        "drive_id": DRIVE_ID,
        "base_folder": BASE_FOLDER,
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


def download_bytes_from_url(url: str) -> bytes:
    """Download a file from a JotForm URL or decode a base64 data URL."""
    if not url:
        raise ValueError("No URL supplied")

    if url.startswith("data:"):
        _header, encoded = url.split(",", 1)
        encoded += "=" * ((-len(encoded)) % 4)
        return base64.b64decode(encoded, validate=False)

    response = requests.get(url, timeout=60, stream=True)
    response.raise_for_status()
    return response.content


def get_file_extension(url: str) -> str:
    """Infer a suitable file extension from a URL or data URL."""
    if url.startswith("data:"):
        return ".jpg" if "image/jpeg" in url else ".png"
    return os.path.splitext(url.split("?")[0])[1] or ".bin"


def build_submission_summary(full_name: str, fields: dict, document_urls: dict) -> str:
    """Create the submission_details.txt content."""
    summary_lines = [
        f"Submission for: {fields.get('full_name', full_name) or 'Unknown'}",
        f"Candidate ID: {fields.get('candidate_id') or 'N/A'}",
        f"Email: {fields.get('email') or 'N/A'}",
        f"Mobile: {fields.get('mobile_number') or 'N/A'}",
        f"State: {fields.get('state') or 'N/A'}",
        f"Car Registration: {fields.get('car_registration') or 'N/A'}",
        f"Submission Date: {fields.get('submission_date') or 'N/A'}",
        "",
        "Document Links:",
    ]
    summary_lines.extend(f"- {key}: {value or 'Not provided'}" for key, value in document_urls.items())
    return "\n".join(summary_lines) + "\n"


def candidate_local_folder(cid: str) -> str:
    """Return the local storage folder for a candidate."""
    return os.path.join(STORAGE_ROOT, sanitize_filename(cid))


def save_local_submission_files(local_folder: str, summary_text: str, document_urls: dict) -> dict:
    """Save submission details and documents to local storage.

    This is the main fail-safe in pure SPA mode. The browser can later read
    these local files through the API and upload them to SharePoint using Graph.
    """
    os.makedirs(local_folder, exist_ok=True)
    info_path = os.path.join(local_folder, "submission_details.txt")

    with open(info_path, "w", encoding="utf-8") as handle:
        handle.write(summary_text)

    saved_files = ["submission_details.txt"]
    for name, url in document_urls.items():
        if not url:
            continue
        try:
            file_path = os.path.join(local_folder, f"{sanitize_filename(name)}{get_file_extension(url)}")
            with open(file_path, "wb") as handle:
                handle.write(download_bytes_from_url(url))
            saved_files.append(os.path.basename(file_path))
        except Exception as exc:
            print(f"[STORAGE DEBUG] Failed to save local file {name}: {exc}")

    return {"storage_type": "local", "folder": local_folder, "files": saved_files, "details_file": info_path}


def save_candidate_submission(cid: str, candidate: CandidateIn, document_urls: dict) -> None:
    """Save the JotForm submission locally for later browser-based SharePoint sync."""
    local_folder = candidate_local_folder(cid)
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
        },
        document_urls,
    )

    storage_result = save_local_submission_files(local_folder, summary_text, document_urls)

    with get_conn() as conn:
        conn.execute(
            """
            UPDATE candidates
            SET local_folder_path = ?,
                sharepoint_sync_status = CASE
                    WHEN sharepoint_folder_url IS NOT NULL AND sharepoint_folder_url != '' THEN sharepoint_sync_status
                    ELSE 'Pending browser sync'
                END
            WHERE id = ?
            """,
            (local_folder, cid),
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
        return handle_document_request(raw)

    print("[WEBHOOK DEBUG] unsupported JotForm webhook payload")
    raise HTTPException(400, "Unsupported JotForm webhook payload")


def handle_document_request(raw: dict):
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
    save_candidate_submission(cid, candidate, document_urls)

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
    """List locally saved files for a candidate so the browser can sync them to SharePoint."""
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM candidates WHERE id = ?", (cid,)).fetchone()

    if not row:
        raise HTTPException(404, "Candidate not found")

    candidate = row_to_dict(row)
    local_folder = candidate.get("local_folder_path") or candidate_local_folder(cid)
    if not os.path.isdir(local_folder):
        return {
            "candidate_id": cid,
            "folder_name": sanitize_filename(cid),
            "folder_exists": False,
            "files": [],
            "message": "No local files were found for this candidate.",
        }

    files = []
    for filename in sorted(os.listdir(local_folder)):
        file_path = os.path.join(local_folder, filename)
        if not os.path.isfile(file_path):
            continue
        files.append({
            "filename": filename,
            "size_bytes": os.path.getsize(file_path),
            "download_url": f"/api/candidates/{cid}/files/{quote(filename, safe='')}",
        })

    return {
        "candidate_id": cid,
        "candidate_name": candidate.get("name") or "",
        "folder_name": sanitize_filename(cid),
        "folder_exists": True,
        "drive_id": DRIVE_ID,
        "base_folder": BASE_FOLDER,
        "files": files,
    }


@app.get("/api/candidates/{cid}/files/{filename:path}")
def download_candidate_local_file(cid: str, filename: str):
    """Serve one locally saved candidate file to the browser."""
    safe_filename = os.path.basename(filename)
    if safe_filename != filename:
        raise HTTPException(400, "Invalid filename")

    with get_conn() as conn:
        row = conn.execute("SELECT local_folder_path FROM candidates WHERE id = ?", (cid,)).fetchone()

    if not row:
        raise HTTPException(404, "Candidate not found")

    local_folder = row["local_folder_path"] or candidate_local_folder(cid)
    file_path = os.path.join(local_folder, safe_filename)

    if not os.path.isfile(file_path):
        raise HTTPException(404, "Local file not found")

    return FileResponse(file_path, filename=safe_filename)


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
