# ═══════════════════════════════════════════════════════════════════════════════
# HELP AT HAND SUPPORT - Recruitment Pipeline API
# ═══════════════════════════════════════════════════════════════════════════════
# FastAPI app for managing recruitment candidates, receiving JotForm webhooks,
# and saving uploaded submission documents to a SharePoint document library.

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
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from msal import PublicClientApplication, SerializableTokenCache
from pydantic import BaseModel
from starlette.middleware.sessions import SessionMiddleware

load_dotenv()

# ───────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ───────────────────────────────────────────────────────────────────────────────

DB_PATH = os.getenv("DB_PATH", "hahs.db")
STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./local_storage")
ALLOW_LOCAL_FALLBACK = os.getenv("ALLOW_LOCAL_FALLBACK", "false").lower() == "true"

# SharePoint destination stays fixed in the backend environment.
SITE_ID = os.getenv("SITE_ID", "").strip()
DRIVE_ID = os.getenv("DRIVE_ID", "").strip()
BASE_FOLDER = os.getenv("BASE_FOLDER", "HR Demo Candidate Log").strip().strip("/")

# Client ID and Tenant ID are entered from the Integration tab and stored in DB.
GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All"]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
REDIRECT_URI = os.getenv("REDIRECT_URI")
SESSION_SECRET_KEY = os.getenv("SESSION_SECRET_KEY")

app = FastAPI(title="HelpAtHandSupport API")
app.add_middleware(SessionMiddleware, secret_key=SESSION_SECRET_KEY)

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
        conn.execute("""
            CREATE TABLE IF NOT EXISTS microsoft_integration (
                id TEXT PRIMARY KEY,
                client_id TEXT NOT NULL,
                tenant_id TEXT NOT NULL,
                token_cache TEXT,
                connected_user TEXT,
                connected_at TEXT,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)

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

class MicrosoftIntegrationIn(BaseModel):
    """Integration-tab payload for Microsoft delegated OAuth connection."""
    client_id: str
    tenant_id: str


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
    sharepoint_folder_id: Optional[str] = None
    sharepoint_folder_url: Optional[str] = None
    sharepoint_folder_path: Optional[str] = None


class CandidateOut(CandidateIn):
    """Candidate output model including database metadata."""
    id: str
    created_at: str

# ───────────────────────────────────────────────────────────────────────────────
# MICROSOFT INTEGRATION / GRAPH AUTHENTICATION
# ───────────────────────────────────────────────────────────────────────────────

def get_microsoft_integration() -> Optional[dict]:
    """Return the saved Microsoft integration, if configured."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM microsoft_integration WHERE id = ?",
            (INTEGRATION_ID,),
        ).fetchone()
    return row_to_dict(row) if row else None


def save_microsoft_integration(data: MicrosoftIntegrationIn) -> dict:
    """Save Client ID and Tenant ID from the Integration tab."""
    client_id = data.client_id.strip()
    tenant_id = data.tenant_id.strip()

    if not client_id or not tenant_id:
        raise HTTPException(400, "Client ID and Tenant ID are required")

    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO microsoft_integration (id, client_id, tenant_id)
            VALUES (?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                client_id = excluded.client_id,
                tenant_id = excluded.tenant_id,
                token_cache = NULL,
                connected_user = NULL,
                connected_at = NULL
            """,
            (INTEGRATION_ID, client_id, tenant_id),
        )
        row = conn.execute(
            "SELECT * FROM microsoft_integration WHERE id = ?",
            (INTEGRATION_ID,),
        ).fetchone()
    return row_to_dict(row)


def load_integration_token_cache(integration: dict) -> SerializableTokenCache:
    """Load the MSAL token cache stored in SQLite."""
    cache = SerializableTokenCache()
    if integration.get("token_cache"):
        cache.deserialize(integration["token_cache"])
    return cache


def save_integration_token_cache(cache: SerializableTokenCache) -> None:
    """Persist the MSAL token cache to SQLite."""
    if cache.has_state_changed:
        with get_conn() as conn:
            conn.execute(
                "UPDATE microsoft_integration SET token_cache = ? WHERE id = ?",
                (cache.serialize(), INTEGRATION_ID),
            )


def build_msal_app(
    integration: dict,
    cache: Optional[SerializableTokenCache] = None,
) -> PublicClientApplication:
    """Build MSAL public-client app from saved integration settings."""
    authority = f"https://login.microsoftonline.com/{integration['tenant_id']}"
    return PublicClientApplication(
        client_id=integration["client_id"],
        authority=authority,
        token_cache=cache,
    )


def get_graph_access_token() -> str:
    """Get a delegated Microsoft Graph token from the cached integration."""
    integration = get_microsoft_integration()
    if not integration:
        raise HTTPException(401, "Microsoft integration has not been configured.")

    cache = load_integration_token_cache(integration)
    msal_app = build_msal_app(integration, cache)
    accounts = msal_app.get_accounts()

    if not accounts:
        raise HTTPException(
            401,
            "Microsoft integration is not connected. Connect it from the Integration tab.",
        )

    result = msal_app.acquire_token_silent(scopes=GRAPH_SCOPES, account=accounts[0])
    save_integration_token_cache(cache)

    if not result or "access_token" not in result:
        raise HTTPException(
            401,
            f"Could not get Microsoft Graph token. Reconnect Microsoft integration. Result: {result}",
        )

    return result["access_token"]


def get_graph_headers(content_type: Optional[str] = None) -> dict:
    """Build Microsoft Graph request headers."""
    headers = {
        "Authorization": f"Bearer {get_graph_access_token()}",
        "Accept": "application/json",
    }
    if content_type:
        headers["Content-Type"] = content_type
    return headers


def graph_request(method: str, url: str, **kwargs) -> requests.Response:
    """Small wrapper for Microsoft Graph calls."""
    timeout = kwargs.pop("timeout", 120)
    return requests.request(method, url, timeout=timeout, **kwargs)


@app.post("/api/integration/microsoft")
def configure_microsoft_integration(data: MicrosoftIntegrationIn):
    integration = save_microsoft_integration(data)
    return {
        "message": "Microsoft integration settings saved.",
        "connect_url": "/api/integration/microsoft/connect",
        "client_id": integration["client_id"],
        "tenant_id": integration["tenant_id"],
    }


@app.get("/api/integration/microsoft/status")
def microsoft_integration_status():
    integration = get_microsoft_integration()
    if not integration:
        return {
            "configured": False,
            "connected": False,
            "client_id": "",
            "tenant_id": "",
            "connected_user": "",
            "connected_at": "",
        }

    return {
        "configured": True,
        "connected": bool(integration.get("token_cache")),
        "client_id": integration.get("client_id") or "",
        "tenant_id": integration.get("tenant_id") or "",
        "connected_user": integration.get("connected_user") or "",
        "connected_at": integration.get("connected_at") or "",
    }


@app.get("/api/integration/microsoft/connect")
def connect_microsoft(request: Request):
    integration = get_microsoft_integration()
    if not integration:
        raise HTTPException(400, "Microsoft integration has not been configured yet.")

    cache = load_integration_token_cache(integration)
    msal_app = build_msal_app(integration, cache)
    flow = msal_app.initiate_auth_code_flow(
        scopes=GRAPH_SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    request.session["auth_flow"] = flow
    return RedirectResponse(flow["auth_uri"])


@app.get("/auth/callback")
def auth_callback(request: Request):
    integration = get_microsoft_integration()
    if not integration:
        raise HTTPException(400, "Microsoft integration has not been configured.")

    cache = load_integration_token_cache(integration)
    msal_app = build_msal_app(integration, cache)
    flow = request.session.get("auth_flow")

    if not flow:
        raise HTTPException(400, "Missing auth flow. Start again from the Integration tab.")

    result = msal_app.acquire_token_by_auth_code_flow(flow, dict(request.query_params))
    if "error" in result:
        raise HTTPException(400, detail=result)

    save_integration_token_cache(cache)
    connected_user = result.get("id_token_claims", {}).get("preferred_username", "")

    with get_conn() as conn:
        conn.execute(
            """
            UPDATE microsoft_integration
            SET connected_user = ?, connected_at = datetime('now')
            WHERE id = ?
            """,
            (connected_user, INTEGRATION_ID),
        )

    return {
        "message": "Microsoft SharePoint connection successful.",
        "connected_user": connected_user,
    }


@app.get("/api/debug/token-claims")
def debug_token_claims():
    token = get_graph_access_token()
    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    claims = json.loads(base64.urlsafe_b64decode(payload))

    return {
        "aud": claims.get("aud"),
        "tenant_id": claims.get("tid"),
        "app_id": claims.get("appid"),
        "roles": claims.get("roles"),
        "scp": claims.get("scp"),
        "user": claims.get("preferred_username") or claims.get("upn"),
    }

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


def build_sharepoint_path(*parts: str) -> str:
    """Build a SharePoint document-library path safely for Graph path-based URLs."""
    cleaned_parts = [part.strip("/") for part in parts if part and part.strip("/")]
    return "/".join(cleaned_parts)


def encode_graph_path(path: str) -> str:
    """Encode a SharePoint path while keeping folder separators."""
    return quote(path, safe="/")


def get_sharepoint_folder_by_path(path: str) -> Optional[dict]:
    """Return a SharePoint folder driveItem by path, or None if it does not exist."""
    if not DRIVE_ID:
        return None

    encoded_path = encode_graph_path(path)
    response = graph_request(
        "GET",
        f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/root:/{encoded_path}",
        headers=get_graph_headers(),
    )

    if response.status_code == 404:
        return None

    if response.status_code >= 400:
        print("Graph status:", response.status_code)
        print("Graph response:", response.text)

    response.raise_for_status()
    return response.json()


def create_sharepoint_folder_under_base(folder_name: str) -> dict:
    """Create a candidate folder under BASE_FOLDER in the SharePoint document library."""
    if BASE_FOLDER:
        encoded_base_folder = encode_graph_path(BASE_FOLDER)
        create_folder_url = f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/root:/{encoded_base_folder}:/children"
    else:
        create_folder_url = f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/root/children"

    response = graph_request(
        "POST",
        create_folder_url,
        headers=get_graph_headers("application/json"),
        json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        },
    )
    response.raise_for_status()
    return response.json()


def create_submission_folder(cid: str) -> dict:
    """Get or create the SharePoint folder for a candidate."""
    if not DRIVE_ID:
        message = "DRIVE_ID is missing. Cannot upload to the SharePoint document library."
        if not ALLOW_LOCAL_FALLBACK:
            raise HTTPException(500, message)
        local_folder = os.path.join(STORAGE_ROOT, cid)
        os.makedirs(local_folder, exist_ok=True)
        return {"storage_type": "local", "folder_name": cid, "local_folder": local_folder}

    folder_name = sanitize_filename(cid)
    folder_path = build_sharepoint_path(BASE_FOLDER, folder_name)

    existing_folder = get_sharepoint_folder_by_path(folder_path)
    if existing_folder:
        print(f"[STORAGE DEBUG] SharePoint folder already exists: {existing_folder.get('webUrl')}")
        return {
            "storage_type": "sharepoint",
            "folder_name": folder_name,
            "folder_path": folder_path,
            "folder_id": existing_folder.get("id"),
            "web_url": existing_folder.get("webUrl"),
        }

    try:
        payload = create_sharepoint_folder_under_base(folder_name)
        print(f"[STORAGE DEBUG] SharePoint folder created: {payload.get('webUrl')}")
        return {
            "storage_type": "sharepoint",
            "folder_name": folder_name,
            "folder_path": folder_path,
            "folder_id": payload.get("id"),
            "web_url": payload.get("webUrl"),
        }
    except requests.HTTPError as exc:
        error_body = safe_json(exc.response) if getattr(exc, "response", None) is not None else str(exc)
        message = f"SharePoint folder creation failed: {error_body}"
        print(f"[STORAGE DEBUG] {message}")
        if not ALLOW_LOCAL_FALLBACK:
            raise HTTPException(500, message)

        local_folder = os.path.join(STORAGE_ROOT, folder_name)
        os.makedirs(local_folder, exist_ok=True)
        return {"storage_type": "local", "folder_name": folder_name, "local_folder": local_folder}


def save_local_submission_files(local_folder: str, summary_text: str, document_urls: dict) -> dict:
    """Save submission details and documents to local storage. Used only when fallback is enabled."""
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


def upload_sharepoint_file(
    folder_id: str,
    filename: str,
    content: bytes,
    content_type: str,
) -> dict:
    """Upload one file into a SharePoint folder driveItem using Microsoft Graph."""
    encoded_filename = quote(filename, safe="")
    response = graph_request(
        "PUT",
        f"{GRAPH_BASE_URL}/drives/{DRIVE_ID}/items/{folder_id}:/{encoded_filename}:/content",
        headers=get_graph_headers(content_type),
        data=content,
    )
    response.raise_for_status()
    return response.json()


def save_sharepoint_submission_files(folder_ref: dict, summary_text: str, document_urls: dict) -> dict:
    """Upload submission details and document files to the candidate's SharePoint folder."""
    folder_id = folder_ref.get("folder_id")
    if not folder_id:
        raise HTTPException(500, "SharePoint folder ID is missing; cannot upload files.")

    upload_results = {
        "storage_type": "sharepoint",
        "folder": folder_ref.get("folder_name"),
        "folder_id": folder_id,
        "web_url": folder_ref.get("web_url"),
        "files": [],
    }

    try:
        upload_sharepoint_file(
            folder_id,
            "submission_details.txt",
            summary_text.encode("utf-8"),
            "text/plain",
        )
        upload_results["files"].append("submission_details.txt")
        print("[STORAGE DEBUG] Uploaded submission_details.txt")
    except Exception as exc:
        print(f"[STORAGE DEBUG] Failed to upload submission_details.txt: {exc}")

    for name, url in document_urls.items():
        if not url:
            continue

        filename = f"{sanitize_filename(name)}{get_file_extension(url)}"
        try:
            upload_sharepoint_file(
                folder_id,
                filename,
                download_bytes_from_url(url),
                "application/octet-stream",
            )
            upload_results["files"].append(filename)
            print(f"[STORAGE DEBUG] Uploaded {filename}")
        except Exception as exc:
            print(f"[STORAGE DEBUG] SharePoint upload failed for {name}: {exc}")

    return upload_results


def save_submission_files(
    folder_ref: dict,
    cid: str,
    full_name: str,
    fields: dict,
    document_urls: dict,
) -> dict:
    """Save typed submission details and uploaded documents to SharePoint/local fallback."""
    fields = {**fields, "candidate_id": cid}
    summary_text = build_submission_summary(full_name, fields, document_urls)

    if folder_ref.get("storage_type") == "sharepoint":
        return save_sharepoint_submission_files(folder_ref, summary_text, document_urls)

    if folder_ref.get("storage_type") == "local" and ALLOW_LOCAL_FALLBACK:
        return save_local_submission_files(
            folder_ref.get("local_folder") or STORAGE_ROOT,
            summary_text,
            document_urls,
        )

    raise HTTPException(500, "Submission storage is not configured for SharePoint.")


def update_candidate_sharepoint_info(cid: str, folder_ref: dict) -> None:
    """Store the candidate's SharePoint folder metadata in the database."""
    if folder_ref.get("storage_type") != "sharepoint":
        return

    with get_conn() as conn:
        conn.execute(
            """
            UPDATE candidates
            SET sharepoint_folder_id = ?,
                sharepoint_folder_url = ?,
                sharepoint_folder_path = ?
            WHERE id = ?
            """,
            (
                folder_ref.get("folder_id") or "",
                folder_ref.get("web_url") or "",
                folder_ref.get("folder_path") or "",
                cid,
            ),
        )


def save_candidate_submission(cid: str, candidate: CandidateIn, document_urls: dict) -> None:
    """Create/find the candidate SharePoint folder and save the JotForm files."""
    folder_ref = create_submission_folder(cid)
    save_submission_files(
        folder_ref,
        cid,
        candidate.name,
        {
            "full_name": candidate.name,
            "email": candidate.email,
            "mobile_number": candidate.mobile_number,
            "state": candidate.state,
            "car_registration": candidate.car_registration,
            "submission_date": candidate.date,
        },
        document_urls,
    )
    update_candidate_sharepoint_info(cid, folder_ref)
    print("[STORAGE DEBUG] Submission files saved to folder", folder_ref)

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
