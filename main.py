# ═══════════════════════════════════════════════════════════════════════════════
# HELP AT HAND SUPPORT - Recruitment Pipeline API
# ═══════════════════════════════════════════════════════════════════════════════
# A FastAPI application for managing recruitment candidates and document requests
# with JotForm webhook integration for form submissions.

import base64
from urllib import response
import msal
import os
import requests
from dotenv import load_dotenv

load_dotenv()

import re
import ast
import json
import sqlite3
import uuid
from datetime import datetime
from contextlib import contextmanager
from typing import Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel

# ───────────────────────────────────────────────────────────────────────────────
# APPLICATION & CONFIGURATION
# ───────────────────────────────────────────────────────────────────────────────

app = FastAPI(title="HelpAtHandSupport API")
DB_PATH = "hahs.db"
STORAGE_ROOT = os.path.join(os.getcwd(), "uploaded_submissions")
DRIVE_ID = os.getenv("DRIVE_ID", "").strip()

# ───────────────────────────────────────────────────────────────────────────────
# AUTHENTICATION SETUP (MICROSOFT GRAPH API)
# ───────────────────────────────────────────────────────────────────────────────

CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_ID = os.getenv("SITE_ID")
DRIVE_ID = os.getenv("DRIVE_ID")
BASE_FOLDER = os.getenv("BASE_FOLDER", "HR Demo Candidate Log")
STORAGE_ROOT = os.getenv("STORAGE_ROOT", "./local_storage")

msal_app = msal.ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)


def get_graph_access_token() -> str:
    result = msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" in result:
        return result["access_token"]
    raise HTTPException(500, f"Microsoft Graph authentication failed: {result.get('error_description')}")

def get_graph_headers() -> dict:
    return {"Authorization": f"Bearer {get_graph_access_token()}"}

# Test Drive ID directly
response1 = requests.get(
    f"https://graph.microsoft.com/v1.0/drives/{os.getenv('DRIVE_ID')}/root",
    headers={"Authorization": f"Bearer {get_graph_access_token()}"}
)
print("Drive test:", response1.status_code, response1.json())

# Test new Site ID
response2 = requests.get(
    f"https://graph.microsoft.com/v1.0/sites/{os.getenv('SITE_ID')}/drive",
    headers={"Authorization": f"Bearer {get_graph_access_token()}"}
)
print("Site test:", response2.status_code, response2.json())

response3 = requests.get(
    f"https://graph.microsoft.com/v1.0/drive/{os.getenv('DRIVE_ID')}/root",
    headers={"Authorization": f"Bearer {get_graph_access_token()}"}
)
print(response3.status_code)
print(response3.json())

# Test authentication on startup
try:
    get_graph_access_token()
    print("Authentication successful")
except Exception as e:
    print(f"Authentication failed: {e}")

# ───────────────────────────────────────────────────────────────────────────────
# DATABASE MANAGEMENT
# ───────────────────────────────────────────────────────────────────────────────

@contextmanager
def get_conn():
    """
    Context manager for database connections.
    Ensures connections are properly committed and closed.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    """
    Initialize database tables on application startup.
    Creates candidates table with all required fields for storing recruitment data,
    and supporting tables for document requests and references.
    Automatically adds missing columns to existing tables (schema migration).
    """
    with get_conn() as conn:
        # Main candidates table - stores all candidate information
        conn.execute("""
            CREATE TABLE IF NOT EXISTS candidates (
                id                  TEXT PRIMARY KEY,
                name                TEXT NOT NULL,
                role                TEXT NOT NULL,
                skills              TEXT,
                stage               TEXT NOT NULL,
                date                TEXT,
                mobile_number       TEXT,
                email               TEXT,
                state               TEXT,
                car_registration    TEXT,
                ndis_worker_check   TEXT,
                police_check        TEXT,
                working_with_children TEXT,
                id_100_points       TEXT,
                first_aid_cpr       TEXT,
                ndis_orientation    TEXT,
                covid_training      TEXT,
                car_insurance       TEXT,
                car_rego_proof      TEXT,
                face_id_picture     TEXT,
                certificates_study  TEXT,
                signature           TEXT,
                captcha_passed      TEXT,
                confirmation_agreed TEXT,
                reference_1         TEXT,
                reference_2         TEXT,
                created_at          TEXT DEFAULT (datetime('now'))
            )
        """)
        
        # Schema migration - add missing columns if they don't exist
        existing_cols = {row['name'] for row in conn.execute("PRAGMA table_info(candidates)").fetchall()}
        for col, col_type in [
            ('mobile_number', 'TEXT'),
            ('email', 'TEXT'),
            ('state', 'TEXT'),
            ('car_registration', 'TEXT'),
            ('ndis_worker_check', 'TEXT'),
            ('police_check', 'TEXT'),
            ('working_with_children', 'TEXT'),
            ('id_100_points', 'TEXT'),
            ('first_aid_cpr', 'TEXT'),
            ('ndis_orientation', 'TEXT'),
            ('covid_training', 'TEXT'),
            ('car_insurance', 'TEXT'),
            ('car_rego_proof', 'TEXT'),
            ('face_id_picture', 'TEXT'),
            ('certificates_study', 'TEXT'),
            ('signature', 'TEXT'),
            ('captcha_passed', 'TEXT'),
            ('confirmation_agreed', 'TEXT'),
            ('reference_1', 'TEXT'),
            ('reference_2', 'TEXT')
        ]:
            if col not in existing_cols:
                conn.execute(f"ALTER TABLE candidates ADD COLUMN {col} {col_type}")

        # Document request tracking table (for future use)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS document_requests (
                id                      TEXT PRIMARY KEY,
                first_name              TEXT NOT NULL,
                last_name               TEXT NOT NULL,
                mobile_number           TEXT,
                email                   TEXT,
                state                   TEXT,
                car_registration        TEXT,
                ndis_worker_check       TEXT,
                police_check            TEXT,
                working_with_children   TEXT,
                id_100_points           TEXT,
                first_aid_cpr           TEXT,
                ndis_orientation        TEXT,
                covid_training          TEXT,
                car_insurance           TEXT,
                car_rego_proof          TEXT,
                face_id_picture         TEXT,
                certificates_study      TEXT,
                signature               TEXT,
                captcha_passed          TEXT,
                confirmation_agreed     TEXT,
                created_at              TEXT DEFAULT (datetime('now'))
            )
        """)
        
        # Document references tracking table (for future use)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS document_references (
                id                  TEXT PRIMARY KEY,
                document_request_id TEXT NOT NULL,
                first_name          TEXT,
                last_name           TEXT,
                phone_number        TEXT,
                email               TEXT,
                created_at          TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(document_request_id) REFERENCES document_requests(id)
            )
        """)

init_db()



# ───────────────────────────────────────────────────────────────────────────────
# DATA MODELS & VALIDATION
# ───────────────────────────────────────────────────────────────────────────────

# Valid values for candidate roles and pipeline stages
VALID_ROLES = {"Support Worker", "Client"}
VALID_STAGES = {
    "Shortlisted",
    "Screening Call",
    "Documents Requested",
    "Documents Received",
    "Interviewed",
    "Hired",
    "Declined"
}


class CandidateIn(BaseModel):
    """Candidate input model - used for creating and updating candidates."""
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
    reference_1: Optional[str] = None  # JSON string with name, phone, email
    reference_2: Optional[str] = None  # JSON string with name, phone, email


class CandidateOut(CandidateIn):
    """Candidate output model - includes database metadata."""
    id: str
    created_at: str


def validate(data: CandidateIn) -> None:
    """
    Validate candidate input data.
    Raises HTTPException with 400 status if validation fails.
    """
    if not data.name.strip():
        raise HTTPException(400, "Name is required")
    if data.role not in VALID_ROLES:
        raise HTTPException(400, f"Role must be one of: {VALID_ROLES}")
    if data.stage not in VALID_STAGES:
        raise HTTPException(400, f"Stage must be one of: {VALID_STAGES}")


def row_to_dict(row) -> dict:
    """Convert SQLite Row object to dictionary."""
    return dict(row)



# ───────────────────────────────────────────────────────────────────────────────
# API ROUTES
# ───────────────────────────────────────────────────────────────────────────────

# WEBHOOK ROUTES

@app.post("/jotform-webhook")
async def jotform_webhook(request: Request):
    """
    Handle JotForm webhook submissions.
    Routes document requests to the appropriate handler.
    Supports form ID 261460903084858 or any form titled "document request".
    """
    form = await request.form()
    data = {k: v for k, v in form.items()}
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


# CANDIDATE CRUD ROUTES

@app.get("/api/candidates", response_model=list[CandidateOut])
def list_candidates(stage: Optional[str] = None, role: Optional[str] = None, q: Optional[str] = None):
    """
    List candidates with optional filtering by stage, role, or search query.
    Returns candidates ordered by creation date (newest first).
    """
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
        params += [f"%{q}%", f"%{q}%"]
    sql += " ORDER BY created_at DESC"
    with get_conn() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [row_to_dict(r) for r in rows]


@app.post("/api/candidates", response_model=CandidateOut, status_code=201)
def create_candidate(data: CandidateIn):
    """
    Create a new candidate record.
    Validates input data before insertion.
    """
    validate(data)
    cid = str(uuid.uuid4())
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO candidates (id, name, role, skills, stage, date, mobile_number, email, state, car_registration, ndis_worker_check, police_check, working_with_children, id_100_points, first_aid_cpr, ndis_orientation, covid_training, car_insurance, car_rego_proof, face_id_picture, certificates_study, signature, captcha_passed, confirmation_agreed, reference_1, reference_2) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                cid,
                data.name.strip(),
                data.role,
                data.skills or "",
                data.stage,
                data.date,
                data.mobile_number or "",
                data.email or "",
                data.state or "",
                data.car_registration or "",
                data.ndis_worker_check or "",
                data.police_check or "",
                data.working_with_children or "",
                data.id_100_points or "",
                data.first_aid_cpr or "",
                data.ndis_orientation or "",
                data.covid_training or "",
                data.car_insurance or "",
                data.car_rego_proof or "",
                data.face_id_picture or "",
                data.certificates_study or "",
                data.signature or "",
                data.captcha_passed or "",
                data.confirmation_agreed or "",
                data.reference_1 or "",
                data.reference_2 or "",
            )
        )
        row = conn.execute("SELECT * FROM candidates WHERE id= ?", (cid,)).fetchone()
    print(f"[WEBHOOK DEBUG] candidate UUID: {cid}")
    return row_to_dict(row)


@app.get("/api/candidates/{cid}", response_model=CandidateOut)
def get_candidate(cid: str):
    """Retrieve a specific candidate by ID."""
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM candidates WHERE id= ?", (cid,)).fetchone()
    if not row:
        raise HTTPException(404, "Candidate not found")
    return row_to_dict(row)


@app.put("/api/candidates/{cid}", response_model=CandidateOut)
def update_candidate(cid: str, data: CandidateIn):
    """
    Update an existing candidate record with merge strategy.
    Preserves existing values for empty/None fields (manual edits don't overwrite JotForm data).
    Only updates fields that are explicitly provided.
    """
    validate(data)
    with get_conn() as conn:
        existing_row = conn.execute("SELECT * FROM candidates WHERE id= ?", (cid,)).fetchone()
        if not existing_row:
            raise HTTPException(404, "Candidate not found")
        
        # Convert existing row to dict for easier field access
        existing = row_to_dict(existing_row)
        
        # Merge: use new value if provided, otherwise keep existing
        merged = {
            "name": data.name.strip() if data.name else existing.get("name", ""),
            "role": data.role if data.role else existing.get("role", ""),
            "skills": data.skills if data.skills else existing.get("skills", ""),
            "stage": data.stage if data.stage else existing.get("stage", ""),
            "date": data.date if data.date else existing.get("date", ""),
            "mobile_number": data.mobile_number if data.mobile_number else existing.get("mobile_number", ""),
            "email": data.email if data.email else existing.get("email", ""),
            "state": data.state if data.state else existing.get("state", ""),
            "car_registration": data.car_registration if data.car_registration else existing.get("car_registration", ""),
            "ndis_worker_check": data.ndis_worker_check if data.ndis_worker_check else existing.get("ndis_worker_check", ""),
            "police_check": data.police_check if data.police_check else existing.get("police_check", ""),
            "working_with_children": data.working_with_children if data.working_with_children else existing.get("working_with_children", ""),
            "id_100_points": data.id_100_points if data.id_100_points else existing.get("id_100_points", ""),
            "first_aid_cpr": data.first_aid_cpr if data.first_aid_cpr else existing.get("first_aid_cpr", ""),
            "ndis_orientation": data.ndis_orientation if data.ndis_orientation else existing.get("ndis_orientation", ""),
            "covid_training": data.covid_training if data.covid_training else existing.get("covid_training", ""),
            "car_insurance": data.car_insurance if data.car_insurance else existing.get("car_insurance", ""),
            "car_rego_proof": data.car_rego_proof if data.car_rego_proof else existing.get("car_rego_proof", ""),
            "face_id_picture": data.face_id_picture if data.face_id_picture else existing.get("face_id_picture", ""),
            "certificates_study": data.certificates_study if data.certificates_study else existing.get("certificates_study", ""),
            "signature": data.signature if data.signature else existing.get("signature", ""),
            "captcha_passed": data.captcha_passed if data.captcha_passed else existing.get("captcha_passed", ""),
            "confirmation_agreed": data.confirmation_agreed if data.confirmation_agreed else existing.get("confirmation_agreed", ""),
            "reference_1": data.reference_1 if data.reference_1 else existing.get("reference_1", ""),
            "reference_2": data.reference_2 if data.reference_2 else existing.get("reference_2", ""),
        }
        
        conn.execute(
            "UPDATE candidates SET name=?, role=?, skills=?, stage=?, date=?, mobile_number=?, email=?, state=?, car_registration=?, ndis_worker_check=?, police_check=?, working_with_children=?, id_100_points=?, first_aid_cpr=?, ndis_orientation=?, covid_training=?, car_insurance=?, car_rego_proof=?, face_id_picture=?, certificates_study=?, signature=?, captcha_passed=?, confirmation_agreed=?, reference_1=?, reference_2=? WHERE id=?",
            (
                merged["name"],
                merged["role"],
                merged["skills"],
                merged["stage"],
                merged["date"],
                merged["mobile_number"],
                merged["email"],
                merged["state"],
                merged["car_registration"],
                merged["ndis_worker_check"],
                merged["police_check"],
                merged["working_with_children"],
                merged["id_100_points"],
                merged["first_aid_cpr"],
                merged["ndis_orientation"],
                merged["covid_training"],
                merged["car_insurance"],
                merged["car_rego_proof"],
                merged["face_id_picture"],
                merged["certificates_study"],
                merged["signature"],
                merged["captcha_passed"],
                merged["confirmation_agreed"],
                merged["reference_1"],
                merged["reference_2"],
                cid
            )
        )
        row = conn.execute("SELECT * FROM candidates WHERE id= ?", (cid,)).fetchone()
    return row_to_dict(row)


@app.delete("/api/candidates/{cid}", status_code=204)
def delete_candidate(cid: str):
    """
    Delete a candidate record by ID.
    Returns 204 No Content on success, 404 if candidate not found.
    """
    with get_conn() as conn:
        existing = conn.execute("SELECT id FROM candidates WHERE id= ?", (cid,)).fetchone()
        if not existing:
            raise HTTPException(404, "Candidate not found")
        conn.execute("DELETE FROM candidates WHERE id= ?", (cid,))


# STATISTICS ROUTE

@app.get("/api/stats")
def stats():
    """
    Get recruitment pipeline statistics.
    Returns totals by stage and key metrics (hired, interviewed, active).
    """
    with get_conn() as conn:
        total = conn.execute("SELECT COUNT(*) FROM candidates").fetchone()[0]
        hired = conn.execute("SELECT COUNT(*) FROM candidates WHERE stage='Hired'").fetchone()[0]
        interviewed = conn.execute("SELECT COUNT(*) FROM candidates WHERE stage IN ('Interviewed','Hired')").fetchone()[0]
        active = conn.execute("SELECT COUNT(*) FROM candidates WHERE stage NOT IN ('Hired','Declined')").fetchone()[0]
        by_stage = conn.execute("SELECT stage, COUNT(*) as n FROM candidates GROUP BY stage").fetchall()
    return {
        "total": total,
        "hired": hired,
        "interviewed": interviewed,
        "active": active,
        "by_stage": {r["stage"]: r["n"] for r in by_stage}
    }



# ───────────────────────────────────────────────────────────────────────────────
# FRONTEND SERVING
# ───────────────────────────────────────────────────────────────────────────────

app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/{full_path:path}", include_in_schema=False)
def serve_frontend(full_path: str):
    """
    Catch-all route to serve the single-page application frontend.
    Returns index.html for all non-API routes.
    """
    return FileResponse("static/index.html")



# ───────────────────────────────────────────────────────────────────────────────
# WEBHOOK PROCESSING HELPER FUNCTIONS
# ───────────────────────────────────────────────────────────────────────────────

def parse_raw_request(data: dict) -> dict:
    """
    Parse the rawRequest field from JotForm webhook payload.
    Extracts the JSON-encoded form submission data.
    """
    raw = {}
    if data.get("rawRequest"):
        try:
            raw = json.loads(data.get("rawRequest"))
            print(f"[WEBHOOK DEBUG] rawRequest parsed, keys={list(raw.keys())}")
        except (json.JSONDecodeError, TypeError, ValueError) as e:
            print(f"[WEBHOOK DEBUG] rawRequest parse error: {e}")
    return raw


def sanitize_filename(name: str) -> str:
    """Create a filesystem-safe name from a user-provided value."""
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._-")
    return cleaned or "submission"


def download_bytes_from_url(url: str) -> bytes:
    """Download a file from a JotForm URL or a base64 data URL."""
    if not url:
        raise ValueError("No URL supplied")

    if url.startswith("data:"):
        header, encoded = url.split(",", 1)
        mime_type = header.split(";")[0].replace("data:", "")
        padding = (-len(encoded)) % 4
        if padding:
            encoded += "=" * padding
        return base64.b64decode(encoded, validate=False)

    response = requests.get(url, timeout=60, stream=True)
    response.raise_for_status()
    return response.content


def parse_jotform_image_url(raw_value):
    """
    Extract image URLs from various JotForm data structures.
    Handles dicts with url/href/link/src/file keys, lists, strings, and base64 data.
    Returns the URL or None if not found/invalid.
    """
    if not raw_value:
        return None

    if isinstance(raw_value, dict):
        for key in ('url', 'href', 'link', 'src', 'file'):
            if raw_value.get(key):
                return raw_value.get(key)
        return None

    if isinstance(raw_value, list):
        return parse_jotform_image_url(raw_value[0]) if raw_value else None

    if isinstance(raw_value, str):
        text = raw_value.strip()
        if not text:
            return None
        if text.startswith('http') or text.startswith('data:image/'):
            return text
        try:
            parsed = ast.literal_eval(text)
        except (ValueError, SyntaxError):
            parsed = None
        if parsed is not None:
            return parse_jotform_image_url(parsed)
        match = re.search(r'https?://[^\s\'"\[\]]+', text)
        return match.group(0) if match else None

    return None


def merge_jotform_data_into_candidate(cid: str, data: CandidateIn):
    """
    Merge JotForm data into existing candidate, but ONLY update fields that are empty.
    This preserves any manual edits made on the website.
    """
    with get_conn() as conn:
        existing_row = conn.execute("SELECT * FROM candidates WHERE id= ?", (cid,)).fetchone()
        if not existing_row:
            raise HTTPException(404, "Candidate not found")
        
        existing = row_to_dict(existing_row)
        
        # Merge: use JotForm value only if existing field is empty
        merged = {
            "name": existing.get("name", "") or data.name.strip(),
            "role": existing.get("role", "") or data.role,
            "skills": existing.get("skills", "") or data.skills,
            "stage": existing.get("stage", "") or data.stage,
            "date": existing.get("date", "") or data.date,
            "mobile_number": existing.get("mobile_number", "") or data.mobile_number,
            "email": existing.get("email", "") or data.email,
            "state": existing.get("state", "") or data.state,
            "car_registration": existing.get("car_registration", "") or data.car_registration,
            "ndis_worker_check": existing.get("ndis_worker_check", "") or data.ndis_worker_check,
            "police_check": existing.get("police_check", "") or data.police_check,
            "working_with_children": existing.get("working_with_children", "") or data.working_with_children,
            "id_100_points": existing.get("id_100_points", "") or data.id_100_points,
            "first_aid_cpr": existing.get("first_aid_cpr", "") or data.first_aid_cpr,
            "ndis_orientation": existing.get("ndis_orientation", "") or data.ndis_orientation,
            "covid_training": existing.get("covid_training", "") or data.covid_training,
            "car_insurance": existing.get("car_insurance", "") or data.car_insurance,
            "car_rego_proof": existing.get("car_rego_proof", "") or data.car_rego_proof,
            "face_id_picture": existing.get("face_id_picture", "") or data.face_id_picture,
            "certificates_study": existing.get("certificates_study", "") or data.certificates_study,
            "signature": existing.get("signature", "") or data.signature,
            "captcha_passed": existing.get("captcha_passed", "") or data.captcha_passed,
            "confirmation_agreed": existing.get("confirmation_agreed", "") or data.confirmation_agreed,
            "reference_1": existing.get("reference_1", "") or data.reference_1,
            "reference_2": existing.get("reference_2", "") or data.reference_2,
        }
        
        conn.execute(
            "UPDATE candidates SET name=?, role=?, skills=?, stage=?, date=?, mobile_number=?, email=?, state=?, car_registration=?, ndis_worker_check=?, police_check=?, working_with_children=?, id_100_points=?, first_aid_cpr=?, ndis_orientation=?, covid_training=?, car_insurance=?, car_rego_proof=?, face_id_picture=?, certificates_study=?, signature=?, captcha_passed=?, confirmation_agreed=?, reference_1=?, reference_2=? WHERE id=?",
            (
                merged["name"],
                merged["role"],
                merged["skills"],
                merged["stage"],
                merged["date"],
                merged["mobile_number"],
                merged["email"],
                merged["state"],
                merged["car_registration"],
                merged["ndis_worker_check"],
                merged["police_check"],
                merged["working_with_children"],
                merged["id_100_points"],
                merged["first_aid_cpr"],
                merged["ndis_orientation"],
                merged["covid_training"],
                merged["car_insurance"],
                merged["car_rego_proof"],
                merged["face_id_picture"],
                merged["certificates_study"],
                merged["signature"],
                merged["captcha_passed"],
                merged["confirmation_agreed"],
                merged["reference_1"],
                merged["reference_2"],
                cid
            )
        )
        row = conn.execute("SELECT * FROM candidates WHERE id= ?", (cid,)).fetchone()
    return row_to_dict(row)


def build_reference(raw: dict, name_key: str, phone_key: str, email_key: str) -> Optional[str]:
    """
    Convert JotForm reference fields into a compact JSON string.
    Only includes non-empty name, phone, and email values.
    """
    name_obj = raw.get(name_key, {})
    first = name_obj.get("first", "") if isinstance(name_obj, dict) else ""
    last = name_obj.get("last", "") if isinstance(name_obj, dict) else ""
    phone_obj = raw.get(phone_key, {})
    phone = phone_obj.get("full", "") if isinstance(phone_obj, dict) else ""
    email = raw.get(email_key, "")

    reference = {}
    if first or last:
        reference["name"] = (first + " " + last).strip()
    if phone:
        reference["phone"] = phone
    if email:
        reference["email"] = email
    return json.dumps(reference) if reference else None


def handle_document_request(raw: dict):
    """
    Process a document request form submission from JotForm.
    Extracts candidate data and updates existing candidate or creates new one if not found.
    Merges JotForm data with manually-entered data to preserve manual edits.
    """
    # Extract name
    name_obj = raw.get("q1_name", {})
    first_name = name_obj.get("first", "") if isinstance(name_obj, dict) else ""
    last_name = name_obj.get("last", "") if isinstance(name_obj, dict) else ""

    # Extract phone number
    mobile_obj = raw.get("q26_mobileNumber", {})
    mobile_number = mobile_obj.get("full", "") if isinstance(mobile_obj, dict) else ""

    # Extract email
    email = raw.get("q2_email", "")

    # Extract state
    state = raw.get("q46_state", "")

    # Extract document uploads
    ndis_worker_check = parse_jotform_image_url(raw.get("fileUpload", ""))
    police_check = parse_jotform_image_url(raw.get("policeCheck", ""))
    working_with_children = parse_jotform_image_url(raw.get("workingWith", ""))
    id_100_points = parse_jotform_image_url(raw.get("100Points", ""))
    first_aid_cpr = parse_jotform_image_url(raw.get("firstAidcpr", ""))
    ndis_orientation = parse_jotform_image_url(raw.get("ndisWorker", ""))
    covid_training = parse_jotform_image_url(raw.get("covid19Training", ""))
    car_insurance = parse_jotform_image_url(raw.get("evidenceOf", ""))
    car_rego_proof = parse_jotform_image_url(raw.get("evidenceOf16", ""))
    face_id_picture = parse_jotform_image_url(raw.get("pictureOf", ""))
    certificates_study = parse_jotform_image_url(raw.get("certificatesOf", ""))
    signature = parse_jotform_image_url(raw.get("q8_iConfirm", ""))

    # Extract car registration and confirmations
    car_registration = raw.get("q18_carRegistration", "")
    confirmation_agreed = raw.get("q13_iConfirm", "")

    reference_1 = build_reference(raw, "q38_name38", "q39_phoneNumber", "q44_email44")
    reference_2 = build_reference(raw, "q41_name41", "q40_phoneNumber40", "q45_email45")


    if not first_name or not last_name:
        raise HTTPException(400, "Missing required fields: first name and last name")

    # Parse submission date from timestamp
    submission_date = None
    if raw.get("submitDate"):
        try:
            timestamp_ms = int(raw.get("submitDate"))
            dt = datetime.fromtimestamp(timestamp_ms / 1000)
            submission_date = dt.strftime("%Y-%m-%d")
        except (ValueError, TypeError) as e:
            print(f"[WEBHOOK DEBUG] submitDate parse error: {e}")

    """""
    compile_log = {
        "first_name": first_name or "",
        "last_name": last_name or "",
        "mobile_number": mobile_number or "",
        "email": email or "",
        "state": state or "",
        "car_registration": car_registration or "",
        "ndis_worker_check": ndis_worker_check or "",
        "police_check": police_check or "",
        "working_with_children": working_with_children or "",
        "id_100_points": id_100_points or "",
        "first_aid_cpr": first_aid_cpr or "",
        "ndis_orientation": ndis_orientation or "",
        "covid_training": covid_training or "",
        "car_insurance": car_insurance or "",
        "car_rego_proof": car_rego_proof or "",
        "face_id_picture": face_id_picture or "",
        "certificates_study": certificates_study or "",
        "signature": signature or "",
        "car_registration": car_registration or "",
        "reference_1": reference_1 or "",
        "reference_2": reference_2 or "",
    }
 

    print(F"[WEBHOOK DEBUG] compiled document request data: {json.dumps(compile_log, indent=2)}")
   """
    
    # Check if candidate already exists (by email or name)
    full_name = f"{first_name} {last_name}".strip()
    existing_candidate = None
    
    with get_conn() as conn:
        if email:
            existing_candidate = conn.execute("SELECT id FROM candidates WHERE email = ?", (email,)).fetchone()
        if not existing_candidate and full_name:
            existing_candidate = conn.execute("SELECT id FROM candidates WHERE name = ?", (full_name,)).fetchone()
    
    document_urls = {
        "ndis_worker_check": ndis_worker_check,
        "police_check": police_check,
        "working_with_children": working_with_children,
        "id_100_points": id_100_points,
        "first_aid_cpr": first_aid_cpr,
        "ndis_orientation": ndis_orientation,
        "covid_training": covid_training,
        "car_insurance": car_insurance,
        "car_rego_proof": car_rego_proof,
        "face_id_picture": face_id_picture,
        "certificates_study": certificates_study,
        "signature": signature,
    }

    # If candidate exists, update them; otherwise create new
    if existing_candidate:
        cid = existing_candidate['id']
        print(f"[WEBHOOK DEBUG] updating existing candidate {cid} with JotForm data (only filling empty fields)")
        updated = merge_jotform_data_into_candidate(cid, CandidateIn(
            name=full_name,
            role="Support Worker",
            skills="N/A",
            stage="Documents Received",
            date=submission_date,
            mobile_number=mobile_number,
            email=email,
            state=state,
            car_registration=car_registration,
            ndis_worker_check=str(ndis_worker_check) if ndis_worker_check else None,
            police_check=str(police_check) if police_check else None,
            working_with_children=str(working_with_children) if working_with_children else None,
            id_100_points=str(id_100_points) if id_100_points else None,
            first_aid_cpr=str(first_aid_cpr) if first_aid_cpr else None,
            ndis_orientation=str(ndis_orientation) if ndis_orientation else None,
            covid_training=str(covid_training) if covid_training else None,
            car_insurance=str(car_insurance) if car_insurance else None,
            car_rego_proof=str(car_rego_proof) if car_rego_proof else None,
            face_id_picture=str(face_id_picture) if face_id_picture else None,
            certificates_study=str(certificates_study) if certificates_study else None,
            signature=str(signature) if signature else None,
            confirmation_agreed=str(confirmation_agreed) if confirmation_agreed else None,
            reference_1=reference_1 or None,
            reference_2=reference_2 or None,
        ))
        # Log the candidate UUID after update
        try:
            print(f"[WEBHOOK DEBUG] candidate UUID: {updated.get('id')}")
        except Exception:
            pass

        try:
            folder_ref = create_submission_folder(full_name, email)
            save_submission_files(folder_ref, full_name, {
                "full_name": full_name,
                "email": email,
                "mobile_number": mobile_number,
                "state": state,
                "car_registration": car_registration,
                "submission_date": submission_date,
            }, document_urls)
            print("[STORAGE DEBUG] Submission files saved to folder", folder_ref)
        except Exception as exc:
            print(f"[STORAGE DEBUG] Submission storage failed: {exc}")

        return updated
    else:
        print(f"[WEBHOOK DEBUG] creating new candidate for {full_name}")
        created = create_candidate(CandidateIn(
            name=full_name,
            role="Support Worker",
            skills="N/A",
            stage="Documents Received",
            date=submission_date,
            mobile_number=mobile_number,
            email=email,
            state=state,
            car_registration=car_registration,
            ndis_worker_check=str(ndis_worker_check) if ndis_worker_check else None,
            police_check=str(police_check) if police_check else None,
            working_with_children=str(working_with_children) if working_with_children else None,
            id_100_points=str(id_100_points) if id_100_points else None,
            first_aid_cpr=str(first_aid_cpr) if first_aid_cpr else None,
            ndis_orientation=str(ndis_orientation) if ndis_orientation else None,
            covid_training=str(covid_training) if covid_training else None,
            car_insurance=str(car_insurance) if car_insurance else None,
            car_rego_proof=str(car_rego_proof) if car_rego_proof else None,
            face_id_picture=str(face_id_picture) if face_id_picture else None,
            certificates_study=str(certificates_study) if certificates_study else None,
            signature=str(signature) if signature else None,
            confirmation_agreed=str(confirmation_agreed) if confirmation_agreed else None,
            reference_1=reference_1 or None,
            reference_2=reference_2 or None,
        ))

        try:
            folder_ref = create_submission_folder(full_name, email)
            save_submission_files(folder_ref, full_name, {
                "full_name": full_name,
                "email": email,
                "mobile_number": mobile_number,
                "state": state,
                "car_registration": car_registration,
                "submission_date": submission_date,
            }, document_urls)
            print("[STORAGE DEBUG] Submission files saved to folder", folder_ref)
        except Exception as exc:
            print(f"[STORAGE DEBUG] Submission storage failed: {exc}")

        return created

# ───────────────────────────────────────────────────────────────────────────────
# ONE DRIVE INTEGRATION
# ───────────────────────────────────────────────────────────────────────────────

def get_onedrive_drive_target() -> dict:
    """Return the best Graph endpoint information for OneDrive storage."""
    if DRIVE_ID:
        return {"type": "drive", "id": DRIVE_ID, "base_url": f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"}
    if SITE_ID:
        return {"type": "site", "id": SITE_ID, "base_url": f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive"}
    return {}

def create_submission_folder(full_name: str, email: str):
    safe_name = sanitize_filename(full_name or "submission")
    safe_email = sanitize_filename(email or "unknown")
    folder_name = f"{safe_name}_{safe_email}"
    
    local_folder = os.path.join(STORAGE_ROOT, folder_name)
    os.makedirs(local_folder, exist_ok=True)

    if DRIVE_ID:
        try:
            token = get_graph_access_token()
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            response = requests.post(
                f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{BASE_FOLDER}:/children",
                headers=headers,
                json={"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"},
                timeout=120,
            )
            response.raise_for_status()
            payload = response.json()
            print(f"[STORAGE DEBUG] OneDrive folder created: {payload.get('webUrl')}")
            return {
                "storage_type": "onedrive",
                "folder_name": folder_name,
                "local_folder": local_folder,
                "web_url": payload.get("webUrl"),
            }
        except Exception as exc:
            print(f"[STORAGE DEBUG] OneDrive folder creation failed, using local fallback: {exc}")

    return local_folder

def save_submission_files(folder_ref, full_name: str, fields: dict, document_urls: dict) -> dict:
    """Save typed submission details and uploaded images into the folder."""
    summary_lines = [
        f"Submission for: {fields.get('full_name', full_name) or 'Unknown'}",
        f"Email: {fields.get('email', '') or 'N/A'}",
        f"Mobile: {fields.get('mobile_number', '') or 'N/A'}",
        f"State: {fields.get('state', '') or 'N/A'}",
        f"Car Registration: {fields.get('car_registration', '') or 'N/A'}",
        f"Submission Date: {fields.get('submission_date', '') or 'N/A'}",
        "",
        "Document Links:",
    ]
    for key, value in document_urls.items():
        summary_lines.append(f"- {key}: {value or 'Not provided'}")

    summary_text = "\n".join(summary_lines) + "\n"

    # Get folder name from folder_ref
    if isinstance(folder_ref, dict):
        folder_name = folder_ref.get("folder_name")
        local_folder = folder_ref.get("local_folder") or STORAGE_ROOT
        use_onedrive = folder_ref.get("storage_type") == "onedrive"
    else:
        folder_name = None
        local_folder = folder_ref
        use_onedrive = False

    # ── Local storage path ──
    if not use_onedrive:
        os.makedirs(local_folder, exist_ok=True)
        info_path = os.path.join(local_folder, "submission_details.txt")
        with open(info_path, "w", encoding="utf-8") as handle:
            handle.write(summary_text)

        for name, url in document_urls.items():
            if not url:
                continue
            try:
                file_bytes = download_bytes_from_url(url)
                if url.startswith("data:"):
                    ext = ".png"
                    if "image/jpeg" in url:
                        ext = ".jpg"
                else:
                    ext = os.path.splitext(url.split("?")[0])[1] or ".bin"
                file_path = os.path.join(local_folder, f"{sanitize_filename(name)}{ext}")
                with open(file_path, "wb") as handle:
                    handle.write(file_bytes)
            except Exception as exc:
                print(f"[STORAGE DEBUG] Failed to save local image {name}: {exc}")

        return {"storage_type": "local", "folder": local_folder, "details_file": info_path}

    # ── OneDrive upload path ──
    if not DRIVE_ID or not folder_name:
        print("[STORAGE DEBUG] OneDrive upload requested but DRIVE_ID or folder_name missing — saving locally.")
        os.makedirs(local_folder, exist_ok=True)
        info_path = os.path.join(local_folder, "submission_details.txt")
        with open(info_path, "w", encoding="utf-8") as handle:
            handle.write(summary_text)
        return {"storage_type": "local", "folder": local_folder, "details_file": info_path}

    token = get_graph_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    upload_results = {"storage_type": "onedrive", "folder": folder_name, "files": []}
    base_path = f"{BASE_FOLDER}/{folder_name}"

    # Upload submission details text file
    try:
        response = requests.put(
            f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{base_path}/submission_details.txt:/content",
            headers={**headers, "Content-Type": "text/plain"},
            data=summary_text.encode("utf-8"),
            timeout=120,
        )
        response.raise_for_status()
        upload_results["files"].append("submission_details.txt")
        print(f"[STORAGE DEBUG] Uploaded submission_details.txt")
    except Exception as exc:
        print(f"[STORAGE DEBUG] Failed to upload submission_details.txt: {exc}")

    # Upload each document
    for name, url in document_urls.items():
        if not url:
            continue
        try:
            file_bytes = download_bytes_from_url(url)
            if url.startswith("data:"):
                ext = ".png"
                if "image/jpeg" in url:
                    ext = ".jpg"
            else:
                ext = os.path.splitext(url.split("?")[0])[1] or ".bin"

            filename = f"{sanitize_filename(name)}{ext}"
            response = requests.put(
                f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{base_path}/{filename}:/content",
                headers={**headers, "Content-Type": "application/octet-stream"},
                data=file_bytes,
                timeout=120,
            )
            response.raise_for_status()
            upload_results["files"].append(filename)
            print(f"[STORAGE DEBUG] Uploaded {filename}")
        except Exception as exc:
            print(f"[STORAGE DEBUG] OneDrive upload failed for {name}: {exc}")

    return upload_results
