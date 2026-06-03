"""
Run once to populate the database with sample candidates.
Usage: python seed.py
"""
import sqlite3, uuid
 
DB = "hahs.db"
 
candidates = [
    ("Amantha Kulathunga",  "Support Worker", "Experienced in aged care and disability support. Cert III in Individual Support. First aid certified. Strong communication skills with experience in manual handling and personal care.", "Hired",                 "2026-05-27"),
]
 
conn = sqlite3.connect(DB)
conn.execute("""
    CREATE TABLE IF NOT EXISTS candidates (
        id TEXT PRIMARY KEY, name TEXT NOT NULL, role TEXT NOT NULL,
        skills TEXT, stage TEXT NOT NULL, date TEXT,
        created_at TEXT DEFAULT (datetime('now'))
    )
""")
for name, role, skills, stage, date in candidates:
    conn.execute(
        "INSERT OR IGNORE INTO candidates (id, name, role, skills, stage, date) VALUES (?,?,?,?,?,?)",
        (str(uuid.uuid4()), name, role, skills, stage, date)
    )
conn.commit()
conn.close()
print(f"Seeded {len(candidates)} candidates into {DB}")