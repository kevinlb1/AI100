from __future__ import annotations

import csv
import io
import json
import os
import random
import re
import sqlite3
import sys
import threading
import time
from datetime import datetime, timezone
from html import escape
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from wsgiref.simple_server import make_server
from generate import (
    generate_preferences_by_category,
    generate_preferences_by_category_mode3,
    generate_preferences_by_category_uniform_real_binned,
    generate_preferences_random,
)


DB_PATH = "match_app.db"


def normalize_base_path(value: str) -> str:
    raw = (value or "").strip()
    if not raw or raw == "/":
        return ""
    if not raw.startswith("/"):
        raw = "/" + raw
    return raw.rstrip("/")


APP_BASE_PATH = normalize_base_path(os.environ.get("APP_BASE_PATH", ""))


def app_url(path: str) -> str:
    raw = path or "/"
    if not raw.startswith("/"):
        raw = "/" + raw
    return f"{APP_BASE_PATH}{raw}" if APP_BASE_PATH else raw

run_state_lock = threading.Lock()
run_state = {
    "thread": None,
    "stop_event": None,
    "current_run_id": None,
    "class_id": None,
}
db_init_lock = threading.Lock()
db_initialized = False
server_state_lock = threading.Lock()
server_instance = None
shutdown_requested = threading.Event()

COMMON_NAMES_100 = [
    "Liam", "Noah", "Oliver", "Elijah", "James", "William", "Benjamin", "Lucas", "Henry", "Theodore",
    "Jack", "Levi", "Alexander", "Jackson", "Mateo", "Daniel", "Michael", "Mason", "Sebastian", "Ethan",
    "Logan", "Owen", "Samuel", "Jacob", "Asher", "Aiden", "John", "Joseph", "Wyatt", "David",
    "Leo", "Luke", "Julian", "Hudson", "Grayson", "Matthew", "Ezra", "Gabriel", "Carter", "Isaac",
    "Jayden", "Luca", "Anthony", "Dylan", "Lincoln", "Thomas", "Maverick", "Elias", "Josiah", "Charles",
    "Olivia", "Emma", "Charlotte", "Amelia", "Sophia", "Isabella", "Ava", "Mia", "Evelyn", "Luna",
    "Harper", "Camila", "Sofia", "Scarlett", "Elizabeth", "Eleanor", "Emily", "Chloe", "Mila", "Violet",
    "Penelope", "Gianna", "Aria", "Abigail", "Ella", "Avery", "Hazel", "Nora", "Layla", "Lily",
    "Aurora", "Nova", "Ellie", "Madison", "Grace", "Isla", "Willow", "Zoey", "Naomi", "Elena",
    "Ivy", "Hannah", "Leah", "Lillian", "Addison", "Aubrey", "Lucy", "Stella", "Natalie", "Maya",
]

AI_CATEGORIES_10 = [
    "Health",
    "Environment",
    "Economics",
    "Education",
    "Agriculture",
    "Urban Planning",
    "Energy",
    "Law and Policy",
    "Humanities",
    "Engineering",
]

AI_TOPICS_BY_CATEGORY: dict[str, list[str]] = {
    "Health": [
        "Breast Cancer Screening", "Drug Development", "Hospital Triage", "Medical Imaging Diagnosis", "Sepsis Prediction",
        "Personalized Nutrition", "Mental Health Monitoring", "Remote Patient Care", "Clinical Trial Matching", "Cardiovascular Risk Forecasting",
        "Diabetes Management", "Stroke Detection", "Surgical Scheduling", "Antibiotic Stewardship", "Genomic Analysis",
        "Elder Care Robotics", "Rehabilitation Planning", "Public Health Surveillance", "Vaccine Distribution", "Emergency Department Flow",
    ],
    "Environment": [
        "Air Quality Forecasting", "Wildfire Risk Mapping", "Flood Prediction", "Drought Monitoring", "Coastal Erosion Analysis",
        "Plastic Waste Sorting", "Biodiversity Tracking", "Water Quality Monitoring", "Deforestation Detection", "Soil Carbon Estimation",
        "Methane Leak Detection", "Recycling Optimization", "Copper Mining Impact Assessment", "Hydroelectricity Planning", "Renewable Siting",
        "Urban Heat Island Mapping", "Invasive Species Detection", "Glacier Change Monitoring", "Marine Pollution Detection", "Sustainable Fisheries Management",
    ],
    "Economics": [
        "Inflation Forecasting", "Labor Market Analysis", "Tax Compliance Modeling", "Supply Chain Resilience", "Credit Risk Assessment",
        "Small Business Lending", "Consumer Demand Forecasting", "Housing Affordability Analysis", "Trade Flow Prediction", "Fraud Detection in Payments",
        "Unemployment Duration Modeling", "Productivity Measurement", "Retail Price Optimization", "Economic Policy Simulation", "Digital Currency Monitoring",
        "Energy Price Forecasting", "Insurance Claim Prediction", "Procurement Cost Forecasting", "Regional Growth Analysis", "Development Program Evaluation",
    ],
    "Education": [
        "Dropout Risk Prediction", "Adaptive Tutoring", "Automated Essay Feedback", "Classroom Engagement Analysis", "Curriculum Mapping",
        "Skill Gap Identification", "Course Recommendation", "Academic Integrity Monitoring", "Early Literacy Assessment", "Special Needs Support",
        "Language Learning Personalization", "Student Wellbeing Signals", "Attendance Intervention", "Teacher Workload Planning", "Enrollment Forecasting",
        "Career Path Guidance", "Peer Collaboration Support", "Learning Resource Discovery", "Laboratory Scheduling", "Alumni Outcome Tracking",
    ],
    "Agriculture": [
        "Crop Yield Forecasting", "Precision Irrigation", "Pest Detection", "Plant Disease Diagnosis", "Soil Nutrient Mapping",
        "Livestock Health Monitoring", "Harvest Timing Prediction", "Weed Identification", "Greenhouse Climate Control", "Fertilizer Optimization",
        "Farm Equipment Maintenance", "Seed Variety Selection", "Post Harvest Loss Reduction", "Precision Spraying Control", "Drought Resistant Breeding",
        "Pollination Monitoring", "Aquaculture Management", "Food Safety Traceability", "Farm Labor Allocation", "Carbon Smart Farming",
    ],
    "Urban Planning": [
        "Traffic Signal Optimization", "Transit Demand Forecasting", "Road Safety Analysis", "Parking Demand Prediction", "Zoning Scenario Modeling",
        "Construction Permit Triage", "Waste Collection Routing", "Noise Pollution Mapping", "Pedestrian Flow Modeling", "Housing Development Prioritization",
        "Public Space Utilization", "Disaster Evacuation Planning", "Utility Outage Prediction", "Streetlight Maintenance", "Sidewalk Accessibility Audits",
        "Emergency Response Coverage", "Crime Hotspot Analysis", "Affordable Housing Allocation", "Bike Lane Planning", "Smart City Sensor Fusion",
    ],
    "Energy": [
        "Grid Load Forecasting", "Demand Response Optimization", "Solar Output Prediction", "Wind Farm Forecasting", "Battery Dispatch Planning",
        "Power Outage Prediction", "Building Energy Efficiency", "EV Charging Optimization", "Microgrid Control", "Carbon Emissions Tracking",
        "Nuclear Plant Maintenance", "Pipeline Leak Detection", "Geothermal Site Screening", "Hydrogen Supply Planning", "Transmission Congestion Analysis",
        "Energy Theft Detection", "Smart Meter Analytics", "Industrial Energy Management", "Renewable Curtailment Reduction", "Heat Pump Adoption Modeling",
    ],
    "Law and Policy": [
        "Case Outcome Forecasting", "Legal Document Review", "Contract Clause Extraction", "Regulatory Compliance Monitoring", "Policy Impact Analysis",
        "Court Scheduling Optimization", "Bail Decision Support", "Public Comment Analysis", "Anti Corruption Risk Scoring", "Benefits Eligibility Screening",
        "Immigration Case Prioritization", "Procurement Integrity Monitoring", "Environmental Regulation Auditing", "Consumer Protection Monitoring", "Open Records Triage",
        "Sentencing Consistency Review", "Legislative Bill Summarization", "Civic Service Routing", "Misinformation Policy Tracking", "Election Operations Planning",
    ],
    "Humanities": [
        "Historical Archive Tagging", "Manuscript Transcription", "Oral History Analysis", "Cultural Heritage Preservation", "Museum Collection Search",
        "Linguistic Change Mapping", "Musicology Pattern Analysis", "Literary Theme Discovery", "Translation Quality Review", "Archaeological Site Detection",
        "Art Style Classification", "Philosophy Argument Mapping", "Religious Text Comparison", "Folklore Motif Indexing", "Digital Humanities Corpus Analysis",
        "Digital Exhibition Curation", "Dialect Documentation", "Theater Script Analysis", "Media Bias Mapping", "Journalism Source Verification",
    ],
    "Engineering": [
        "Predictive Maintenance", "Fault Detection in Sensors", "Robotics Path Planning", "Additive Manufacturing QA", "CAD Design Optimization",
        "Structural Health Monitoring", "Materials Discovery", "Control System Tuning", "Satellite Image Processing", "Manufacturing Line Balancing",
        "Quality Defect Detection", "Semiconductor Yield Prediction", "Autonomous Drone Navigation", "Digital Twin Calibration", "Cyber Physical Security",
        "HVAC Control Optimization", "Construction Progress Tracking", "Water Network Leak Detection", "Bridge Inspection Automation", "Telecom Network Planning",
    ],
}

AI_TOPICS_FLAT = [topic for cat in AI_CATEGORIES_10 for topic in AI_TOPICS_BY_CATEGORY[cat]]


def shuffled_pool_for_class(pool: list[str], class_id: int, salt: int) -> list[str]:
    rng = random.Random(((class_id + 1) * 1_000_003) ^ salt)
    out = list(pool)
    rng.shuffle(out)
    return out


def db_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def get_meta(conn: sqlite3.Connection, key: str, default: str) -> str:
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute("INSERT INTO meta(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))


def table_has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    return column in cols


def get_current_class_id(conn: sqlite3.Connection) -> int:
    raw = get_meta(conn, "current_class_id", "1")
    try:
        class_id = max(1, int(raw))
    except (TypeError, ValueError):
        class_id = 1
    exists = conn.execute("SELECT 1 FROM classes WHERE id=?", (class_id,)).fetchone()
    if exists:
        return class_id
    first = conn.execute("SELECT id FROM classes ORDER BY id LIMIT 1").fetchone()
    if first:
        class_id = int(first[0])
        set_meta(conn, "current_class_id", str(class_id))
        return class_id
    conn.execute("INSERT INTO classes(id, name) VALUES (1, 'Class 1')")
    set_meta(conn, "current_class_id", "1")
    return 1


def set_current_class_id(conn: sqlite3.Connection, class_id: int) -> None:
    set_meta(conn, "current_class_id", str(class_id))


def get_class_meta(conn: sqlite3.Connection, class_id: int, key: str, default: str) -> str:
    row = conn.execute("SELECT value FROM class_meta WHERE class_id=? AND key=?", (class_id, key)).fetchone()
    return row[0] if row else default


def set_class_meta(conn: sqlite3.Connection, class_id: int, key: str, value: str) -> None:
    conn.execute(
        """
        INSERT INTO class_meta(class_id, key, value) VALUES (?, ?, ?)
        ON CONFLICT(class_id, key) DO UPDATE SET value=excluded.value
        """,
        (class_id, key, value),
    )


def create_class(conn: sqlite3.Connection, name: str, n: int) -> int:
    clean_name = (name or "").strip()[:200]
    if not clean_name:
        clean_name = "Untitled class"
    n = max(4, int(n))
    cur = conn.execute("INSERT INTO classes(name) VALUES (?)", (clean_name,))
    class_id = int(cur.lastrowid)
    set_class_meta(conn, class_id, "n", str(n))
    set_class_meta(conn, class_id, "finalized", "0")
    return class_id


def ensure_class_meta_defaults(conn: sqlite3.Connection, class_id: int, n_default: str = "8", finalized_default: str = "0") -> None:
    if conn.execute("SELECT 1 FROM class_meta WHERE class_id=? AND key='n'", (class_id,)).fetchone() is None:
        set_class_meta(conn, class_id, "n", n_default)
    if conn.execute("SELECT 1 FROM class_meta WHERE class_id=? AND key='finalized'", (class_id,)).fetchone() is None:
        set_class_meta(conn, class_id, "finalized", finalized_default)


def initialize_db() -> None:
    conn = db_conn()
    with conn:
        conn.execute("CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT NOT NULL)")
        conn.execute("CREATE TABLE IF NOT EXISTS classes(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL)")
        conn.execute("CREATE TABLE IF NOT EXISTS class_meta(class_id INTEGER NOT NULL, key TEXT NOT NULL, value TEXT NOT NULL, PRIMARY KEY(class_id, key))")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS students(
                class_id INTEGER NOT NULL,
                id INTEGER NOT NULL,
                name TEXT NOT NULL,
                topic_title TEXT NOT NULL,
                PRIMARY KEY(class_id, id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS preferences(
                class_id INTEGER NOT NULL,
                student_id INTEGER NOT NULL,
                topic_id INTEGER NOT NULL,
                score INTEGER NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(class_id, student_id, topic_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS match_runs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                class_id INTEGER NOT NULL DEFAULT 1,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                status TEXT NOT NULL,
                utility REAL,
                penalty INTEGER,
                overlap_count INTEGER,
                finalized_snapshot INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        conn.execute("CREATE TABLE IF NOT EXISTS progress_logs(run_id INTEGER NOT NULL, idx INTEGER NOT NULL, message TEXT NOT NULL, PRIMARY KEY(run_id, idx))")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS selected_topics(
                run_id INTEGER NOT NULL,
                topic_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                partition TEXT NOT NULL,
                PRIMARY KEY(run_id, topic_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS assignments(
                run_id INTEGER NOT NULL,
                student_id INTEGER NOT NULL,
                main_topic INTEGER NOT NULL,
                main_title TEXT NOT NULL,
                main_score INTEGER NOT NULL,
                shadow_topic INTEGER NOT NULL,
                shadow_title TEXT NOT NULL,
                shadow_score INTEGER NOT NULL,
                PRIMARY KEY(run_id, student_id)
            )
            """
        )
        conn.execute("CREATE TABLE IF NOT EXISTS overlaps(run_id INTEGER NOT NULL, s1 INTEGER NOT NULL, s2 INTEGER NOT NULL, PRIMARY KEY(run_id, s1, s2))")

        if not table_has_column(conn, "students", "class_id"):
            conn.execute("ALTER TABLE students RENAME TO students_legacy")
            conn.execute(
                """
                CREATE TABLE students(
                    class_id INTEGER NOT NULL,
                    id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    topic_title TEXT NOT NULL,
                    PRIMARY KEY(class_id, id)
                )
                """
            )
            conn.execute("INSERT INTO students(class_id, id, name, topic_title) SELECT 1, id, name, topic_title FROM students_legacy")
            conn.execute("DROP TABLE students_legacy")

        if not table_has_column(conn, "preferences", "class_id"):
            conn.execute("ALTER TABLE preferences RENAME TO preferences_legacy")
            conn.execute(
                """
                CREATE TABLE preferences(
                    class_id INTEGER NOT NULL,
                    student_id INTEGER NOT NULL,
                    topic_id INTEGER NOT NULL,
                    score INTEGER NOT NULL,
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY(class_id, student_id, topic_id)
                )
                """
            )
            conn.execute(
                """
                INSERT INTO preferences(class_id, student_id, topic_id, score, updated_at)
                SELECT 1, student_id, topic_id, score, updated_at FROM preferences_legacy
                """
            )
            conn.execute("DROP TABLE preferences_legacy")

        if not table_has_column(conn, "match_runs", "class_id"):
            conn.execute("ALTER TABLE match_runs ADD COLUMN class_id INTEGER NOT NULL DEFAULT 1")

        if conn.execute("SELECT COUNT(*) FROM meta").fetchone()[0] == 0:
            set_meta(conn, "n", "8")
            set_meta(conn, "finalized", "0")

        if conn.execute("SELECT COUNT(*) FROM classes").fetchone()[0] == 0:
            class_name = get_meta(conn, "class_name", "Class 1")
            conn.execute("INSERT INTO classes(id, name) VALUES (1, ?)", (class_name,))

        if conn.execute("SELECT COUNT(*) FROM class_meta").fetchone()[0] == 0:
            n_default = get_meta(conn, "n", "8")
            finalized_default = get_meta(conn, "finalized", "0")
            conn.execute("INSERT INTO class_meta(class_id, key, value) VALUES (1, 'n', ?)", (n_default,))
            conn.execute("INSERT INTO class_meta(class_id, key, value) VALUES (1, 'finalized', ?)", (finalized_default,))

        class_rows = conn.execute("SELECT id FROM classes ORDER BY id").fetchall()
        for r in class_rows:
            ensure_class_meta_defaults(conn, int(r[0]))

        current = get_meta(conn, "current_class_id", "")
        if not current.strip():
            set_meta(conn, "current_class_id", "1")

    ensure_students_and_preferences()
    conn.close()


def ensure_db_initialized() -> None:
    global db_initialized
    if db_initialized:
        return
    with db_init_lock:
        if db_initialized:
            return
        initialize_db()
        db_initialized = True


def ensure_students_and_preferences(class_id: int | None = None) -> None:
    conn = db_conn()
    class_id = class_id or get_current_class_id(conn)
    n = int(get_class_meta(conn, class_id, "n", "8"))
    now = datetime.now(timezone.utc).isoformat()
    with conn:
        for i in range(n):
            name = f"Student {i + 1}"
            topic = f"Topic {i + 1}"
            conn.execute(
                """
                INSERT INTO students(class_id, id, name, topic_title) VALUES (?, ?, ?, ?)
                ON CONFLICT(class_id, id) DO NOTHING
                """,
                (class_id, i, name, topic),
            )

        for i in range(n):
            for j in range(n):
                default_score = 4 if i == j else 3
                conn.execute(
                    """
                    INSERT INTO preferences(class_id, student_id, topic_id, score, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(class_id, student_id, topic_id) DO NOTHING
                    """,
                    (class_id, i, j, default_score, now),
                )

    conn.close()


def html_response(start_response, body: str, status: str = "200 OK"):
    start_response(status, [("Content-Type", "text/html; charset=utf-8")])
    return [body.encode("utf-8")]


def json_response(start_response, payload: dict, status: str = "200 OK"):
    start_response(status, [("Content-Type", "application/json; charset=utf-8")])
    return [json.dumps(payload).encode("utf-8")]


def text_response(start_response, body: str, content_type: str = "text/plain; charset=utf-8", status: str = "200 OK", headers: list[tuple[str, str]] | None = None):
    hdrs = [("Content-Type", content_type)]
    if headers:
        hdrs.extend(headers)
    start_response(status, hdrs)
    return [body.encode("utf-8")]


def read_json(environ) -> dict:
    try:
        size = int(environ.get("CONTENT_LENGTH") or 0)
    except ValueError:
        size = 0
    raw = environ["wsgi.input"].read(size).decode("utf-8") if size > 0 else "{}"
    return json.loads(raw or "{}")


def parse_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return False


def compute_home_fingerprint(conn: sqlite3.Connection) -> str:
    class_id = get_current_class_id(conn)
    n = int(get_class_meta(conn, class_id, "n", "8"))
    finalized = get_class_meta(conn, class_id, "finalized", "0")
    latest = conn.execute("SELECT id, status FROM match_runs WHERE class_id=? ORDER BY id DESC LIMIT 1", (class_id,)).fetchone()
    latest_id = str(int(latest["id"])) if latest else ""
    latest_status = str(latest["status"]) if latest else ""
    max_updated = conn.execute(
        "SELECT COALESCE(MAX(updated_at), '') FROM preferences WHERE class_id=? AND student_id < ? AND topic_id < ?",
        (class_id, n, n),
    ).fetchone()[0]
    return f"{class_id}|{n}|{finalized}|{latest_id}|{latest_status}|{max_updated}"


def extract_missing_packages(text: str) -> list[str]:
    if not text:
        return []
    matches = re.findall(r"No module named ['\"]([^'\"]+)['\"]", text)
    seen: set[str] = set()
    ordered: list[str] = []
    for pkg in matches:
        if pkg not in seen:
            seen.add(pkg)
            ordered.append(pkg)
    return ordered


def clear_class_matching_data(conn: sqlite3.Connection, class_id: int) -> None:
    run_ids = [int(r[0]) for r in conn.execute("SELECT id FROM match_runs WHERE class_id=?", (class_id,)).fetchall()]
    for rid in run_ids:
        conn.execute("DELETE FROM progress_logs WHERE run_id=?", (rid,))
        conn.execute("DELETE FROM selected_topics WHERE run_id=?", (rid,))
        conn.execute("DELETE FROM assignments WHERE run_id=?", (rid,))
        conn.execute("DELETE FROM overlaps WHERE run_id=?", (rid,))
    conn.execute("DELETE FROM match_runs WHERE class_id=?", (class_id,))


def sample_unique_labels(pool: list[str], n: int, seed: int, kind: str) -> list[str]:
    if n > len(pool):
        raise ValueError(f"Cannot sample {n} unique {kind}; list has {len(pool)} entries.")
    rng = random.Random(seed)
    return rng.sample(pool, n)


def randomize_class_data(conn: sqlite3.Connection, class_id: int, mode: str, seed: int) -> tuple[bool, str]:
    n = int(get_class_meta(conn, class_id, "n", "8"))
    if n < 4:
        return False, "Matrix size must be at least 4."
    mode = str(mode or "").strip()
    if mode not in {"category_uniform", "category", "category_mode3", "random"}:
        return False, "Invalid random distribution mode."

    category_count = max(1, min(len(AI_CATEGORIES_10), n))
    try:
        names = sample_unique_labels(COMMON_NAMES_100, n, seed ^ 0x4C1B3A77, "student names")
        if mode == "random":
            pref, _r, _base, cat = generate_preferences_random(n=n, C=category_count, seed=seed)
        elif mode == "category":
            pref, _r, _base, cat = generate_preferences_by_category(n=n, C=category_count, seed=seed)
        elif mode == "category_mode3":
            pref, _r, _base, cat = generate_preferences_by_category_mode3(n=n, C=category_count, seed=seed)
        else:
            pref, _r, _base, cat = generate_preferences_by_category_uniform_real_binned(n=n, C=category_count, seed=seed)
        if mode == "random":
            topics = sample_unique_labels(AI_TOPICS_FLAT, n, seed ^ 0x9E3779B1, "topic titles")
        else:
            cat_rng = random.Random(seed ^ 0xA531BEEF)
            per_cat_orders: dict[str, list[str]] = {}
            per_cat_pos: dict[str, int] = {}
            for cname in AI_CATEGORIES_10:
                order = list(AI_TOPICS_BY_CATEGORY[cname])
                cat_rng.shuffle(order)
                per_cat_orders[cname] = order
                per_cat_pos[cname] = 0
            topics = []
            for j in range(n):
                c_idx = int(cat[j]) % len(AI_CATEGORIES_10)
                cname = AI_CATEGORIES_10[c_idx]
                order = per_cat_orders[cname]
                pos = per_cat_pos[cname]
                if pos >= len(order):
                    cat_rng.shuffle(order)
                    per_cat_pos[cname] = 0
                    pos = 0
                topic_name = order[pos]
                per_cat_pos[cname] = pos + 1
                topics.append(f"{cname}: {topic_name}")
    except ValueError as exc:
        return False, str(exc)

    now = datetime.now(timezone.utc).isoformat()
    with conn:
        clear_class_matching_data(conn, class_id)
        conn.execute("DELETE FROM preferences WHERE class_id=?", (class_id,))
        conn.execute("DELETE FROM students WHERE class_id=?", (class_id,))
        for sid in range(n):
            conn.execute(
                "INSERT INTO students(class_id, id, name, topic_title) VALUES (?, ?, ?, ?)",
                (class_id, sid, names[sid], topics[sid]),
            )
        for sid in range(n):
            for tid in range(n):
                conn.execute(
                    "INSERT INTO preferences(class_id, student_id, topic_id, score, updated_at) VALUES (?, ?, ?, ?, ?)",
                    (class_id, sid, tid, int(pref[sid][tid]), now),
                )
    return True, f"Replaced class data using '{mode}' distribution."


def add_students_to_class(conn: sqlite3.Connection, class_id: int, count: int) -> tuple[bool, str]:
    count = int(count)
    if count <= 0:
        return False, "count must be a positive integer."
    current_n = int(get_class_meta(conn, class_id, "n", "8"))
    new_n = current_n + count
    with conn:
        clear_class_matching_data(conn, class_id)
        set_class_meta(conn, class_id, "n", str(new_n))
    ensure_students_and_preferences(class_id)
    return True, f"Added {count} student(s). Class size is now {new_n}."


def remove_students_from_class(conn: sqlite3.Connection, class_id: int, student_ids: list[int]) -> tuple[bool, str]:
    current_n = int(get_class_meta(conn, class_id, "n", "8"))
    unique_ids = sorted({int(x) for x in student_ids if isinstance(x, int) or (isinstance(x, str) and str(x).strip().lstrip('-').isdigit())})
    unique_ids = [x for x in unique_ids if 0 <= x < current_n]
    if not unique_ids:
        return False, "No valid students selected."
    new_n = current_n - len(unique_ids)
    if new_n < 4:
        return False, f"Cannot remove {len(unique_ids)} student(s): class must keep at least 4 students."

    students = get_students(conn, class_id)
    matrix = [get_pref_row(conn, class_id, i, current_n) for i in range(current_n)]
    keep_ids = [i for i in range(current_n) if i not in set(unique_ids)]

    new_students = [students[i] for i in keep_ids]
    new_matrix = [[matrix[old_i][old_j] for old_j in keep_ids] for old_i in keep_ids]
    removed_labels = [f"S{i + 1} ({students[i]['name']})" for i in unique_ids]
    now = datetime.now(timezone.utc).isoformat()

    with conn:
        clear_class_matching_data(conn, class_id)
        conn.execute("DELETE FROM preferences WHERE class_id=?", (class_id,))
        conn.execute("DELETE FROM students WHERE class_id=?", (class_id,))
        set_class_meta(conn, class_id, "n", str(new_n))
        for sid_new, s in enumerate(new_students):
            conn.execute(
                "INSERT INTO students(class_id, id, name, topic_title) VALUES (?, ?, ?, ?)",
                (class_id, sid_new, str(s["name"]), str(s["topic_title"])),
            )
        for i in range(new_n):
            for j in range(new_n):
                conn.execute(
                    "INSERT INTO preferences(class_id, student_id, topic_id, score, updated_at) VALUES (?, ?, ?, ?, ?)",
                    (class_id, i, j, int(new_matrix[i][j]), now),
                )
    return True, f"Removed {len(unique_ids)} student(s): {', '.join(removed_labels)}."


def export_class_csv(conn: sqlite3.Connection, class_id: int) -> str:
    n = int(get_class_meta(conn, class_id, "n", "8"))
    students = get_students(conn, class_id)
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(["student_id", "student_name", "topic_title"] + [f"pref_{i + 1}" for i in range(n)])
    for s in students:
        sid = int(s["id"])
        pref = get_pref_row(conn, class_id, sid, n)
        writer.writerow([sid + 1, s["name"], s["topic_title"], *pref])
    return output.getvalue()


def import_class_csv(conn: sqlite3.Connection, class_id: int, csv_text: str) -> tuple[bool, str]:
    if not csv_text.strip():
        return False, "CSV file is empty."
    try:
        reader = csv.DictReader(io.StringIO(csv_text))
    except Exception as exc:  # noqa: BLE001
        return False, f"Invalid CSV: {exc}"
    if not reader.fieldnames:
        return False, "CSV header is missing."
    fieldnames = [str(f or "").strip() for f in reader.fieldnames]
    pref_cols: list[tuple[int, str]] = []
    for f in fieldnames:
        m = re.fullmatch(r"pref_(\d+)", f)
        if m:
            pref_cols.append((int(m.group(1)), f))
    pref_cols.sort(key=lambda x: x[0])
    if not pref_cols:
        return False, "CSV must contain preference columns named pref_1, pref_2, ..."
    pref_ordered = [name for _idx, name in pref_cols]
    rows = [dict(r) for r in reader]
    if not rows:
        return False, "CSV has no data rows."
    n = len(rows)
    if len(pref_ordered) != n:
        return False, f"CSV must be square: found {n} students but {len(pref_ordered)} preference columns."
    if n < 4:
        return False, "At least 4 students are required."

    parsed_students: list[tuple[str, str]] = []
    parsed_prefs: list[list[int]] = []
    veto_max = n // 4
    for sid, row in enumerate(rows):
        name = str(row.get("student_name", "")).strip()
        title = str(row.get("topic_title", "")).strip()
        if not name:
            return False, f"Row {sid + 2}: student_name is required."
        if not title:
            return False, f"Row {sid + 2}: topic_title is required."
        prefs: list[int] = []
        for col in pref_ordered:
            raw = str(row.get(col, "")).strip()
            try:
                score = int(raw)
            except ValueError:
                return False, f"Row {sid + 2}: {col} must be an integer 0..5."
            if score < 0 or score > 5:
                return False, f"Row {sid + 2}: {col} must be in 0..5."
            prefs.append(score)
        if prefs[sid] < 4:
            return False, f"Row {sid + 2}: own topic preference must be at least 4."
        if sum(1 for x in prefs if x == 0) > veto_max:
            return False, f"Row {sid + 2}: too many vetoes (max {veto_max})."
        parsed_students.append((name[:200], title[:200]))
        parsed_prefs.append(prefs)

    now = datetime.now(timezone.utc).isoformat()
    with conn:
        clear_class_matching_data(conn, class_id)
        conn.execute("DELETE FROM preferences WHERE class_id=?", (class_id,))
        conn.execute("DELETE FROM students WHERE class_id=?", (class_id,))
        set_class_meta(conn, class_id, "n", str(n))
        for sid, (name, title) in enumerate(parsed_students):
            conn.execute(
                "INSERT INTO students(class_id, id, name, topic_title) VALUES (?, ?, ?, ?)",
                (class_id, sid, name, title),
            )
        for sid, prefs in enumerate(parsed_prefs):
            for tid, score in enumerate(prefs):
                conn.execute(
                    "INSERT INTO preferences(class_id, student_id, topic_id, score, updated_at) VALUES (?, ?, ?, ?, ?)",
                    (class_id, sid, tid, score, now),
                )
    return True, f"Imported CSV for {n} students."


def get_students(conn: sqlite3.Connection, class_id: int) -> list[dict]:
    n = int(get_class_meta(conn, class_id, "n", "8"))
    return [
        dict(r)
        for r in conn.execute(
            "SELECT id, name, topic_title FROM students WHERE class_id=? AND id < ? ORDER BY id",
            (class_id, n),
        ).fetchall()
    ]


def get_pref_row(conn: sqlite3.Connection, class_id: int, sid: int, n: int) -> list[int]:
    rows = conn.execute(
        "SELECT topic_id, score FROM preferences WHERE class_id=? AND student_id=? AND topic_id < ? ORDER BY topic_id",
        (class_id, sid, n),
    ).fetchall()
    if len(rows) != n:
        ensure_students_and_preferences(class_id)
        rows = conn.execute(
            "SELECT topic_id, score FROM preferences WHERE class_id=? AND student_id=? AND topic_id < ? ORDER BY topic_id",
            (class_id, sid, n),
        ).fetchall()
    return [int(r[1]) for r in rows]


def student_has_non_default_preferences(pref_row: list[int], sid: int) -> bool:
    for tid, score in enumerate(pref_row):
        default_score = 4 if tid == sid else 3
        if score != default_score:
            return True
    return False


def render_home() -> str:
    conn = db_conn()
    class_id = get_current_class_id(conn)
    class_name_row = conn.execute("SELECT name FROM classes WHERE id=?", (class_id,)).fetchone()
    class_name = str(class_name_row[0]) if class_name_row else f"Class {class_id}"
    students = get_students(conn, class_id)
    n = int(get_class_meta(conn, class_id, "n", "8"))
    finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
    home_fingerprint = compute_home_fingerprint(conn)
    latest = conn.execute("SELECT id, status FROM match_runs WHERE class_id=? ORDER BY id DESC LIMIT 1", (class_id,)).fetchone()
    latest_with_selection = conn.execute(
        """
        SELECT mr.id
        FROM match_runs mr
        WHERE mr.class_id = ?
          AND mr.status != 'running'
          AND EXISTS (SELECT 1 FROM selected_topics st WHERE st.run_id = mr.id)
        ORDER BY mr.id DESC
        LIMIT 1
        """,
        (class_id,),
    ).fetchone()
    latest_group_run_id = int(latest_with_selection[0]) if latest_with_selection else None
    selected_topic_ids_ordered: list[int] = []
    selected_topic_ids: set[int] = set()
    main_topic_by_student: dict[int, int] = {}
    shadow_topic_by_student: dict[int, int] = {}
    overlap_violation_students: set[int] = set()
    if latest_group_run_id is not None:
        selected_topic_ids_ordered = [
            int(r[0])
            for r in conn.execute(
                """
                SELECT topic_id
                FROM selected_topics
                WHERE run_id=?
                ORDER BY CASE partition WHEN 'A' THEN 0 WHEN 'B' THEN 1 ELSE 2 END, topic_id
                """,
                (latest_group_run_id,),
            ).fetchall()
            if 0 <= int(r[0]) < n
        ]
        selected_topic_ids = set(selected_topic_ids_ordered)
        assignment_rows = conn.execute(
            "SELECT student_id, main_topic, shadow_topic FROM assignments WHERE run_id=? ORDER BY student_id",
            (latest_group_run_id,),
        ).fetchall()
        main_topic_by_student = {
            int(r[0]): int(r[1])
            for r in assignment_rows
            if 0 <= int(r[0]) < n and 0 <= int(r[1]) < n
        }
        shadow_topic_by_student = {
            int(r[0]): int(r[2])
            for r in assignment_rows
            if 0 <= int(r[0]) < n and 0 <= int(r[2]) < n
        }
        overlap_rows = conn.execute("SELECT s1, s2 FROM overlaps WHERE run_id=? ORDER BY s1, s2", (latest_group_run_id,)).fetchall()
        for r in overlap_rows:
            s1 = int(r[0])
            s2 = int(r[1])
            if 0 <= s1 < n:
                overlap_violation_students.add(s1)
            if 0 <= s2 < n:
                overlap_violation_students.add(s2)
    pref_rows = [(s["id"], get_pref_row(conn, class_id, s["id"], n)) for s in students]
    conn.close()

    if selected_topic_ids_ordered:
        ordered_topic_ids = selected_topic_ids_ordered + [tid for tid in range(n) if tid not in selected_topic_ids]
    else:
        ordered_topic_ids = list(range(n))
    group_rank = {tid: idx for idx, tid in enumerate(selected_topic_ids_ordered)}
    has_selected_columns = len(selected_topic_ids) > 0
    dense_mode = n >= 24
    names_by_id = {s["id"]: s["name"] for s in students}
    topic_titles_by_id = {s["id"]: s["topic_title"] for s in students}
    category_index = {name: idx + 1 for idx, name in enumerate(AI_CATEGORIES_10)}
    short_topic_code_by_id: dict[int, str] = {}
    fallback_topic_num_within_category: dict[str, int] = {}
    for tid in range(n):
        title = str(topic_titles_by_id.get(tid, ""))
        if ":" not in title:
            continue
        category_name, topic_name = title.split(":", 1)
        category_name = category_name.strip()
        topic_name = topic_name.strip()
        cnum = category_index.get(category_name)
        if cnum is None:
            continue
        topic_list = AI_TOPICS_BY_CATEGORY.get(category_name, [])
        tnum = 0
        if topic_list:
            try:
                tnum = topic_list.index(topic_name) + 1
            except ValueError:
                tnum = 0
        if tnum <= 0:
            next_num = fallback_topic_num_within_category.get(category_name, 0) + 1
            fallback_topic_num_within_category[category_name] = next_num
            tnum = next_num
        short_topic_code_by_id[tid] = f"C{cnum}:T{tnum}"
    group_palette = [
        "#22c55e", "#06b6d4", "#f59e0b", "#a78bfa", "#ef4444", "#10b981",
        "#3b82f6", "#eab308", "#ec4899", "#14b8a6", "#84cc16", "#f97316",
    ]

    def color_for_group(gid: int) -> str:
        return group_palette[gid % len(group_palette)]

    def hex_to_rgba(hex_color: str, alpha: float) -> str:
        h = hex_color.lstrip("#")
        if len(h) != 6:
            return f"rgba(59,130,246,{alpha})"
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        return f"rgba({r},{g},{b},{alpha})"
    if latest:
        latest_status = str(latest["status"])
        if latest_status.startswith("error:"):
            latest_status = "error"
        latest_html = f"<p>Latest run: #{latest['id']} ({latest_status})</p>"
    else:
        latest_html = "<p>No runs yet.</p>"

    topic_headers = []
    for tid in ordered_topic_ids:
        topic_title = topic_titles_by_id.get(tid, "")
        col_class = "topic-head"
        group_tag_html = ""
        header_style = ""
        if has_selected_columns and tid in group_rank:
            gid = group_rank[tid]
            color = color_for_group(gid)
            col_class += " group-head"
            group_tag_html = f"<div class='topic-group' style='color:{color};'>G{gid + 1}</div>"
            header_style = f" style='background:{hex_to_rgba(color, 0.12)};'"
        elif has_selected_columns:
            col_class += " unmatched-topic-col"
        short_code = escape(short_topic_code_by_id.get(tid, f"T{tid + 1}"))
        topic_label_html = (
            f"<div class='topic-code topic-code-short'>{short_code}</div>"
            f"<div class='topic-code topic-code-full'>{escape(topic_title)}</div>"
        )
        topic_headers.append(
            f"<th class='{col_class}' title='{escape(topic_title)}'{header_style}>"
            f"{topic_label_html}"
            f"<div class='topic-title'>{escape(topic_title)}</div>"
            f"{group_tag_html}"
            f"</th>"
        )
    topic_col_defs = []
    for tid in ordered_topic_ids:
        col_cls = "topic-col"
        if has_selected_columns and tid not in selected_topic_ids:
            col_cls += " unmatched-topic-col"
        topic_col_defs.append(f"<col class='{col_cls}'>")
    colgroup = "<col class='student-col'>" + "".join(topic_col_defs)

    def row_sort_key(row: tuple[int, list[int]]) -> tuple[int, int, int]:
        sid = int(row[0])
        main_topic = main_topic_by_student.get(sid)
        if main_topic is None or main_topic not in group_rank:
            return (1, 10_000, sid)
        shadow_topic = shadow_topic_by_student.get(sid)
        shadow_rank = group_rank.get(shadow_topic, 10_000)
        return (0, group_rank[main_topic], shadow_rank, sid)

    def build_topic_chip(topic_id: int | None, is_shadow: bool) -> str:
        if topic_id is None or topic_id < 0 or topic_id >= n:
            return ""
        if topic_id in group_rank:
            gid = group_rank[topic_id]
            color = color_for_group(gid)
            label = f"G{gid + 1}"
        else:
            color = "#64748b"
            label = f"T{topic_id + 1}"
        title = topic_titles_by_id.get(topic_id, f"Topic {topic_id + 1}")
        chip_class = "group-chip shadow-chip" if is_shadow else "group-chip main-chip"
        return (
            f"<span class='{chip_class}' "
            f"title='{'Shadow' if is_shadow else 'Main'} topic: {escape(title)}' "
            f"style='color:{color}; border-color:{color}; background:{hex_to_rgba(color, 0.16)};'>{label}</span>"
        )

    if has_selected_columns:
        pref_rows = sorted(pref_rows, key=row_sort_key)

    row_group_ids = []
    for sid, _pref in pref_rows:
        if has_selected_columns:
            row_group_ids.append(group_rank.get(main_topic_by_student.get(sid, -1), -1))
        else:
            row_group_ids.append(-1)

    group_bounds: dict[int, tuple[int, int]] = {}
    if has_selected_columns:
        for idx, gid in enumerate(row_group_ids):
            if gid < 0:
                continue
            if gid not in group_bounds:
                group_bounds[gid] = (idx, idx)
            else:
                start, _end = group_bounds[gid]
                group_bounds[gid] = (start, idx)

    matrix_rows = []
    for row_idx, (sid, pref) in enumerate(pref_rows):
        changed = student_has_non_default_preferences(pref, sid)
        cells = []
        row_gid = row_group_ids[row_idx]
        shadow_topic = shadow_topic_by_student.get(sid)
        shadow_gid = group_rank.get(shadow_topic, -1) if shadow_topic is not None else -1
        shadow_color = color_for_group(shadow_gid) if shadow_gid >= 0 else "#64748b"
        for tid in ordered_topic_ids:
            score = pref[tid]
            cell_class = f"score-cell score-{score}"
            if has_selected_columns and tid not in selected_topic_ids:
                cell_class += " unmatched-topic-cell"
            cell_style = ""
            score_html = str(score)
            if has_selected_columns:
                col_gid = group_rank.get(tid, -1)
                if row_gid >= 0 and col_gid == row_gid:
                    gcolor = color_for_group(row_gid)
                    start, end = group_bounds.get(row_gid, (row_idx, row_idx))
                    if start == end:
                        block_pos = "group-box-single"
                    elif row_idx == start:
                        block_pos = "group-box-top"
                    elif row_idx == end:
                        block_pos = "group-box-bottom"
                    else:
                        block_pos = "group-box-mid"
                    cell_class += f" active-pair-cell {block_pos}"
                    cell_style = f" style='--gcolor:{gcolor};'"
            if shadow_topic is not None and tid == shadow_topic:
                score_html = f"<span class='shadow-pref-circle' style='--shadow-gcolor:{shadow_color};'>{score}</span>"
            cells.append(f"<td class='{cell_class}'{cell_style}>{score_html}</td>")
        row_class = "matrix-row" if changed else "matrix-row matrix-row-dim"
        student_name = escape(names_by_id.get(sid, ""))
        main_topic = main_topic_by_student.get(sid)
        has_overlap_violation = sid in overlap_violation_students
        main_chip = build_topic_chip(main_topic, is_shadow=False)
        shadow_chip = build_topic_chip(shadow_topic, is_shadow=True)
        violation_html = "<span class='shadow-violation-star' title='Main/shadow overlap violation'>*</span>" if has_overlap_violation else ""
        matrix_rows.append(
            f"<tr class='{row_class}' onclick='window.location.href=\"{app_url(f'/student?sid={sid}')}\"'>"
            f"<td class='student-cell' title='{student_name}'>"
            f"<span class='student-main-group'>{main_chip}</span>"
            f"<span class='student-name'>{student_name}</span>"
            f"<span class='student-group-right'>{violation_html}{shadow_chip}</span>"
            f"</td>{''.join(cells)}</tr>"
        )
    return f"""
    <html><head><title>Topic Match</title><style>{base_css()}</style></head><body>
    <main class='container home-layout {'matrix-dense' if dense_mode else ''}'>
      <h1>Topic Matching Portal</h1>
      <p class='muted'>Class: <strong>{escape(class_name)}</strong></p>
      <p>Student editing of preferences: <strong id='lockStatus'>{'Locked' if finalized else 'Unlocked'}</strong></p>
      <div class='top-right-actions'>
        <a class='button-link' href='{app_url("/admin")}'>Open Admin Dashboard</a>
        <a class='button-link button-danger' href='{app_url("/shutdown")}'>Stop server</a>
      </div>
      <div class='matrix-controls'>
        <button id='topicTitleToggleBtn' type='button'>Expand topic titles</button>
        <button id='matrixFlowToggleBtn' type='button'>Fit/Scroll: Fit</button>
        <button id='unmatchedToggleBtn' type='button'>Hide unmatched topics</button>
      </div>
      <div id='matrixWrap' class='matrix-wrap matrix-fit table-mode-compressed'>
        <table class='matrix-table'>
          <colgroup>{colgroup}</colgroup>
          <thead>
            <tr><th class='student-head'>Student</th>{''.join(topic_headers)}</tr>
          </thead>
          <tbody>
            {"".join(matrix_rows)}
          </tbody>
        </table>
      </div>
      {latest_html}
    </main>
    <script>
      const BASE_PATH = {json.dumps(APP_BASE_PATH)};
      function appUrl(path) {{
        const raw = String(path || '/');
        const normalized = raw.startsWith('/') ? raw : '/' + raw;
        return BASE_PATH ? BASE_PATH + normalized : normalized;
      }}
      const THEME_MODE_KEY = 'adminThemeMode';
      function applyThemeMode(mode) {{
        const root = document.documentElement;
        const next = (mode === 'light' || mode === 'dark') ? mode : 'system';
        if (next === 'system') root.removeAttribute('data-theme');
        else root.setAttribute('data-theme', next);
      }}
      function initThemeMode() {{
        let mode = 'system';
        try {{
          const stored = localStorage.getItem(THEME_MODE_KEY);
          if (stored) mode = stored;
        }} catch (_err) {{}}
        applyThemeMode(mode);
      }}
      const homeFingerprint = {json.dumps(home_fingerprint)};
      const hasMatchedColumns = {str(has_selected_columns).lower()};
      const uiStateStorageKey = 'topic_match_home_ui_state_v1';
      let topicTitleExpanded = false;
      let matrixFlowMode = 'fit'; // fit | scroll
      let hideUnmatchedTopics = false;
      function loadUiState() {{
        try {{
          const raw = window.localStorage.getItem(uiStateStorageKey);
          if (!raw) return null;
          const parsed = JSON.parse(raw);
          if (!parsed || typeof parsed !== 'object') return null;
          return parsed;
        }} catch (_err) {{
          return null;
        }}
      }}
      function saveUiState() {{
        try {{
          const state = {{
            topicTitleExpanded: !!topicTitleExpanded,
            matrixFlowMode: matrixFlowMode === 'scroll' ? 'scroll' : 'fit',
            hideUnmatchedTopics: !!hideUnmatchedTopics,
          }};
          window.localStorage.setItem(uiStateStorageKey, JSON.stringify(state));
        }} catch (_err) {{
          // ignore storage errors
        }}
      }}
      function setTopicTitleExpanded(expanded) {{
        const wrap = document.getElementById('matrixWrap');
        const btn = document.getElementById('topicTitleToggleBtn');
        if (!wrap || !btn) return;
        topicTitleExpanded = !!expanded;
        wrap.classList.remove('table-mode-compressed', 'table-mode-all');
        wrap.classList.add(topicTitleExpanded ? 'table-mode-all' : 'table-mode-compressed');
        if (topicTitleExpanded) {{
          btn.innerText = 'Compress topic titles';
        }} else {{
          btn.innerText = 'Expand topic titles';
        }}
        saveUiState();
      }}
      function toggleTopicTitleMode() {{
        setTopicTitleExpanded(!topicTitleExpanded);
      }}
      function setMatrixFlowMode(mode) {{
        const wrap = document.getElementById('matrixWrap');
        const btn = document.getElementById('matrixFlowToggleBtn');
        if (!wrap || !btn) return;
        if (mode !== 'fit' && mode !== 'scroll') mode = 'fit';
        matrixFlowMode = mode;
        const effectiveMode = hideUnmatchedTopics ? 'fit' : matrixFlowMode;
        wrap.classList.remove('matrix-fit', 'matrix-scroll');
        wrap.classList.add('matrix-' + effectiveMode);
        btn.innerText = 'Fit/Scroll: ' + (effectiveMode === 'fit' ? 'Fit' : 'Scroll');
        btn.title = (hideUnmatchedTopics && matrixFlowMode === 'scroll') ? 'Unmatched topics hidden: forced to fit mode.' : '';
        saveUiState();
      }}
      function toggleMatrixFlowMode() {{
        setMatrixFlowMode(matrixFlowMode === 'fit' ? 'scroll' : 'fit');
      }}
      function setHideUnmatched(hidden) {{
        const wrap = document.getElementById('matrixWrap');
        const btn = document.getElementById('unmatchedToggleBtn');
        hideUnmatchedTopics = !!hidden;
        if (wrap) {{
          wrap.classList.toggle('hide-unmatched', hideUnmatchedTopics);
        }}
        if (btn) {{
          btn.innerText = hideUnmatchedTopics ? 'Show unmatched topics' : 'Hide unmatched topics';
        }}
        setMatrixFlowMode(matrixFlowMode);
        saveUiState();
      }}
      function toggleHideUnmatched() {{
        if (!hasMatchedColumns) return;
        setHideUnmatched(!hideUnmatchedTopics);
      }}
      async function refreshLockStatus() {{
        try {{
          const r = await fetch(appUrl('/api/finalized'));
          if (!r.ok) return;
          const data = await r.json();
          document.getElementById('lockStatus').innerText = data.finalized ? 'Locked' : 'Unlocked';
        }} catch (err) {{
          // keep current text on transient errors
        }}
      }}
      async function refreshHomeIfChanged() {{
        try {{
          const r = await fetch(appUrl('/api/home_fingerprint'));
          if (!r.ok) return;
          const data = await r.json();
          if ((data.fingerprint || '') !== homeFingerprint) {{
            window.location.reload();
          }}
        }} catch (err) {{
          // ignore transient poll errors
        }}
      }}
      const topicTitleToggleBtn = document.getElementById('topicTitleToggleBtn');
      if (topicTitleToggleBtn) {{
        topicTitleToggleBtn.addEventListener('click', toggleTopicTitleMode);
      }}
      const matrixFlowToggleBtn = document.getElementById('matrixFlowToggleBtn');
      if (matrixFlowToggleBtn) {{
        matrixFlowToggleBtn.addEventListener('click', toggleMatrixFlowMode);
      }}
      const unmatchedToggleBtn = document.getElementById('unmatchedToggleBtn');
      if (unmatchedToggleBtn) {{
        if (!hasMatchedColumns) {{
          unmatchedToggleBtn.disabled = true;
          unmatchedToggleBtn.style.opacity = '0.6';
          unmatchedToggleBtn.style.cursor = 'default';
          unmatchedToggleBtn.title = 'No matched topics to filter.';
        }} else {{
          unmatchedToggleBtn.addEventListener('click', toggleHideUnmatched);
        }}
      }}
      const savedUiState = loadUiState();
      const initialExpand = !!(savedUiState && savedUiState.topicTitleExpanded);
      const initialFlow = (savedUiState && savedUiState.matrixFlowMode === 'scroll') ? 'scroll' : 'fit';
      const initialHideUnmatched = hasMatchedColumns
        ? (savedUiState ? !!savedUiState.hideUnmatchedTopics : true)
        : false;
      setTopicTitleExpanded(initialExpand);
      setMatrixFlowMode(initialFlow);
      setHideUnmatched(initialHideUnmatched);
      initThemeMode();
      setInterval(refreshLockStatus, 2000);
      setInterval(refreshHomeIfChanged, 2000);
    </script>
    </body></html>
    """


def base_css() -> str:
    return """
    :root {
      color-scheme: light dark;
      --bg:#f3f6fb;
      --surface:#ffffff;
      --surface-2:#f8fbff;
      --text:#0f172a;
      --muted:#475569;
      --border:#ccd7e4;
      --border-2:#dbe4ef;
      --shadow:0 2px 9px rgba(0,0,0,.08);
      --row-hover:#e0efff;
      --row-hover-ring:#7fb3ff;
      --selected-col:rgba(16,185,129,.20);
      --topic-bg:#dbeafe;
      --topic-border:#8ab5ff;
      --topic-hover:rgba(37,99,235,.2);
      --topic-selected-unmatched-bg:#cfd8e3;
      --topic-selected-unmatched-border:#70839a;
      --topic-unselected-unmatched-bg:#f8fafc;
      --topic-unselected-unmatched-border:#e2e8f0;
      --score-0:#fee2e2;
      --score-1:#ffedd5;
      --score-2:#fef9c3;
      --score-3:#e0f2fe;
      --score-4:#dcfce7;
      --score-5:#bbf7d0;
    }
    @media (prefers-color-scheme: dark) {
      :root {
        --bg:#0b0b0b;
        --surface:#151515;
        --surface-2:#1b1b1b;
        --text:#e8e8e8;
        --muted:#a3a3a3;
        --border:#3a3a3a;
        --border-2:#2b2b2b;
        --shadow:0 2px 14px rgba(0,0,0,.45);
        --row-hover:#232323;
        --row-hover-ring:#646464;
        --selected-col:rgba(16,185,129,.24);
        --topic-bg:#2a2a2a;
        --topic-border:#4a4a4a;
        --topic-hover:rgba(120,120,120,.28);
        --topic-selected-unmatched-bg:#4b5563;
        --topic-selected-unmatched-border:#a3afbf;
        --topic-unselected-unmatched-bg:#20242b;
        --topic-unselected-unmatched-border:#394150;
        --score-0:#3a2626;
        --score-1:#3a2f24;
        --score-2:#3a3523;
        --score-3:#27323a;
        --score-4:#233427;
        --score-5:#1f3b2a;
      }
    }
    :root[data-theme="light"] {
      --bg:#f3f6fb;
      --surface:#ffffff;
      --surface-2:#f8fbff;
      --text:#0f172a;
      --muted:#475569;
      --border:#ccd7e4;
      --border-2:#dbe4ef;
      --shadow:0 2px 9px rgba(0,0,0,.08);
      --row-hover:#e0efff;
      --row-hover-ring:#7fb3ff;
      --selected-col:rgba(16,185,129,.20);
      --topic-bg:#dbeafe;
      --topic-border:#8ab5ff;
      --topic-hover:rgba(37,99,235,.2);
      --topic-selected-unmatched-bg:#cfd8e3;
      --topic-selected-unmatched-border:#70839a;
      --topic-unselected-unmatched-bg:#f8fafc;
      --topic-unselected-unmatched-border:#e2e8f0;
      --score-0:#fee2e2;
      --score-1:#ffedd5;
      --score-2:#fef9c3;
      --score-3:#e0f2fe;
      --score-4:#dcfce7;
      --score-5:#bbf7d0;
      color-scheme: light;
    }
    :root[data-theme="dark"] {
      --bg:#0b0b0b;
      --surface:#151515;
      --surface-2:#1b1b1b;
      --text:#e8e8e8;
      --muted:#a3a3a3;
      --border:#3a3a3a;
      --border-2:#2b2b2b;
      --shadow:0 2px 14px rgba(0,0,0,.45);
      --row-hover:#232323;
      --row-hover-ring:#646464;
      --selected-col:rgba(16,185,129,.24);
      --topic-bg:#2a2a2a;
      --topic-border:#4a4a4a;
      --topic-hover:rgba(120,120,120,.28);
      --topic-selected-unmatched-bg:#4b5563;
      --topic-selected-unmatched-border:#a3afbf;
      --topic-unselected-unmatched-bg:#20242b;
      --topic-unselected-unmatched-border:#394150;
      --score-0:#3a2626;
      --score-1:#3a2f24;
      --score-2:#3a3523;
      --score-3:#27323a;
      --score-4:#233427;
      --score-5:#1f3b2a;
      color-scheme: dark;
    }
    body { font-family: Arial, sans-serif; background:var(--bg); color:var(--text); margin:0; }
    .container { max-width:1200px; margin:20px auto; background:var(--surface); border-radius:12px; padding:18px; box-shadow:var(--shadow); }
    .home-layout { background:transparent; border-radius:0; box-shadow:none; padding:0; }
    .card { border:1px solid var(--border-2); border-radius:8px; padding:12px; margin:12px 0; }
    button, .button-link { background:#2563eb; color:white; border:none; border-radius:6px; padding:8px 12px; text-decoration:none; cursor:pointer; transition:transform .12s ease, filter .12s ease, box-shadow .12s ease; }
    .button-danger { background:#b91c1c; }
    .top-right-actions { position:fixed; top:14px; right:14px; z-index:1000; display:flex; gap:8px; align-items:center; }
    button:hover, .button-link:hover, button:focus-visible, .button-link:focus-visible { filter:brightness(1.12); box-shadow:0 0 0 3px rgba(37,99,235,.22), 0 6px 16px rgba(37,99,235,.25); outline:none; }
    .button-danger:hover, .button-danger:focus-visible { box-shadow:0 0 0 3px rgba(185,28,28,.24), 0 6px 16px rgba(185,28,28,.3); }
    button.pressed { transform:translateY(1px) scale(.98); filter:brightness(.9); }
    .muted { color:var(--muted); }
    .buckets { display:flex; flex-direction:column; gap:10px; }
    .bucket { background:var(--surface-2); border:2px dashed var(--border); min-height:90px; border-radius:8px; padding:8px; }
    .bucket h3 { margin:0 0 8px 0; font-size:14px; }
    .bucket-topics { display:flex; flex-wrap:wrap; gap:6px; }
    .topic { background:var(--topic-bg); border:1px solid var(--topic-border); border-radius:16px; padding:6px 10px; margin-bottom:6px; font-size:12px; cursor:grab; display:inline-flex; align-items:center; gap:5px; }
    .topic:hover { filter:brightness(1.06); box-shadow:0 0 0 2px rgba(37,99,235,.18), 0 4px 10px var(--topic-hover); }
    .topic-main-match { font-weight:700; }
    .topic-shadow-match { opacity:.66; }
    .topic-selected-unmatched {
      background:var(--topic-selected-unmatched-bg) !important;
      border-color:var(--topic-selected-unmatched-border) !important;
      color:var(--text) !important;
      opacity:.96;
      box-shadow:inset 0 0 0 1px color-mix(in srgb, var(--topic-selected-unmatched-border) 85%, transparent);
    }
    .topic-unselected-unmatched {
      background:var(--topic-unselected-unmatched-bg) !important;
      border-color:var(--topic-unselected-unmatched-border) !important;
      color:var(--muted) !important;
      opacity:.42;
      box-shadow:none;
    }
    .row { display:flex; gap:14px; flex-wrap:wrap; }
    .col { flex:1; min-width:320px; }
    .inline-finalize { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-top:8px; }
    .theme-toggle { display:inline-flex; align-items:center; gap:0; border:1px solid var(--border); border-radius:8px; overflow:hidden; margin-left:8px; vertical-align:middle; }
    .theme-toggle-btn { background:var(--surface-2); color:var(--text); border:none; border-right:1px solid var(--border); border-radius:0; padding:6px 10px; }
    .theme-toggle-btn:last-child { border-right:none; }
    .theme-toggle-btn:hover, .theme-toggle-btn:focus-visible { filter:none; box-shadow:none; background:var(--row-hover); }
    .theme-toggle-btn.active { background:#2563eb; color:#fff; }
    table { border-collapse:collapse; width:100%; }
    th,td { border:1px solid var(--border); padding:6px; text-align:left; }
    .matrix-table { table-layout:fixed; width:100%; border-collapse:separate; border-spacing:0; }
    .matrix-table th, .matrix-table td {
      border-top:0;
      border-left:0;
      border-right:1px solid var(--border);
      border-bottom:1px solid var(--border);
    }
    .matrix-table thead th { border-top:1px solid var(--border); }
    .matrix-table th:first-child, .matrix-table td:first-child { border-left:1px solid var(--border); }
    .matrix-table .student-col { width:clamp(104px, 16vw, 160px); }
    .matrix-table .topic-col { width:auto; }
    .matrix-table .student-head, .matrix-table .student-cell { white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .matrix-table .student-cell { display:flex; align-items:center; gap:6px; }
    .matrix-table .student-name { min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:700; }
    .matrix-table .student-main-group { display:inline-flex; align-items:center; min-width:0; }
    .matrix-table .student-group-right { margin-left:auto; display:inline-flex; align-items:center; gap:4px; min-width:0; }
    .matrix-table .topic-head { text-align:center; font-size:11px; padding:3px 1px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; color:var(--text); }
    .matrix-table .group-head { font-weight:700; }
    .matrix-table .topic-group { font-size:9px; line-height:1; font-weight:700; }
    .matrix-table .topic-code { font-weight:700; line-height:1.05; }
    .matrix-table .topic-code-full { font-size:10px; font-weight:600; color:var(--muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .matrix-table .topic-title { display:none; }
    .matrix-controls { display:flex; justify-content:flex-end; gap:8px; margin:6px 0 8px 0; }
    .matrix-wrap { width:100%; overflow-x:hidden; }
    .matrix-wrap.matrix-scroll { overflow-x:auto; }
    .matrix-wrap.matrix-fit { overflow-x:hidden; }
    .matrix-wrap.table-mode-compressed .topic-code-short { display:block; }
    .matrix-wrap.table-mode-compressed .topic-code-full { display:none; }
    .matrix-wrap.table-mode-all .topic-code-short { display:none; }
    .matrix-wrap.table-mode-all .topic-code-full { display:block; }
    .matrix-wrap.hide-unmatched .unmatched-topic-col,
    .matrix-wrap.hide-unmatched .unmatched-topic-cell { display:none; }
    .matrix-wrap.matrix-scroll .matrix-table {
      table-layout:auto;
      width:max-content;
      min-width:100%;
    }
    .matrix-wrap.matrix-fit .matrix-table {
      table-layout:fixed;
      width:100%;
      min-width:0;
    }
    .matrix-wrap.matrix-fit .topic-head,
    .matrix-wrap.hide-unmatched .topic-head {
      white-space:normal;
      overflow:visible;
      text-overflow:clip;
      overflow-wrap:anywhere;
      word-break:break-word;
    }
    .matrix-wrap.matrix-fit .topic-code-full,
    .matrix-wrap.hide-unmatched .topic-code-full {
      white-space:normal;
      overflow:visible;
      text-overflow:clip;
      overflow-wrap:anywhere;
      word-break:break-word;
    }
    .matrix-wrap.hide-unmatched {
      overflow-x:hidden;
    }
    .matrix-wrap.hide-unmatched .matrix-table {
      table-layout:fixed;
      width:100%;
      min-width:0;
    }
    .matrix-table .score-cell { text-align:center; font-weight:600; font-size:11px; padding:4px 1px; }
    .shadow-pref-circle {
      display:inline-flex;
      align-items:center;
      justify-content:center;
      width:1.55em;
      height:1.55em;
      border:2px solid var(--shadow-gcolor, #64748b);
      border-radius:999px;
      line-height:1;
      box-sizing:border-box;
    }
    .group-chip { display:inline-block; font-size:10px; font-weight:700; color:#0f766e; background:#ccfbf1; border:1px solid #99f6e4; border-radius:999px; padding:1px 6px; vertical-align:middle; position:relative; line-height:1.1; }
    .shadow-chip { font-size:7px; padding:1px 4px; }
    .shadow-violation-star { color:#ef4444; font-size:12px; font-weight:900; line-height:1; display:inline-block; }
    .active-pair-cell { font-weight:700; background-clip:padding-box; }
    .group-box-top, .group-box-mid, .group-box-bottom, .group-box-single {
      border-left:3px solid var(--gcolor, #22c55e) !important;
      border-right:3px solid var(--gcolor, #22c55e) !important;
    }
    .group-box-top {
      border-top:3px solid var(--gcolor, #22c55e) !important;
      border-top-left-radius:8px;
      border-top-right-radius:8px;
    }
    .group-box-bottom {
      border-bottom:3px solid var(--gcolor, #22c55e) !important;
      border-bottom-left-radius:8px;
      border-bottom-right-radius:8px;
    }
    .group-box-single {
      border-top:3px solid var(--gcolor, #22c55e) !important;
      border-bottom:3px solid var(--gcolor, #22c55e) !important;
      border-radius:8px;
    }
    .matrix-row:hover .active-pair-cell { filter:brightness(1.06); }
    .matrix-row { cursor:pointer; transition:background-color .12s ease, box-shadow .12s ease; }
    .matrix-row:hover { background:var(--row-hover); box-shadow:inset 0 0 0 2px var(--row-hover-ring); }
    .matrix-row-dim .student-cell { opacity:.40; }
    .matrix-row-dim .score-cell { opacity:.40; }
    .matrix-row-dim .active-pair-cell { opacity:1; }
    .score-cell { text-align:center; font-weight:600; }
    .score-0 { background:var(--score-0); }
    .score-1 { background:var(--score-1); }
    .score-2 { background:var(--score-2); }
    .score-3 { background:var(--score-3); }
    .score-4 { background:var(--score-4); }
    .score-5 { background:var(--score-5); }
    """


def render_student(sid: int) -> str:
    conn = db_conn()
    class_id = get_current_class_id(conn)
    n = int(get_class_meta(conn, class_id, "n", "8"))
    students = get_students(conn, class_id)
    sid = max(0, min(sid, len(students) - 1))
    pref = get_pref_row(conn, class_id, sid, n)
    finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
    latest_id_row = conn.execute(
        """
        SELECT mr.id
        FROM match_runs mr
        WHERE mr.class_id=?
          AND mr.status IN ('optimal','feasible','interrupted')
          AND EXISTS (SELECT 1 FROM selected_topics st WHERE st.run_id=mr.id)
        ORDER BY mr.id DESC
        LIMIT 1
        """,
        (class_id,),
    ).fetchone()
    assignment = None
    assignment_visual = {"main_topic": None, "shadow_topic": None, "main_color": "", "shadow_color": "", "selected_topics": []}
    if latest_id_row:
        run_id = int(latest_id_row[0])
        selected_topic_ids_ordered = [
            int(r[0])
            for r in conn.execute(
                """
                SELECT topic_id
                FROM selected_topics
                WHERE run_id=?
                ORDER BY CASE partition WHEN 'A' THEN 0 WHEN 'B' THEN 1 ELSE 2 END, topic_id
                """,
                (run_id,),
            ).fetchall()
        ]
        assignment = conn.execute("SELECT * FROM assignments WHERE run_id=? AND student_id=?", (run_id, sid)).fetchone()
        assignment_visual["selected_topics"] = selected_topic_ids_ordered
        if assignment:
            group_rank = {tid: idx for idx, tid in enumerate(selected_topic_ids_ordered)}
            group_palette = [
                "#22c55e", "#06b6d4", "#f59e0b", "#a78bfa", "#ef4444", "#10b981",
                "#3b82f6", "#eab308", "#ec4899", "#14b8a6", "#84cc16", "#f97316",
            ]

            def color_for_gid(gid: int) -> str:
                return group_palette[gid % len(group_palette)]

            main_topic_id = int(assignment["main_topic"])
            shadow_topic_id = int(assignment["shadow_topic"])
            main_gid = group_rank.get(main_topic_id, -1)
            shadow_gid = group_rank.get(shadow_topic_id, -1)
            assignment_visual = {
                "main_topic": main_topic_id,
                "shadow_topic": shadow_topic_id,
                "main_color": color_for_gid(main_gid) if main_gid >= 0 else "#2563eb",
                "shadow_color": color_for_gid(shadow_gid) if shadow_gid >= 0 else "#64748b",
                "selected_topics": selected_topic_ids_ordered,
            }
    conn.close()

    topics_json = json.dumps([{"id": s["id"], "short": f"S{s['id'] + 1}", "name": s["name"], "title": s["topic_title"]} for s in students])
    pref_json = json.dumps(pref)
    editable = "false" if finalized else "true"
    assignment_html = ""
    if assignment:
        assignment_html = f"<p><strong>Your assignment:</strong> Main = {escape(assignment['main_title'])} (score {assignment['main_score']}), Shadow = {escape(assignment['shadow_title'])} (score {assignment['shadow_score']}).</p>"

    return f"""
    <html><head><title>Student {sid + 1}</title><style>{base_css()}</style></head><body>
    <main class='container'>
      <p><a href='{app_url('/')}'>← back</a></p>
      <h1>Student {sid + 1} preferences</h1>
      <p class='muted'>Drag topics to change their rankings. At most {n // 4} topics can be given score zero, which corresponds to a veto.</p>
      <p>Editing enabled: <strong id='editingState'>{'Yes' if not finalized else 'No (finalized by admin)'}</strong></p>
      {assignment_html}
      <div class='card'>
        <h2>Your Info</h2>
        <label>
          Student name:
          <input id='studentNameInput' type='text' value='{escape(students[sid]["name"])}' maxlength='200' style='width:min(520px, 100%); margin-right:8px;'>
        </label>
        <span id='nameMsg' class='muted'></span>
        <br><br>
        <label>
          Topic title:
          <input id='topicTitleInput' type='text' value='{escape(students[sid]["topic_title"])}' maxlength='200' style='width:min(520px, 100%); margin-right:8px;'>
        </label>
        <span id='titleMsg' class='muted'></span>
      </div>
      <div id='app'></div>
      <p class='muted'>Preferences save automatically.</p>
      <span id='msg' class='muted'></span>
    </main>
    <script>
      const BASE_PATH = {json.dumps(APP_BASE_PATH)};
      function appUrl(path) {{
        const raw = String(path || '/');
        const normalized = raw.startsWith('/') ? raw : '/' + raw;
        return BASE_PATH ? BASE_PATH + normalized : normalized;
      }}
      const THEME_MODE_KEY = 'adminThemeMode';
      function applyThemeMode(mode) {{
        const root = document.documentElement;
        const next = (mode === 'light' || mode === 'dark') ? mode : 'system';
        if (next === 'system') root.removeAttribute('data-theme');
        else root.setAttribute('data-theme', next);
      }}
      function initThemeMode() {{
        let mode = 'system';
        try {{
          const stored = localStorage.getItem(THEME_MODE_KEY);
          if (stored) mode = stored;
        }} catch (_err) {{}}
        applyThemeMode(mode);
      }}
      const sid = {sid};
      let editable = {editable};
      let topics = {topics_json};
      let scores = {pref_json};
      const assignmentVisual = {json.dumps(assignment_visual)};
      const selectedTopicSet = new Set(Array.isArray(assignmentVisual.selected_topics) ? assignmentVisual.selected_topics : []);
      const vetoMax = Math.floor(topics.length / 4);
      const prefMsg = document.getElementById('msg');
      let prefSaveTimer = null;
      let prefSaveInFlight = false;
      let prefSaveQueued = false;
      let lastSavedScores = JSON.stringify(scores);

      function findTopicById(topicId) {{
        return topics.find(t => t.id === topicId);
      }}
      function hexToRgba(hexColor, alpha) {{
        const h = String(hexColor || '').replace('#', '');
        if (h.length !== 6) return `rgba(37,99,235,${{alpha}})`;
        const r = parseInt(h.slice(0,2), 16);
        const g = parseInt(h.slice(2,4), 16);
        const b = parseInt(h.slice(4,6), 16);
        return `rgba(${{r}},${{g}},${{b}},${{alpha}})`;
      }}

      function applyEditableState() {{
        const nameInput = document.getElementById('studentNameInput');
        const titleInput = document.getElementById('topicTitleInput');
        const editingState = document.getElementById('editingState');
        nameInput.disabled = !editable;
        titleInput.disabled = !editable;
        if (!editable) {{
          editingState.innerText = 'No (finalized by admin)';
          prefMsg.innerText = 'Matching finalized; edits are locked.';
        }} else {{
          editingState.innerText = 'Yes';
          if (prefMsg.innerText === 'Matching finalized; edits are locked.') {{
            prefMsg.innerText = '';
          }}
        }}
      }}

      function render() {{
        const app = document.getElementById('app');
        const buckets = [];
        for (let s=5; s>=0; s--) {{
          const title = (s === 0)
            ? `Score 0 <span id='vetoCount' class='muted' style='font-weight:500; font-size:12px;'>(0 / ${{vetoMax}} vetoes)</span>`
            : `Score ${{s}}`;
          buckets.push(`<div class='bucket' data-score='${{s}}' ondragover='event.preventDefault()' ondrop='dropTopic(event, ${{s}})'><h3>${{title}}</h3><div id='bucket-${{s}}' class='bucket-topics'></div></div>`);
        }}
        app.innerHTML = `<div class='buckets'>${{buckets.join('')}}</div>`;
        const orderedTopics = [...topics].sort((a, b) => {{
          if (scores[b.id] !== scores[a.id]) return scores[b.id] - scores[a.id];
          return a.id - b.id;
        }});
        orderedTopics.forEach(t => {{
          const div = document.createElement('div');
          div.className='topic';
          if (assignmentVisual.main_topic === t.id) {{
            div.classList.add('topic-main-match');
            if (assignmentVisual.main_color) {{
              div.style.borderColor = assignmentVisual.main_color;
              div.style.background = hexToRgba(assignmentVisual.main_color, 0.24);
              div.style.boxShadow = `0 0 0 2px ${{hexToRgba(assignmentVisual.main_color, 0.32)}}`;
            }}
          }} else if (assignmentVisual.shadow_topic === t.id) {{
            div.classList.add('topic-shadow-match');
            if (assignmentVisual.shadow_color) {{
              div.style.borderColor = assignmentVisual.shadow_color;
              div.style.background = hexToRgba(assignmentVisual.shadow_color, 0.18);
              div.style.boxShadow = `0 0 0 1px ${{hexToRgba(assignmentVisual.shadow_color, 0.28)}}`;
            }}
          }} else if (selectedTopicSet.size > 0) {{
            if (selectedTopicSet.has(t.id)) {{
              div.classList.add('topic-selected-unmatched');
            }} else {{
              div.classList.add('topic-unselected-unmatched');
            }}
          }}
          div.draggable = editable;
          div.title = `${{t.short}}: ${{t.title}} (${{t.name}})`;
          div.innerText = `${{t.short}}: ${{t.title}} (${{t.name}})`;
          div.dataset.topicId = String(t.id);
          div.ondragstart = ev => {{
            ev.dataTransfer.effectAllowed = 'move';
            ev.dataTransfer.setData('text/topic', String(t.id));
            const ghost = div.cloneNode(true);
            ghost.style.position = 'absolute';
            ghost.style.top = '-1000px';
            ghost.style.left = '-1000px';
            ghost.style.opacity = '1';
            ghost.style.boxShadow = '0 4px 12px rgba(0,0,0,.25)';
            document.body.appendChild(ghost);
            ev.dataTransfer.setDragImage(ghost, Math.floor(ghost.offsetWidth / 2), Math.floor(ghost.offsetHeight / 2));
            setTimeout(() => ghost.remove(), 0);
          }};
          document.getElementById(`bucket-${{scores[t.id]}}`).appendChild(div);
        }});
        updateVetoCount();
      }}

      function updateVetoCount() {{
        const v = scores.filter(x => x === 0).length;
        const el = document.getElementById('vetoCount');
        if (el) el.innerText = `(${{v}} / ${{vetoMax}} vetoes)`;
      }}

      function dropTopic(ev, targetScore) {{
        if (!editable) return;
        const topicId = Number(ev.dataTransfer.getData('text/topic'));
        if (topicId === sid && targetScore < 4) return alert('Your own topic must stay at least 4.');
        const current = scores[topicId];
        if (current === targetScore) return;
        const currentVetoes = scores.filter(x => x === 0).length;
        if (targetScore === 0 && currentVetoes >= vetoMax) return alert('Maximum vetoes reached. Move one out of 0 first.');
        scores[topicId] = targetScore;
        render();
        schedulePreferenceSave();
      }}

      async function savePreferencesNow(force=false) {{
        if (!editable) return;
        const serialized = JSON.stringify(scores);
        if (!force && serialized === lastSavedScores) return;
        if (prefSaveInFlight) {{
          prefSaveQueued = true;
          return;
        }}
        prefSaveInFlight = true;
        prefMsg.innerText = 'Saving...';
        try {{
          const res = await fetch(appUrl(`/api/student/${{sid}}/preferences`), {{
            method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{scores}})
          }});
          const data = await res.json();
          if (res.ok && !data.error) {{
            lastSavedScores = serialized;
          }}
          prefMsg.innerText = data.message || data.error || (res.ok ? 'Saved' : 'Save failed');
        }} catch (err) {{
          prefMsg.innerText = 'Save failed';
        }} finally {{
          prefSaveInFlight = false;
          if (prefSaveQueued) {{
            prefSaveQueued = false;
            await savePreferencesNow(true);
          }}
        }}
      }}

      function schedulePreferenceSave() {{
        if (!editable) return;
        if (prefSaveTimer) clearTimeout(prefSaveTimer);
        prefSaveTimer = setTimeout(() => {{
          prefSaveTimer = null;
          savePreferencesNow();
        }}, 350);
      }}

      function wireAutosaveField(inputId, msgId, endpoint, payloadKey, onLocalValueChange) {{
        const input = document.getElementById(inputId);
        const msg = document.getElementById(msgId);
        let lastSaved = (input.value || '').trim();
        let timer = null;
        let inFlight = false;

        async function saveNow() {{
          if (!editable || inFlight) return;
          const value = (input.value || '').trim();
          if (value === lastSaved) return;
          inFlight = true;
          msg.innerText = 'Saving...';
          try {{
            const url = appUrl(endpoint);
            const res = await fetch(url, {{
              method:'POST',
              headers:{{'Content-Type':'application/json'}},
              body:JSON.stringify({{ [payloadKey]: value }})
            }});
            const data = await res.json();
            if (res.ok && !data.error) {{
              lastSaved = value;
            }}
            msg.innerText = data.message || data.error || (res.ok ? 'Saved' : 'Save failed');
          }} catch (err) {{
            msg.innerText = 'Save failed';
          }} finally {{
            inFlight = false;
          }}
        }}

        input.addEventListener('input', () => {{
          if (onLocalValueChange) onLocalValueChange(input.value || '');
          if (timer) clearTimeout(timer);
          timer = setTimeout(saveNow, 450);
        }});
        input.addEventListener('blur', async () => {{
          if (timer) {{
            clearTimeout(timer);
            timer = null;
          }}
          await saveNow();
        }});
      }}

      wireAutosaveField('studentNameInput', 'nameMsg', `/api/student/${{sid}}/name`, 'name', (value) => {{
        const own = findTopicById(sid);
        if (!own) return;
        own.name = value;
        render();
      }});
      wireAutosaveField('topicTitleInput', 'titleMsg', `/api/student/${{sid}}/topic_title`, 'title', (value) => {{
        const own = findTopicById(sid);
        if (!own) return;
        own.title = value;
        render();
      }});

      async function refreshStudentMeta() {{
        try {{
          const res = await fetch(appUrl('/api/students_meta'));
          if (!res.ok) return;
          const data = await res.json();
          if (!Array.isArray(data.students)) return;
          const nextEditable = !Boolean(data.finalized);
          if (editable !== nextEditable) {{
            editable = nextEditable;
            applyEditableState();
            render();
          }}
          let changed = false;
          data.students.forEach(s => {{
            const id = Number(s.id);
            if (id === sid) return;
            const t = findTopicById(id);
            if (!t) return;
            const nextName = String(s.name ?? '');
            const nextTitle = String(s.topic_title ?? '');
            if (t.name !== nextName || t.title !== nextTitle) {{
              t.name = nextName;
              t.title = nextTitle;
              changed = true;
            }}
          }});
          if (changed) render();
        }} catch (err) {{
          // keep current values if refresh fails
        }}
      }}
      setInterval(refreshStudentMeta, 2000);

      initThemeMode();
      applyEditableState();
      render();
    </script>
    </body></html>
    """


def render_admin() -> str:
    return f"""
    <html><head><title>Admin</title><style>{base_css()}</style></head><body>
    <main class='container'>
      <p><a href='{app_url('/')}'>← back</a></p>
      <h1>Admin dashboard</h1>
      <div class='top-right-actions'>
        <a class='button-link button-danger' href='{app_url("/shutdown")}'>Stop server</a>
      </div>
      <div style='margin:12px 0;'>
        <div class='inline-finalize'>
          <button id='finalCtrlBtn' onclick='toggleFinalized(this)'>Freeze student preferences</button>
          <span id='finalStateText' class='muted'>Student preferences are currently editable.</span>
        </div>
      </div>
      <div style='margin:12px 0;'>
        <button id='runCtrlBtn' onclick='toggleRun(this)'>Run matching</button>
        <button id='undoMatchBtn' onclick='undoMatching(this)' style='display:none;'>Undo matching</button>
        <p id='status' class='muted'></p>
        <p id='pkgWarn' class='muted' style='display:none;color:#b91c1c;'></p>
        <p id='adminMsg' class='muted'></p>
      </div>
      <section class='card'>
        <h2>Class Management</h2>
        <p>
          <label>Class name:
            <input id='classNameInput' type='text' maxlength='200' style='width:min(420px, 100%);'>
          </label>
          <span id='classNameMsg' class='muted'></span>
        </p>
        <p>
          <label>Select current class:
            <select id='classSelect' onchange='onClassSelectChanged(this)'></select>
          </label>
          <button type='button' onclick='createNewClass(this)'>Create new class</button>
          <button type='button' class='button-danger' onclick='deleteCurrentClass(this)'>Delete current class</button>
        </p>
        <p><strong>Database management</strong></p>
        <p>
          <button onclick='addStudents(this)'>Add students</button>
          <button onclick='removeStudents(this)'>Remove students</button>
          <button onclick='resetDb(this)'>Reset database</button>
          <button onclick='randomizePreferences(this)'>Replace with random preferences</button>
          <label>Random preference distribution:
            <select id='randomPrefModeSelect'>
              <option value='category_uniform'>Category (uniform)</option>
              <option value='category'>Category (favorite-based)</option>
              <option value='category_mode3'>Category (peaked at 3)</option>
              <option value='random'>Random</option>
            </select>
          </label>
        </p>
        <p style='margin-top:10px;'>
          <button onclick='downloadCsv()'>Save CSV</button>
          <input id='csvFileInput' type='file' accept='.csv,text/csv'>
          <button onclick='importCsv(this)'>Load CSV</button>
        </p>
      </section>
      <p style='margin-top:10px;'>
        <span>Theme:</span>
        <span class='theme-toggle' role='group' aria-label='Theme mode'>
          <button type='button' id='themeBtnSystem' class='theme-toggle-btn' onclick='onThemeModeChange("system")'>System</button>
          <button type='button' id='themeBtnLight' class='theme-toggle-btn' onclick='onThemeModeChange("light")'>Light</button>
          <button type='button' id='themeBtnDark' class='theme-toggle-btn' onclick='onThemeModeChange("dark")'>Dark</button>
        </span>
      </p>
      <section id='progressSection' class='card' style='display:none;'>
        <h2 style='display:flex;align-items:center;justify-content:space-between;gap:10px;'>
          <span>Solver progress</span>
          <button id='autoFollowBtn' type='button' onclick='toggleAutoFollow(this)'>Auto-follow: On</button>
        </h2>
        <div id='progress'></div>
      </section>
      <section id='resultsSection' class='card' style='display:none;'>
        <h2>Matching results</h2>
        <div id='results'></div>
      </section>
    </main>
    <script>
      const BASE_PATH = {json.dumps(APP_BASE_PATH)};
      function appUrl(path) {{
        const raw = String(path || '/');
        const normalized = raw.startsWith('/') ? raw : '/' + raw;
        return BASE_PATH ? BASE_PATH + normalized : normalized;
      }}
      const THEME_MODE_KEY = 'adminThemeMode';
      function applyThemeMode(mode) {{
        const root = document.documentElement;
        let next = String(mode || 'system');
        if (next !== 'light' && next !== 'dark') next = 'system';
        if (next === 'system') {{
          root.removeAttribute('data-theme');
        }} else {{
          root.setAttribute('data-theme', next);
        }}
        const map = {{
          system: document.getElementById('themeBtnSystem'),
          light: document.getElementById('themeBtnLight'),
          dark: document.getElementById('themeBtnDark'),
        }};
        Object.entries(map).forEach(([k, el]) => {{
          if (!el) return;
          if (k === next) el.classList.add('active');
          else el.classList.remove('active');
        }});
        try {{ localStorage.setItem(THEME_MODE_KEY, next); }} catch (err) {{}}
      }}
      function onThemeModeChange(mode) {{
        applyThemeMode(mode || 'system');
      }}
      function initThemeMode() {{
        let mode = 'system';
        try {{
          const stored = localStorage.getItem(THEME_MODE_KEY);
          if (stored) mode = stored;
        }} catch (err) {{}}
        applyThemeMode(mode);
      }}
      async function post(path, body={{}}) {{
        const r = await fetch(appUrl(path), {{method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify(body)}});
        const data = await r.json();
        data.__ok = r.ok;
        return data;
      }}
      function pulseButton(btn) {{
        if (!btn) return;
        btn.classList.add('pressed');
        setTimeout(() => btn.classList.remove('pressed'), 150);
      }}
      function setAdminMsg(text, isError=false) {{
        const el = document.getElementById('adminMsg');
        el.style.color = isError ? '#b91c1c' : '#334155';
        el.innerText = text;
      }}
      function setPackageWarning(missingPackages, latestError) {{
        const el = document.getElementById('pkgWarn');
        if (!el) return;
        if (Array.isArray(missingPackages) && missingPackages.length > 0) {{
          el.style.display = 'block';
          el.innerText = `Missing package(s): ${{missingPackages.join(', ')}}. Install into the Python env running app.py.`;
          return;
        }}
        if (latestError) {{
          const low = String(latestError).toLowerCase();
          if (low.includes('no module named')) {{
            el.style.display = 'block';
            el.innerText = `Dependency warning: ${{latestError}}`;
            return;
          }}
        }}
        el.style.display = 'none';
        el.innerText = '';
      }}
      function populateClassSelect(classes, currentClassId) {{
        const select = document.getElementById('classSelect');
        if (!select) return;
        if (!Array.isArray(classes) || classes.length === 0) return;
        const prev = select.value;
        select.innerHTML = '';
        (classes || []).forEach(c => {{
          const opt = document.createElement('option');
          opt.value = String(c.id);
          opt.textContent = c.name || `Class ${{c.id}}`;
          select.appendChild(opt);
        }});
        const target = String(currentClassId ?? '');
        if (target && Array.from(select.options).some(o => o.value === target)) {{
          select.value = target;
        }} else if (prev && Array.from(select.options).some(o => o.value === prev)) {{
          select.value = prev;
        }}
      }}
      async function createNewClass(btn) {{
        pulseButton(btn);
        const name = prompt('New class name:');
        if (name === null) {{
          await poll();
          return;
        }}
        const nRaw = prompt('Number of students:', '8');
        if (nRaw === null) {{
          await poll();
          return;
        }}
        const n = Math.max(4, Number(nRaw) || 8);
        const data = await post('/api/admin/create_class', {{name, n}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function deleteCurrentClass(btn) {{
        pulseButton(btn);
        const select = document.getElementById('classSelect');
        if (!select) return;
        const currentId = Number(select.value || 0);
        const currentName = select.options[select.selectedIndex] ? String(select.options[select.selectedIndex].text || '').trim() : `Class ${{currentId}}`;
        if (!Number.isFinite(currentId) || currentId <= 0) {{
          setAdminMsg('No valid class selected.', true);
          return;
        }}
        const ok = window.confirm(
          `Delete current class "${{currentName}}"?\\n\\nThis removes all students, preferences, and matching history for this class.`
        );
        if (!ok) return;
        const data = await post('/api/admin/delete_class', {{class_id: currentId}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function onClassSelectChanged(selectEl) {{
        const value = String(selectEl.value || '');
        const classId = Number(value);
        if (!Number.isFinite(classId) || classId <= 0) return;
        const data = await post('/api/admin/select_class', {{class_id: classId}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      function wireClassNameAutosave() {{
        const input = document.getElementById('classNameInput');
        const msg = document.getElementById('classNameMsg');
        if (!input || !msg) return;
        let timer = null;
        let inFlight = false;
        let lastSaved = (input.value || '').trim();

        async function saveNow() {{
          if (inFlight) return;
          const value = (input.value || '').trim();
          if (!value || value === lastSaved) return;
          inFlight = true;
          msg.innerText = 'Saving...';
          const data = await post('/api/admin/set_class_name', {{name: value}});
          if (data.__ok && !data.error) {{
            lastSaved = value;
          }}
          msg.innerText = data.message || data.error || (data.__ok ? 'Saved' : 'Save failed');
          inFlight = false;
          await poll();
        }}

        input.addEventListener('input', () => {{
          if (timer) clearTimeout(timer);
          timer = setTimeout(saveNow, 450);
        }});
        input.addEventListener('blur', async () => {{
          if (timer) {{
            clearTimeout(timer);
            timer = null;
          }}
          await saveNow();
        }});

        window.__setClassNameSavedValue = (value) => {{
          const v = String(value || '').trim();
          lastSaved = v;
          if (document.activeElement !== input) input.value = v;
        }};
      }}
      async function addStudents(btn) {{
        pulseButton(btn);
        const raw = prompt('How many students to add?', '1');
        if (raw === null) return;
        const count = Number(raw);
        if (!Number.isFinite(count) || count <= 0 || !Number.isInteger(count)) {{
          setAdminMsg('Enter a positive whole number of students to add.', true);
          return;
        }}
        const data = await post('/api/admin/add_students', {{count}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function removeStudents(btn) {{
        pulseButton(btn);
        let students = [];
        try {{
          const r = await fetch(appUrl('/api/students_meta'));
          if (r.ok) {{
            const data = await r.json();
            students = (data.students || []).slice().sort((a,b) => Number(a.id) - Number(b.id));
          }}
        }} catch (err) {{
          // fall through
        }}
        if (!students.length) {{
          setAdminMsg('No students available to remove.', true);
          return;
        }}
        const roster = students.map(s => `S${{Number(s.id) + 1}}: ${{String(s.name || '').trim() || '(unnamed)'}}`).join('\\n');
        const raw = prompt(
          `Enter student numbers to remove (comma-separated).\\n\\n${{roster}}\\n\\nExample: 2,5,9`,
          ''
        );
        if (raw === null) return;
        const ids = Array.from(new Set(
          raw.split(',')
            .map(x => Number(String(x).trim()) - 1)
            .filter(x => Number.isInteger(x) && x >= 0)
        ));
        if (!ids.length) {{
          setAdminMsg('No valid student numbers provided.', true);
          return;
        }}
        const selected = students.filter(s => ids.includes(Number(s.id)));
        if (!selected.length) {{
          setAdminMsg('None of the selected students exist.', true);
          return;
        }}
        const label = selected.map(s => `S${{Number(s.id) + 1}} (${{String(s.name || '').trim() || 'unnamed'}})`).join(', ');
        const ok = window.confirm(
          `Remove ${{selected.length}} student(s): ${{label}}?\\n\\nTheir topics and preferences will be deleted.`
        );
        if (!ok) return;
        const data = await post('/api/admin/remove_students', {{student_ids: selected.map(s => Number(s.id))}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function randomizePreferences(btn) {{
        pulseButton(btn);
        const mode = String(document.getElementById('randomPrefModeSelect').value || 'category_uniform');
        const ok = window.confirm(
          'Replace current class names, topics, and preferences with randomized data?\\n\\nThis will clear existing matching results.'
        );
        if (!ok) return;
        const data = await post('/api/admin/randomize_preferences', {{mode}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      function downloadCsv() {{
        window.location.href = appUrl('/api/admin/export_csv');
      }}
      async function importCsv(btn) {{
        pulseButton(btn);
        const input = document.getElementById('csvFileInput');
        const file = input && input.files && input.files[0] ? input.files[0] : null;
        if (!file) {{
          setAdminMsg('Choose a CSV file first.', true);
          return;
        }}
        let text = '';
        try {{
          text = await file.text();
        }} catch (err) {{
          setAdminMsg('Could not read CSV file.', true);
          return;
        }}
        const data = await post('/api/admin/import_csv', {{csv_text: text}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        if (data.__ok && input) input.value = '';
        await poll();
      }}
      async function resetDb(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/reset');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
      }}
      function setFinalizedControl(finalized) {{
        const btn = document.getElementById('finalCtrlBtn');
        const text = document.getElementById('finalStateText');
        if (!btn || !text) return;
        if (finalized) {{
          btn.dataset.finalized = '1';
          btn.innerText = 'Unfreeze student preferences';
          btn.style.background = '#0f766e';
          text.innerText = 'Student preferences are currently frozen.';
        }} else {{
          btn.dataset.finalized = '0';
          btn.innerText = 'Freeze student preferences';
          btn.style.background = '';
          text.innerText = 'Student preferences are currently editable.';
        }}
      }}
      async function toggleFinalized(btn) {{
        pulseButton(btn);
        const currentlyFinalized = (btn.dataset.finalized || '0') === '1';
        const data = await post('/api/admin/finalize', {{finalized: !currentlyFinalized}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function runMatch(btn) {{
        pulseButton(btn);
        const progressSection = document.getElementById('progressSection');
        const progress = document.getElementById('progress');
        progressSection.style.display = 'block';
        progress.innerText = 'Starting run...';
        const data = await post('/api/admin/run');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function interruptMatch(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/interrupt');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function toggleRun(btn) {{
        if ((btn.dataset.mode || 'run') === 'interrupt') {{
          await interruptMatch(btn);
        }} else {{
          await runMatch(btn);
        }}
      }}
      async function undoMatching(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/undo_matching');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      function setRunControlMode(running) {{
        const btn = document.getElementById('runCtrlBtn');
        if (!btn) return;
        if (running) {{
          btn.dataset.mode = 'interrupt';
          btn.innerText = 'Interrupt run';
          btn.style.background = '#b45309';
        }} else {{
          btn.dataset.mode = 'run';
          btn.innerText = 'Run matching';
          btn.style.background = '';
        }}
      }}
      function setUndoVisibility(data) {{
        const btn = document.getElementById('undoMatchBtn');
        if (!btn) return;
        const hasCompletedRun = !!data.latest_run && String(data.latest_run.status || '') !== 'running';
        btn.style.display = hasCompletedRun ? '' : 'none';
      }}
      function escHtml(value) {{
        return String(value ?? '')
          .replaceAll('&', '&amp;')
          .replaceAll('<', '&lt;')
          .replaceAll('>', '&gt;');
      }}
      let fullLogOpen = false;
      let autoFollowProgress = true;
      let lastProgressSignature = '';
      function refreshAutoFollowBtn() {{
        const btn = document.getElementById('autoFollowBtn');
        if (!btn) return;
        btn.innerText = autoFollowProgress ? 'Auto-follow: On' : 'Auto-follow: Off';
      }}
      function toggleAutoFollow(btn) {{
        pulseButton(btn);
        autoFollowProgress = !autoFollowProgress;
        refreshAutoFollowBtn();
      }}
      function bindFullLogState(rootEl) {{
        if (!rootEl) return;
        const details = rootEl.querySelector('.full-log-details');
        if (!details) return;
        details.open = !!fullLogOpen;
        details.addEventListener('toggle', () => {{
          fullLogOpen = !!details.open;
        }});
      }}
      function compactSolverRows(logs) {{
        const rows = [];
        let lastObjNum = null;
        function parseNum(raw) {{
          const v = Number(raw);
          return Number.isFinite(v) ? v : null;
        }}
        function formatRounded(n) {{
          return n === null ? '' : String(Math.round(n));
        }}
        function parseBoundFromSolverLine(msg) {{
          const patterns = [
            /bound\\s*[:=]\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i,
            /best_bound\\s*[:=]\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i,
            /next\\s*[:=]\\s*\\[\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i,
            /next\\s*[:=]\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i,
          ];
          for (const p of patterns) {{
            const m = msg.match(p);
            if (!m) continue;
            const n = parseNum(m[1]);
            if (n !== null) return n;
          }}
          return null;
        }}
        function parseNextRange(msg) {{
          const m = msg.match(/next\\s*[:=]\\s*\\[\\s*([-]?[0-9]+(?:\\.[0-9]+)?)\\s*,\\s*([-]?[0-9]+(?:\\.[0-9]+)?)\\s*\\]/i);
          if (!m) return null;
          const lo = parseNum(m[1]);
          const hi = parseNum(m[2]);
          if (lo === null || hi === null) return null;
          return {{ low: lo, high: hi }};
        }}
        (logs || []).forEach(line => {{
          const msg = String(line || '').trim();
          if (!msg) return;
          const low = msg.toLowerCase();

          if (low.startsWith('next solution found')) {{
            const tMatch = msg.match(/t\\s*=\\s*([0-9]+(?:\\.[0-9]+)?)\\s*s/i);
            const utilMatch = msg.match(/util\\s*=\\s*([0-9]+(?:\\.[0-9]+)?)/i);
            const penMatch = msg.match(/pen\\s*=\\s*([-]?[0-9]+)/i);
            const objMatch = msg.match(/obj\\s*=\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i);
            const boundMatch = msg.match(/bound\\s*=\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i);
            const gapMatch = msg.match(/gap\\s*=\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i);
            if (!tMatch) return;
            const objNum = objMatch ? parseNum(objMatch[1]) : null;
            if (objNum !== null) lastObjNum = objNum;
            const boundNum = boundMatch ? parseNum(boundMatch[1]) : null;
            const parsedGap = gapMatch ? parseNum(gapMatch[1]) : null;
            const impliedGap = parsedGap !== null ? parsedGap : ((boundNum !== null && objNum !== null) ? Math.abs(boundNum - objNum) : null);
            rows.push({{
              timeSec: Math.round(Number(tMatch[1])),
              util: utilMatch ? String(Math.round(Number(utilMatch[1]))) : '',
              pen: penMatch ? String(penMatch[1]) : '',
              obj: formatRounded(objNum),
              bound: formatRounded(boundNum),
              gap: formatRounded(impliedGap),
            }});
            return;
          }}

          if (low.startsWith('solver:') && low.includes('#bound')) {{
            const tMatch = msg.match(/([0-9]+(?:\\.[0-9]+)?)\\s*s/i);
            const nextRange = parseNextRange(msg);
            const boundNum = nextRange ? nextRange.high : parseBoundFromSolverLine(msg);
            const gapMatch = msg.match(/gap\\s*[:=]\\s*([-]?[0-9]+(?:\\.[0-9]+)?)/i);
            if (!tMatch) return;
            const parsedGap = gapMatch ? parseNum(gapMatch[1]) : null;
            const impliedGap = nextRange
              ? Math.abs(nextRange.high - nextRange.low)
              : (parsedGap !== null ? parsedGap : ((boundNum !== null && lastObjNum !== null) ? Math.abs(boundNum - lastObjNum) : null));
            if (boundNum === null && impliedGap === null) return;
            rows.push({{
              timeSec: Math.round(Number(tMatch[1])),
              util: '',
              pen: '',
              obj: '',
              bound: formatRounded(boundNum),
              gap: formatRounded(impliedGap),
            }});
          }}
        }});
        return rows;
      }}
      function renderCompactSolverTable(logs) {{
        const rows = compactSolverRows(logs);
        let completionLine = '';
        const all = Array.isArray(logs) ? logs : [];
        for (let i = all.length - 1; i >= 0; i--) {{
          const msg = String(all[i] || '').trim();
          if (!msg) continue;
          const low = msg.toLowerCase();
          const doneMatch = msg.match(/run completed in\\s*([0-9]+(?:\\.[0-9]+)?)\\s*s/i);
          if (doneMatch) {{
            const sec = Math.round(Number(doneMatch[1]));
            completionLine = `Run completed (${{sec}}s).`;
            break;
          }}
          const statusMatch = msg.match(/solve ended with status\\s*=\\s*([a-z_]+)/i);
          if (statusMatch) {{
            completionLine = `Run ended with status ${{statusMatch[1].toUpperCase()}}.`;
            break;
          }}
        }}
        if (!rows.length) {{
          return completionLine || 'No solver table rows yet.';
        }}
        const fmtTime = (sec) => String(Math.max(0, Number(sec) || 0)).padStart(7, ' ') + 's';
        const widths = {{ time: 10, util: 6, pen: 5, obj: 8, bound: 8, gap: 8 }};
        const pad = (s, w) => String(s ?? '').padStart(w, ' ');
        const sep = '  ';
        const header =
          pad('Time', widths.time) + sep +
          pad('Util', widths.util) + sep +
          pad('Pen', widths.pen) + sep +
          pad('Obj', widths.obj) + sep +
          pad('Bound', widths.bound) + sep +
          pad('Gap', widths.gap);
        const line = '-'.repeat(header.length);
        const out = [];
        out.push(header);
        out.push(line);
        rows.forEach(r => {{
          out.push(
            pad(fmtTime(r.timeSec), widths.time) + sep +
            pad(r.util, widths.util) + sep +
            pad(r.pen, widths.pen) + sep +
            pad(r.obj, widths.obj) + sep +
            pad(r.bound, widths.bound) + sep +
            pad(r.gap, widths.gap)
          );
        }});
        if (completionLine) {{
          out.push('');
          out.push(completionLine);
        }}
        return out.join('\\n');
      }}
      function renderLogPanels(logs) {{
        const compactText = escHtml(renderCompactSolverTable(logs));
        const fullText = escHtml((logs || []).join('\\n'));
        return (
          `<pre style="background:#0b1220;color:#dbeafe;padding:8px;white-space:pre-wrap;word-break:break-word;overflow:visible;">${{compactText}}</pre>` +
          `<details class="full-log-details"${{fullLogOpen ? ' open' : ''}}><summary>Full log</summary>` +
          `<pre style="background:#0b1220;color:#dbeafe;padding:8px;white-space:pre-wrap;word-break:break-word;overflow:visible;">${{fullText}}</pre>` +
          `</details>`
        );
      }}
      function renderResults(data) {{
        const results = document.getElementById('results');
        if (!data.latest_run || data.latest_run.status === 'running') {{
          results.innerHTML = '';
          return false;
        }}
        const run = data.latest_run;
        let html = `<p>Run #${{run.id}} status: <strong>${{run.status}}</strong>; utility=${{run.utility ?? 'n/a'}}, penalty=${{run.penalty ?? 'n/a'}}, overlap violations=${{run.overlap_count ?? 'n/a'}}</p>`;

        if (data.assignments.length) {{
          const avgMain = (data.assignments.reduce((acc,x)=>acc+x.main_score,0)/data.assignments.length).toFixed(2);
          const avgShadow = (data.assignments.reduce((acc,x)=>acc+x.shadow_score,0)/data.assignments.length).toFixed(2);
          html += `<p><strong>Average scores:</strong> main=${{avgMain}}, shadow=${{avgShadow}}</p>`;
        }}

        if (data.progress_logs.length) {{
          html += '<h3>Solver output</h3>';
          html += renderLogPanels(data.progress_logs);
        }}
        results.innerHTML = html;
        bindFullLogState(results);
        return true;
      }}

      let pollTimer = null;
      function scheduleNextPoll(ms) {{
        if (pollTimer) clearTimeout(pollTimer);
        pollTimer = setTimeout(() => {{ poll(); }}, ms);
      }}
      async function poll() {{
        let nextMs = 1200;
        try {{
          const r = await fetch(appUrl('/api/admin/status'));
          const data = await r.json();
          document.getElementById('status').innerText = data.running ? 'Matching in progress...' : '';
          setRunControlMode(!!data.running);
          setUndoVisibility(data);
          setFinalizedControl(!!data.finalized);
          setPackageWarning(data.missing_packages || [], data.latest_error || '');
          let classes = Array.isArray(data.classes) ? data.classes : [];
          let classId = data.class_id;
          if (!classes.length) {{
            try {{
              const clsResp = await fetch(appUrl('/api/admin/classes'));
              if (clsResp.ok) {{
                const clsData = await clsResp.json();
                if (Array.isArray(clsData.classes) && clsData.classes.length) {{
                  classes = clsData.classes;
                  if (classId === undefined || classId === null) classId = clsData.current_class_id;
                }}
              }}
            }} catch (_err) {{
              // ignore fallback failure; keep existing select contents
            }}
          }}
          populateClassSelect(classes, classId);
          const classNameInput = document.getElementById('classNameInput');
          if (window.__setClassNameSavedValue) {{
            window.__setClassNameSavedValue(data.class_name || '');
          }} else if (classNameInput && document.activeElement !== classNameInput) {{
            classNameInput.value = data.class_name || '';
          }}
          const progressSection = document.getElementById('progressSection');
          const progress = document.getElementById('progress');
          if (data.running) {{
            nextMs = 250;
            progressSection.style.display = 'block';
            const logs = data.progress_logs.length ? data.progress_logs : ['Starting run...'];
            const signature = String(logs.length) + '|' + String(logs[logs.length - 1] || '');
            if (signature !== lastProgressSignature) {{
              progress.innerHTML = renderLogPanels(logs);
              bindFullLogState(progress);
              lastProgressSignature = signature;
              if (autoFollowProgress) {{
                window.scrollTo({{ top: document.body.scrollHeight, behavior: 'smooth' }});
              }}
            }} else {{
              bindFullLogState(progress);
            }}
          }} else {{
            nextMs = 1200;
            progressSection.style.display = 'none';
            progress.innerHTML = '';
            lastProgressSignature = '';
          }}
          const hasResults = renderResults(data);
          document.getElementById('resultsSection').style.display = hasResults ? 'block' : 'none';
        }} catch (err) {{
          nextMs = 1200;
          document.getElementById('status').innerText = 'Server unavailable';
        }} finally {{
          scheduleNextPoll(nextMs);
        }}
      }}
      window.addEventListener('pageshow', () => {{ poll(); }});
      document.addEventListener('visibilitychange', () => {{
        if (!document.hidden) poll();
      }});
      wireClassNameAutosave();
      initThemeMode();
      refreshAutoFollowBtn();
      poll();
    </script>
    </body></html>
    """


def save_preferences(sid: int, scores: list[int]) -> tuple[bool, str]:
    conn = db_conn()
    class_id = get_current_class_id(conn)
    n = int(get_class_meta(conn, class_id, "n", "8"))
    if sid < 0 or sid >= n:
        conn.close()
        return False, "Invalid student id."
    finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
    if finalized:
        conn.close()
        return False, "Matching finalized; edits are locked."
    if len(scores) != n:
        conn.close()
        return False, "Invalid score vector length."
    veto_max = n // 4
    if sum(1 for s in scores if s == 0) > veto_max:
        conn.close()
        return False, f"Too many vetoes. Max allowed is {veto_max}."
    if scores[sid] < 4:
        conn.close()
        return False, "Your own topic must have score at least 4."
    if any((not isinstance(s, int)) or s < 0 or s > 5 for s in scores):
        conn.close()
        return False, "Scores must be integers in [0,5]."

    now = datetime.now(timezone.utc).isoformat()
    with conn:
        for j, score in enumerate(scores):
            conn.execute(
                "UPDATE preferences SET score=?, updated_at=? WHERE class_id=? AND student_id=? AND topic_id=?",
                (score, now, class_id, sid, j),
            )
    conn.close()
    return True, "Preferences saved."


def save_topic_title(sid: int, title: str) -> tuple[bool, str]:
    conn = db_conn()
    class_id = get_current_class_id(conn)
    n = int(get_class_meta(conn, class_id, "n", "8"))
    if sid < 0 or sid >= n:
        conn.close()
        return False, "Invalid student id."
    finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
    if finalized:
        conn.close()
        return False, "Matching finalized; edits are locked."
    clean = title.strip()
    if not clean:
        conn.close()
        return False, "Topic title cannot be empty."
    with conn:
        conn.execute("UPDATE students SET topic_title=? WHERE class_id=? AND id=?", (clean[:200], class_id, sid))
    conn.close()
    return True, "Topic title saved."


def save_student_name(sid: int, name: str) -> tuple[bool, str]:
    conn = db_conn()
    class_id = get_current_class_id(conn)
    n = int(get_class_meta(conn, class_id, "n", "8"))
    if sid < 0 or sid >= n:
        conn.close()
        return False, "Invalid student id."
    finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
    if finalized:
        conn.close()
        return False, "Matching finalized; edits are locked."
    clean = name.strip()
    if not clean:
        conn.close()
        return False, "Student name cannot be empty."
    with conn:
        conn.execute("UPDATE students SET name=? WHERE class_id=? AND id=?", (clean[:200], class_id, sid))
    conn.close()
    return True, "Student name saved."


def collect_problem(conn: sqlite3.Connection, class_id: int) -> tuple[list[str], list[list[int]]]:
    students = get_students(conn, class_id)
    n = len(students)
    topics = [s["topic_title"] for s in students]
    matrix = []
    for i in range(n):
        matrix.append(get_pref_row(conn, class_id, i, n))
    return topics, matrix


def ensure_local_venv_packages_on_path() -> None:
    root = Path(__file__).resolve().parent
    site_pkgs = root / ".venv" / "Lib" / "site-packages"
    if site_pkgs.exists():
        p = str(site_pkgs)
        if p not in sys.path:
            sys.path.insert(0, p)


def run_matching_background(run_id: int, class_id: int, stop_event: threading.Event) -> None:
    conn = db_conn()
    started_at_perf = time.perf_counter()

    def log(msg: str) -> None:
        with conn:
            idx = conn.execute("SELECT COALESCE(MAX(idx), -1)+1 FROM progress_logs WHERE run_id=?", (run_id,)).fetchone()[0]
            conn.execute("INSERT INTO progress_logs(run_id, idx, message) VALUES (?, ?, ?)", (run_id, idx, msg))

    try:
        try:
            import ortools  # noqa: F401
        except Exception:
            ensure_local_venv_packages_on_path()
        topics, matrix = collect_problem(conn, class_id)
        from match import solve_with_preferences_live
        result = solve_with_preferences_live(
            topic_titles=topics,
            preferences=matrix,
            M=4,
            time_limit_s=120,
            progress_cb=log,
            stop_event=stop_event,
        )

        with conn:
            conn.execute(
                "UPDATE match_runs SET finished_at=?, status=?, utility=?, penalty=?, overlap_count=? WHERE id=?",
                (
                    datetime.now(timezone.utc).isoformat(),
                    result["status"].lower(),
                    result.get("utility"),
                    result.get("penalty"),
                    len(result.get("overlaps", [])),
                    run_id,
                ),
            )
            conn.execute("DELETE FROM selected_topics WHERE run_id=?", (run_id,))
            conn.execute("DELETE FROM assignments WHERE run_id=?", (run_id,))
            conn.execute("DELETE FROM overlaps WHERE run_id=?", (run_id,))
            for t in result["selected_topics"]:
                conn.execute(
                    "INSERT INTO selected_topics(run_id, topic_id, title, partition) VALUES (?, ?, ?, ?)",
                    (run_id, t["id"], t["title"], t["partition"]),
                )
            for a in result["assignments"]:
                conn.execute(
                    "INSERT INTO assignments(run_id, student_id, main_topic, main_title, main_score, shadow_topic, shadow_title, shadow_score) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        run_id,
                        a["student"],
                        a["main_topic"],
                        a["main_title"],
                        a["main_score"],
                        a["shadow_topic"],
                        a["shadow_title"],
                        a["shadow_score"],
                    ),
                )
            for o in result.get("overlaps", []):
                conn.execute("INSERT INTO overlaps(run_id, s1, s2) VALUES (?, ?, ?)", (run_id, o[0], o[1]))
        elapsed_s = time.perf_counter() - started_at_perf
        log(f"run completed in {elapsed_s:.2f}s")
    except Exception as exc:  # noqa: BLE001
        with conn:
            conn.execute(
                "UPDATE match_runs SET finished_at=?, status=? WHERE id=?",
                (datetime.now(timezone.utc).isoformat(), f"error: {exc}", run_id),
            )
            idx = conn.execute("SELECT COALESCE(MAX(idx), -1)+1 FROM progress_logs WHERE run_id=?", (run_id,)).fetchone()[0]
            conn.execute("INSERT INTO progress_logs(run_id, idx, message) VALUES (?, ?, ?)", (run_id, idx, f"ERROR: {exc}"))
    finally:
        conn.close()
        with run_state_lock:
            run_state["thread"] = None
            run_state["stop_event"] = None
            run_state["current_run_id"] = None
            run_state["class_id"] = None


def request_server_shutdown() -> bool:
    with server_state_lock:
        server = server_instance
    if server is None:
        return False

    # Stop any active solver run before shutting down the web server loop.
    with run_state_lock:
        run_thread = run_state.get("thread")
        stop_event = run_state.get("stop_event")
        if stop_event is not None:
            stop_event.set()
    shutdown_requested.set()
    return True


def application(environ, start_response):
    ensure_db_initialized()
    raw_path = environ.get("PATH_INFO", "/") or "/"
    path = urlparse(raw_path).path or "/"
    method = environ.get("REQUEST_METHOD", "GET")

    if method == "GET" and path == "/":
        return html_response(start_response, render_home())

    if method == "GET" and path == "/student":
        qs = parse_qs(environ.get("QUERY_STRING", ""))
        sid = int(qs.get("sid", ["0"])[0])
        return html_response(start_response, render_student(sid))

    if method == "GET" and path == "/admin":
        return html_response(start_response, render_admin())

    if method == "GET" and path == "/launch":
        return html_response(
            start_response,
            f"""
            <html><head><title>Launching app...</title><style>
            body{font-family:Arial,sans-serif;background:#f3f6fb;margin:0;padding:24px;}
            .box{max-width:780px;margin:40px auto;background:#fff;border-radius:10px;padding:24px;box-shadow:0 2px 9px rgba(0,0,0,.08);}
            h1{margin-top:0;}
            .muted{color:#64748b;}
            </style></head>
            <body>
              <div class='box'>
                <h1>Launching app...</h1>
                <p id='msg'>Opening the app in a new tab.</p>
                <p class='muted' id='hint' style='display:none;'>If a pop-up was blocked, click this link: <a href='{app_url("/")}' target='_blank'>Open app</a></p>
              </div>
              <script>
                (function () {{
                  var opened = null;
                  try {{
                    opened = window.open({json.dumps(app_url('/?script_opened=1'))}, '_blank');
                  }} catch (e) {{
                    opened = null;
                  }}
                  if (opened) {{
                    try {{ opened.focus(); }} catch (e) {{}}
                    document.getElementById('msg').textContent = 'App opened. This launcher will close.';
                    setTimeout(function () {{ try {{ window.close(); }} catch (e) {{}} }}, 600);
                  }} else {{
                    document.getElementById('msg').textContent = 'Your browser blocked automatic tab opening.';
                    document.getElementById('hint').style.display = 'block';
                  }}
                }})();
              </script>
            </body></html>
            """,
        )

    stopped_page_html = """
            <html><head><title>Server stopped</title><style>
            body{font-family:Arial,sans-serif;background:#f3f6fb;margin:0;padding:24px;}
            .box{max-width:780px;margin:40px auto;background:#fff;border-radius:10px;padding:24px;box-shadow:0 2px 9px rgba(0,0,0,.08);}
            h1{margin-top:0;}
            .muted{color:#64748b;}
            </style></head>
            <body>
              <div class='box'>
                <h1>Server has stopped.</h1>
                <p>This app is no longer running. Restart the server to use it again.</p>
                <p class='muted'>This tab will try to close in <span id='sec'>5</span> seconds.</p>
                <p id='fallback' class='muted' style='display:none;'>If this tab does not close automatically, you can close it manually.</p>
              </div>
              <script>
                (function () {
                  var remaining = 5;
                  var secEl = document.getElementById('sec');
                  var fallback = document.getElementById('fallback');
                  var t = setInterval(function () {
                    remaining -= 1;
                    if (secEl) secEl.textContent = String(Math.max(remaining, 0));
                    if (remaining <= 0) {
                      clearInterval(t);
                      try { window.close(); } catch (e) {}
                      setTimeout(function () {
                        if (fallback) fallback.style.display = 'block';
                      }, 400);
                    }
                  }, 1000);
                })();
              </script>
            </body></html>
            """

    if method == "GET" and path == "/api/students_meta":
        conn = db_conn()
        class_id = get_current_class_id(conn)
        students = get_students(conn, class_id)
        finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
        conn.close()
        return json_response(start_response, {"students": students, "finalized": finalized})

    if method == "GET" and path == "/api/finalized":
        conn = db_conn()
        class_id = get_current_class_id(conn)
        finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
        conn.close()
        return json_response(start_response, {"finalized": finalized})

    if method == "GET" and path == "/api/home_fingerprint":
        conn = db_conn()
        fingerprint = compute_home_fingerprint(conn)
        conn.close()
        return json_response(start_response, {"fingerprint": fingerprint})

    if method == "GET" and path == "/stopped":
        return html_response(start_response, stopped_page_html)

    if method == "GET" and path == "/shutdown":
        ok = request_server_shutdown()
        if not ok:
            return html_response(start_response, "<html><body><h1>Server is not running.</h1></body></html>", "503 Service Unavailable")
        return html_response(start_response, stopped_page_html)

    if method == "POST" and path.startswith("/api/student/") and path.endswith("/preferences"):
        sid = int(path.split("/")[3])
        data = read_json(environ)
        ok, msg = save_preferences(sid, data.get("scores", []))
        return json_response(start_response, {"message": msg} if ok else {"error": msg}, "200 OK" if ok else "400 Bad Request")

    if method == "POST" and path.startswith("/api/student/") and path.endswith("/topic_title"):
        sid = int(path.split("/")[3])
        data = read_json(environ)
        ok, msg = save_topic_title(sid, str(data.get("title", "")))
        return json_response(start_response, {"message": msg} if ok else {"error": msg}, "200 OK" if ok else "400 Bad Request")

    if method == "POST" and path.startswith("/api/student/") and path.endswith("/name"):
        sid = int(path.split("/")[3])
        data = read_json(environ)
        ok, msg = save_student_name(sid, str(data.get("name", "")))
        return json_response(start_response, {"message": msg} if ok else {"error": msg}, "200 OK" if ok else "400 Bad Request")

    if method == "POST" and path == "/api/admin/set_n":
        data = read_json(environ)
        n = int(data.get("n", 8))
        n = max(4, n)
        confirm_delete = parse_bool(data.get("confirm_delete"))
        conn = db_conn()
        class_id = get_current_class_id(conn)
        prev_n = int(get_class_meta(conn, class_id, "n", "8"))
        doomed_students = []
        if n < prev_n:
            doomed_students = [
                {"id": int(r["id"]), "name": r["name"]}
                for r in conn.execute(
                    "SELECT id, name FROM students WHERE class_id=? AND id >= ? ORDER BY id",
                    (class_id, n),
                ).fetchall()
            ]
            if not confirm_delete:
                conn.close()
                names = ", ".join(f"S{d['id'] + 1} ({d['name']})" for d in doomed_students) or "none"
                return json_response(
                    start_response,
                    {
                        "error": f"Reducing n from {prev_n} to {n} will delete: {names}. Confirm deletion to proceed.",
                        "students_to_delete": doomed_students,
                    },
                    "409 Conflict",
                )
        with conn:
            if n < prev_n:
                conn.execute(
                    "DELETE FROM preferences WHERE class_id=? AND (student_id >= ? OR topic_id >= ?)",
                    (class_id, n, n),
                )
                conn.execute("DELETE FROM students WHERE class_id=? AND id >= ?", (class_id, n))
            set_class_meta(conn, class_id, "n", str(n))
        conn.close()
        ensure_students_and_preferences(class_id)
        if n == prev_n:
            msg = f"Matrix size unchanged (n={n}). Existing preferences were kept."
        elif n > prev_n:
            msg = f"Matrix expanded from n={prev_n} to n={n}. Existing preferences were kept; new rows/columns use defaults."
        else:
            if doomed_students:
                removed = ", ".join(f"S{d['id'] + 1} ({d['name']})" for d in doomed_students)
                msg = f"Matrix reduced from n={prev_n} to n={n}. Deleted students: {removed}."
            else:
                msg = f"Matrix reduced from n={prev_n} to n={n}."
        return json_response(start_response, {"message": msg})

    if method == "POST" and path == "/api/admin/add_students":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot add students while a run is in progress."}, "409 Conflict")
        data = read_json(environ)
        count = int(data.get("count", 0))
        conn = db_conn()
        class_id = get_current_class_id(conn)
        ok, msg = add_students_to_class(conn, class_id, count)
        conn.close()
        return json_response(start_response, {"message": msg} if ok else {"error": msg}, "200 OK" if ok else "400 Bad Request")

    if method == "POST" and path == "/api/admin/remove_students":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot remove students while a run is in progress."}, "409 Conflict")
        data = read_json(environ)
        raw_ids = data.get("student_ids", [])
        if not isinstance(raw_ids, list):
            return json_response(start_response, {"error": "student_ids must be a list."}, "400 Bad Request")
        conn = db_conn()
        class_id = get_current_class_id(conn)
        ok, msg = remove_students_from_class(conn, class_id, raw_ids)
        conn.close()
        return json_response(start_response, {"message": msg} if ok else {"error": msg}, "200 OK" if ok else "400 Bad Request")

    if method == "POST" and path == "/api/admin/reset":
        conn = db_conn()
        class_id = get_current_class_id(conn)
        with conn:
            conn.execute("DELETE FROM students WHERE class_id=?", (class_id,))
            conn.execute("DELETE FROM preferences WHERE class_id=?", (class_id,))
            clear_class_matching_data(conn, class_id)
            set_class_meta(conn, class_id, "n", "8")
            set_class_meta(conn, class_id, "finalized", "0")
        conn.close()
        ensure_students_and_preferences(class_id)
        return json_response(start_response, {"message": "Current class reset."})

    if method == "POST" and path == "/api/admin/finalize":
        data = read_json(environ)
        finalized = parse_bool(data.get("finalized"))
        conn = db_conn()
        class_id = get_current_class_id(conn)
        with conn:
            set_class_meta(conn, class_id, "finalized", "1" if finalized else "0")
        conn.close()
        return json_response(start_response, {"message": f"Student editing is now {'locked' if finalized else 'unlocked'}."})

    if method == "POST" and path == "/api/admin/run":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Run already in progress."}, "409 Conflict")
            conn = db_conn()
            class_id = get_current_class_id(conn)
            finalized = 1 if get_class_meta(conn, class_id, "finalized", "0") == "1" else 0
            with conn:
                cur = conn.execute(
                    "INSERT INTO match_runs(class_id, started_at, status, finalized_snapshot) VALUES (?, ?, ?, ?)",
                    (class_id, datetime.now(timezone.utc).isoformat(), "running", finalized),
                )
                run_id = cur.lastrowid
                conn.execute("INSERT INTO progress_logs(run_id, idx, message) VALUES (?, ?, ?)", (run_id, 0, "Run started."))
            conn.close()

            stop_event = threading.Event()
            t = threading.Thread(target=run_matching_background, args=(run_id, class_id, stop_event), daemon=True)
            run_state["thread"] = t
            run_state["stop_event"] = stop_event
            run_state["current_run_id"] = run_id
            run_state["class_id"] = class_id
            t.start()
        return json_response(start_response, {"message": "Run started."})

    if method == "POST" and path == "/api/admin/undo_matching":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot undo while a run is in progress."}, "409 Conflict")
        conn = db_conn()
        class_id = get_current_class_id(conn)
        with conn:
            clear_class_matching_data(conn, class_id)
        conn.close()
        return json_response(start_response, {"message": "Matching history/results cleared."})

    if method == "GET" and path == "/api/admin/classes":
        conn = db_conn()
        current_class_id = get_current_class_id(conn)
        rows = conn.execute("SELECT id, name FROM classes ORDER BY id").fetchall()
        classes = [{"id": int(r["id"]), "name": r["name"]} for r in rows]
        conn.close()
        return json_response(start_response, {"current_class_id": current_class_id, "classes": classes})

    if method == "POST" and path == "/api/admin/select_class":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot switch class while a run is in progress."}, "409 Conflict")
        data = read_json(environ)
        class_id = int(data.get("class_id", 0))
        conn = db_conn()
        exists = conn.execute("SELECT 1 FROM classes WHERE id=?", (class_id,)).fetchone() is not None
        if not exists:
            conn.close()
            return json_response(start_response, {"error": "Class not found."}, "404 Not Found")
        with conn:
            set_current_class_id(conn, class_id)
        ensure_students_and_preferences(class_id)
        conn.close()
        return json_response(start_response, {"message": "Current class switched."})

    if method == "POST" and path == "/api/admin/create_class":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot create/switch class while a run is in progress."}, "409 Conflict")
        data = read_json(environ)
        name = str(data.get("name", "")).strip()
        n = int(data.get("n", 8))
        conn = db_conn()
        with conn:
            class_id = create_class(conn, name, n)
            set_current_class_id(conn, class_id)
        conn.close()
        ensure_students_and_preferences(class_id)
        return json_response(start_response, {"message": "Class created.", "class_id": class_id})

    if method == "POST" and path == "/api/admin/delete_class":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot delete class while a run is in progress."}, "409 Conflict")
        data = read_json(environ)
        requested_class_id = int(data.get("class_id", 0))
        conn = db_conn()
        current_class_id = get_current_class_id(conn)
        class_id = requested_class_id if requested_class_id > 0 else current_class_id
        exists = conn.execute("SELECT 1 FROM classes WHERE id=?", (class_id,)).fetchone() is not None
        if not exists:
            conn.close()
            return json_response(start_response, {"error": "Class not found."}, "404 Not Found")
        class_count = int(conn.execute("SELECT COUNT(*) FROM classes").fetchone()[0])
        if class_count <= 1:
            conn.close()
            return json_response(start_response, {"error": "Cannot delete the only remaining class."}, "400 Bad Request")
        with conn:
            clear_class_matching_data(conn, class_id)
            conn.execute("DELETE FROM preferences WHERE class_id=?", (class_id,))
            conn.execute("DELETE FROM students WHERE class_id=?", (class_id,))
            conn.execute("DELETE FROM class_meta WHERE class_id=?", (class_id,))
            conn.execute("DELETE FROM classes WHERE id=?", (class_id,))
            next_row = conn.execute("SELECT id FROM classes ORDER BY id LIMIT 1").fetchone()
            next_class_id = int(next_row[0]) if next_row else 1
            set_current_class_id(conn, next_class_id)
        conn.close()
        ensure_students_and_preferences(next_class_id)
        return json_response(start_response, {"message": "Class deleted.", "class_id": next_class_id})

    if method == "POST" and path == "/api/admin/set_class_name":
        data = read_json(environ)
        new_name = str(data.get("name", "")).strip()
        if not new_name:
            return json_response(start_response, {"error": "Class name cannot be empty."}, "400 Bad Request")
        conn = db_conn()
        class_id = get_current_class_id(conn)
        with conn:
            conn.execute("UPDATE classes SET name=? WHERE id=?", (new_name[:200], class_id))
        conn.close()
        return json_response(start_response, {"message": "Class name updated."})

    if method == "POST" and path == "/api/admin/interrupt":
        with run_state_lock:
            stop_event = run_state.get("stop_event")
            if stop_event is None:
                return json_response(start_response, {"message": "No active run."})
            stop_event.set()
        return json_response(start_response, {"message": "Interrupt requested."})

    if method == "GET" and path == "/api/admin/export_csv":
        conn = db_conn()
        class_id = get_current_class_id(conn)
        class_row = conn.execute("SELECT name FROM classes WHERE id=?", (class_id,)).fetchone()
        class_name = (class_row[0] if class_row else f"class_{class_id}").strip() or f"class_{class_id}"
        csv_text = export_class_csv(conn, class_id)
        conn.close()
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", class_name).strip("_") or f"class_{class_id}"
        filename = f"{safe_name}_students_topics_preferences.csv"
        return text_response(
            start_response,
            csv_text,
            content_type="text/csv; charset=utf-8",
            headers=[("Content-Disposition", f'attachment; filename="{filename}"')],
        )

    if method == "POST" and path == "/api/admin/import_csv":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot import while a run is in progress."}, "409 Conflict")
        data = read_json(environ)
        csv_text = str(data.get("csv_text", ""))
        conn = db_conn()
        class_id = get_current_class_id(conn)
        ok, msg = import_class_csv(conn, class_id, csv_text)
        conn.close()
        return json_response(start_response, {"message": msg} if ok else {"error": msg}, "200 OK" if ok else "400 Bad Request")

    if method == "POST" and path == "/api/admin/randomize_preferences":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Cannot randomize while a run is in progress."}, "409 Conflict")
        data = read_json(environ)
        mode = str(data.get("mode", "category"))
        seed = int(time.time() * 1000) & 0x7FFFFFFF
        conn = db_conn()
        class_id = get_current_class_id(conn)
        ok, msg = randomize_class_data(conn, class_id, mode, seed)
        conn.close()
        return json_response(start_response, {"message": msg} if ok else {"error": msg}, "200 OK" if ok else "400 Bad Request")

    if method == "POST" and path == "/api/admin/stop":
        ok = request_server_shutdown()
        if not ok:
            return json_response(start_response, {"error": "Server is not running."}, "503 Service Unavailable")
        return json_response(start_response, {"message": "Server shutdown requested.", "redirect": app_url("/stopped")})

    if method == "GET" and path == "/api/admin/status":
        conn = db_conn()
        class_id = get_current_class_id(conn)
        class_row = conn.execute("SELECT id, name FROM classes WHERE id=?", (class_id,)).fetchone()
        n = int(get_class_meta(conn, class_id, "n", "8"))
        finalized = get_class_meta(conn, class_id, "finalized", "0") == "1"
        latest = conn.execute("SELECT * FROM match_runs WHERE class_id=? ORDER BY id DESC LIMIT 1", (class_id,)).fetchone()
        classes = [dict(r) for r in conn.execute("SELECT id, name FROM classes ORDER BY id").fetchall()]
        selected_topics = []
        assignments = []
        overlaps = []
        logs = []
        if latest:
            rid = int(latest["id"])
            logs = [r[0] for r in conn.execute("SELECT message FROM progress_logs WHERE run_id=? ORDER BY idx", (rid,)).fetchall()]
            if latest["status"] != "running":
                selected_topics = [dict(r) for r in conn.execute("SELECT topic_id, title, partition FROM selected_topics WHERE run_id=? ORDER BY topic_id", (rid,)).fetchall()]
                assignments = [dict(r) for r in conn.execute("SELECT * FROM assignments WHERE run_id=? ORDER BY student_id", (rid,)).fetchall()]
                overlaps = [dict(r) for r in conn.execute("SELECT s1, s2 FROM overlaps WHERE run_id=? ORDER BY s1, s2", (rid,)).fetchall()]
        latest_error = ""
        missing_packages: list[str] = []
        if latest and str(latest["status"]).startswith("error:"):
            latest_error = str(latest["status"])[len("error:"):].strip()
            missing_packages = extract_missing_packages(str(latest["status"]))
        if not missing_packages and logs:
            missing_packages = extract_missing_packages("\n".join(logs))
        conn.close()

        with run_state_lock:
            running = run_state["thread"] is not None and run_state.get("class_id") == class_id
        return json_response(
            start_response,
            {
                "class_id": class_id,
                "class_name": class_row["name"] if class_row else f"Class {class_id}",
                "classes": classes,
                "n": n,
                "finalized": finalized,
                "running": running,
                "progress_logs": logs,
                "latest_run": dict(latest) if latest else None,
                "selected_topics": selected_topics,
                "assignments": assignments,
                "overlaps": overlaps,
                "latest_error": latest_error,
                "missing_packages": missing_packages,
            },
        )

    return html_response(start_response, "Not found", "404 Not Found")


if __name__ == "__main__":
    ensure_db_initialized()
    shutdown_requested.clear()
    with make_server("0.0.0.0", 8000, application) as server:
        with server_state_lock:
            server_instance = server
        try:
            print("Serving on http://0.0.0.0:8000")
            server.timeout = 0.5
            while not shutdown_requested.is_set():
                server.handle_request()
        finally:
            with server_state_lock:
                server_instance = None

