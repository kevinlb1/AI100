from __future__ import annotations

import json
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


DB_PATH = "match_app.db"

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
            conn.execute(
                """
                INSERT INTO students(class_id, id, name, topic_title) VALUES (?, ?, ?, ?)
                ON CONFLICT(class_id, id) DO NOTHING
                """,
                (class_id, i, f"Student {i + 1}", f"Topic {i + 1}"),
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
            for r in conn.execute("SELECT topic_id FROM selected_topics WHERE run_id=? ORDER BY topic_id", (latest_group_run_id,)).fetchall()
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
    selection_note_html = ""
    if has_selected_columns and latest_group_run_id is not None:
        selection_note_html = f"<p class='muted'>Rows/columns are grouped from run #{latest_group_run_id}. Colored cells show active student-group pairs.</p>"

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
        topic_label_html = (
            f"<div class='topic-code'>T{tid + 1}</div>"
            if dense_mode
            else f"<div class='topic-code topic-code-full'>{escape(topic_title)}</div>"
        )
        topic_headers.append(
            f"<th class='{col_class}' title='{escape(topic_title)}'{header_style}>"
            f"{group_tag_html}"
            f"{topic_label_html}"
            f"<div class='topic-title'>{escape(topic_title)}</div>"
            f"</th>"
        )
    colgroup = "<col class='student-col'>" + "".join("<col class='topic-col'>" for _ in ordered_topic_ids)

    def row_sort_key(row: tuple[int, list[int]]) -> tuple[int, int, int]:
        sid = int(row[0])
        main_topic = main_topic_by_student.get(sid)
        if main_topic is None or main_topic not in group_rank:
            return (1, 10_000, sid)
        return (0, group_rank[main_topic], sid)

    def build_topic_chip(topic_id: int | None, is_shadow: bool, has_violation: bool = False) -> str:
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
        if is_shadow and has_violation:
            chip_class += " shadow-chip-violation"
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
        for tid in ordered_topic_ids:
            score = pref[tid]
            cell_class = f"score-cell score-{score}"
            cell_style = ""
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
            cells.append(f"<td class='{cell_class}'{cell_style}>{score}</td>")
        row_class = "matrix-row" if changed else "matrix-row matrix-row-dim"
        student_name = escape(names_by_id.get(sid, ""))
        main_topic = main_topic_by_student.get(sid)
        shadow_topic = shadow_topic_by_student.get(sid)
        has_overlap_violation = sid in overlap_violation_students
        main_chip = build_topic_chip(main_topic, is_shadow=False)
        shadow_chip = build_topic_chip(shadow_topic, is_shadow=True, has_violation=has_overlap_violation)
        group_badge = f"{main_chip}{shadow_chip}"
        matrix_rows.append(
            f"<tr class='{row_class}' onclick='window.location.href=\"/student?sid={sid}\"'>"
            f"<td class='student-cell' title='{student_name}'>"
            f"<span class='student-name'>{student_name}</span>"
            f"<span class='student-group'>{group_badge}</span>"
            f"</td>{''.join(cells)}</tr>"
        )
    return f"""
    <html><head><title>Topic Match</title><style>{base_css()}</style></head><body>
    <main class='container home-layout {'matrix-dense' if dense_mode else ''}'>
      <h1>Topic Matching Portal</h1>
      <p class='muted'>Class: <strong>{escape(class_name)}</strong></p>
      <p>Student editing of preferences: <strong id='lockStatus'>{'Locked' if finalized else 'Unlocked'}</strong></p>
      <div class='top-right-actions'>
        <a class='button-link' href='/admin'>Open Admin Dashboard</a>
        <a class='button-link button-danger' href='/shutdown'>Stop server</a>
      </div>
      {selection_note_html}
      <table class='matrix-table'>
        <colgroup>{colgroup}</colgroup>
        <thead>
          <tr><th class='student-head'>Student</th>{''.join(topic_headers)}</tr>
        </thead>
        <tbody>
          {"".join(matrix_rows)}
        </tbody>
      </table>
      {latest_html}
    </main>
    <script>
      const homeFingerprint = {json.dumps(home_fingerprint)};
      async function refreshLockStatus() {{
        try {{
          const r = await fetch('/api/finalized');
          if (!r.ok) return;
          const data = await r.json();
          document.getElementById('lockStatus').innerText = data.finalized ? 'Locked' : 'Unlocked';
        }} catch (err) {{
          // keep current text on transient errors
        }}
      }}
      async function refreshHomeIfChanged() {{
        try {{
          const r = await fetch('/api/home_fingerprint');
          if (!r.ok) return;
          const data = await r.json();
          if ((data.fingerprint || '') !== homeFingerprint) {{
            window.location.reload();
          }}
        }} catch (err) {{
          // ignore transient poll errors
        }}
      }}
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
    .row { display:flex; gap:14px; flex-wrap:wrap; }
    .col { flex:1; min-width:320px; }
    .inline-finalize { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-top:8px; }
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
    .matrix-table .student-group { margin-left:auto; display:inline-flex; align-items:center; gap:4px; }
    .matrix-table .topic-head { text-align:center; font-size:11px; padding:3px 1px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; color:var(--text); }
    .matrix-table .group-head { font-weight:700; }
    .matrix-table .topic-group { font-size:9px; line-height:1; font-weight:700; }
    .matrix-table .topic-code { font-weight:700; line-height:1.05; }
    .matrix-table .topic-code-full { font-size:10px; font-weight:600; color:var(--muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .matrix-table .topic-title { display:none; }
    .matrix-table .score-cell { text-align:center; font-weight:600; font-size:11px; padding:4px 1px; }
    .group-chip { display:inline-block; font-size:10px; font-weight:700; color:#0f766e; background:#ccfbf1; border:1px solid #99f6e4; border-radius:999px; padding:1px 6px; vertical-align:middle; position:relative; line-height:1.1; }
    .shadow-chip { font-size:7px; padding:1px 4px; }
    .shadow-chip-violation::after {
      content:'*';
      position:absolute;
      top:-6px;
      right:-2px;
      color:#ef4444;
      font-size:12px;
      font-weight:900;
      line-height:1;
    }
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
        "SELECT id FROM match_runs WHERE class_id=? AND status IN ('optimal','feasible','interrupted') ORDER BY id DESC LIMIT 1",
        (class_id,),
    ).fetchone()
    assignment = None
    if latest_id_row:
        assignment = conn.execute("SELECT * FROM assignments WHERE run_id=? AND student_id=?", (latest_id_row[0], sid)).fetchone()
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
      <p><a href='/'>← back</a></p>
      <h1>Student {sid + 1} preferences</h1>
      <p class='muted'>Drag topics between score buckets (0..5). Max vetoes (score 0): floor(n/4).</p>
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
      const sid = {sid};
      let editable = {editable};
      let topics = {topics_json};
      let scores = {pref_json};
      const vetoMax = Math.floor(topics.length / 4);
      const prefMsg = document.getElementById('msg');
      let prefSaveTimer = null;
      let prefSaveInFlight = false;
      let prefSaveQueued = false;
      let lastSavedScores = JSON.stringify(scores);

      function findTopicById(topicId) {{
        return topics.find(t => t.id === topicId);
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
          buckets.push(`<div class='bucket' data-score='${{s}}' ondragover='event.preventDefault()' ondrop='dropTopic(event, ${{s}})'><h3>Score ${{s}}</h3><div id='bucket-${{s}}' class='bucket-topics'></div></div>`);
        }}
        app.innerHTML = `<div class='buckets'>${{buckets.join('')}}</div><p>Veto count: <span id='vetoCount'></span> / ${{vetoMax}}</p>`;
        const orderedTopics = [...topics].sort((a, b) => {{
          if (scores[b.id] !== scores[a.id]) return scores[b.id] - scores[a.id];
          return a.id - b.id;
        }});
        orderedTopics.forEach(t => {{
          const div = document.createElement('div');
          div.className='topic';
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
        document.getElementById('vetoCount').innerText = String(v);
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
          const res = await fetch(`/api/student/${{sid}}/preferences`, {{
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
            const res = await fetch(endpoint, {{
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
          const res = await fetch('/api/students_meta');
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

      applyEditableState();
      render();
    </script>
    </body></html>
    """


def render_admin() -> str:
    return f"""
    <html><head><title>Admin</title><style>{base_css()}</style></head><body>
    <main class='container'>
      <p><a href='/'>← back</a></p>
      <h1>Admin dashboard</h1>
      <div class='top-right-actions'>
        <a class='button-link button-danger' href='/shutdown'>Stop server</a>
      </div>
      <p>
        <label>Theme:
          <select id='themeModeSelect' onchange='onThemeModeChange(this)'>
            <option value='system'>Follow system settings</option>
            <option value='light'>Light mode</option>
            <option value='dark'>Dark mode</option>
          </select>
        </label>
      </p>
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
        <h2>Configuration</h2>
        <p>
          <label>Class name:
            <input id='classNameInput' type='text' maxlength='200' style='width:min(420px, 100%);'>
          </label>
          <button id='setClassNameBtn' onclick='setClassName(this)'>Save class name</button>
        </p>
        <p>
          <label>Select current class:
            <select id='classSelect' onchange='onClassSelectChanged(this)'></select>
          </label>
        </p>
        <label>Number of students/topics: <input id='nInput' type='number' min='4' value='8'></label>
        <button onclick='resetDb(this)'>Reset database</button>
        <button onclick='setN(this)'>Apply matrix size</button>
      </section>
      <section id='progressSection' class='card' style='display:none;'>
        <h2>Solver progress</h2>
        <pre id='progress' style='background:#0b1220;color:#dbeafe;padding:8px;white-space:pre-wrap;word-break:break-word;overflow:visible;'></pre>
      </section>
      <section id='resultsSection' class='card' style='display:none;'>
        <h2>Matching results</h2>
        <div id='results'></div>
      </section>
    </main>
    <script>
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
        const sel = document.getElementById('themeModeSelect');
        if (sel && sel.value !== next) sel.value = next;
        try {{ localStorage.setItem(THEME_MODE_KEY, next); }} catch (err) {{}}
      }}
      function onThemeModeChange(sel) {{
        applyThemeMode(sel ? sel.value : 'system');
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
        const r = await fetch(path, {{method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify(body)}});
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
        const prev = select.value;
        select.innerHTML = '';
        (classes || []).forEach(c => {{
          const opt = document.createElement('option');
          opt.value = String(c.id);
          opt.textContent = c.name || `Class ${{c.id}}`;
          select.appendChild(opt);
        }});
        const createOpt = document.createElement('option');
        createOpt.value = '__create__';
        createOpt.textContent = 'Create new class';
        select.appendChild(createOpt);
        const target = String(currentClassId ?? '');
        if (target && Array.from(select.options).some(o => o.value === target)) {{
          select.value = target;
        }} else if (prev && Array.from(select.options).some(o => o.value === prev)) {{
          select.value = prev;
        }}
      }}
      async function onClassSelectChanged(selectEl) {{
        const value = String(selectEl.value || '');
        if (value === '__create__') {{
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
          return;
        }}
        const classId = Number(value);
        if (!Number.isFinite(classId) || classId <= 0) return;
        const data = await post('/api/admin/select_class', {{class_id: classId}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function setClassName(btn) {{
        pulseButton(btn);
        const value = (document.getElementById('classNameInput').value || '').trim();
        const data = await post('/api/admin/set_class_name', {{name: value}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        await poll();
      }}
      async function setN(btn) {{
        pulseButton(btn);
        const requested = Math.max(4, Number(document.getElementById('nInput').value) || 8);
        const current = Number(window.__currentN || 8);
        let confirmDelete = false;
        if (requested < current) {{
          try {{
            const r = await fetch('/api/students_meta');
            let doomed = [];
            if (r.ok) {{
              const meta = await r.json();
              doomed = (meta.students || [])
                .filter(s => Number(s.id) >= requested)
                .sort((a,b) => Number(a.id) - Number(b.id));
            }}
            const listText = doomed.length
              ? doomed.map(s => `S${{Number(s.id) + 1}}: ${{String(s.name || '').trim() || '(unnamed)'}}`).join('\\n')
              : '(no students found)';
            const ok = window.confirm(
              `Reduce from n=${{current}} to n=${{requested}}?\\n\\nThe following students will be deleted:\\n${{listText}}\\n\\nThis cannot be undone.`
            );
            if (!ok) {{
              document.getElementById('nInput').value = String(current);
              return;
            }}
            confirmDelete = true;
          }} catch (err) {{
            const ok = window.confirm(
              `Reduce from n=${{current}} to n=${{requested}}?\\n\\nLast students will be deleted. Continue?`
            );
            if (!ok) {{
              document.getElementById('nInput').value = String(current);
              return;
            }}
            confirmDelete = true;
          }}
        }}
        const data = await post('/api/admin/set_n', {{n: requested, confirm_delete: confirmDelete}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
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
          html += '<h3>Solver log</h3>';
          html += `<pre style="background:#0b1220;color:#dbeafe;padding:8px;white-space:pre-wrap;word-break:break-word;overflow:visible;">${{escHtml(data.progress_logs.join('\\n'))}}</pre>`;
        }}
        results.innerHTML = html;
        return true;
      }}

      async function poll() {{
        try {{
          const r = await fetch('/api/admin/status');
          const data = await r.json();
          document.getElementById('status').innerText = data.running ? 'Matching in progress...' : '';
          setRunControlMode(!!data.running);
          setUndoVisibility(data);
          setFinalizedControl(!!data.finalized);
          setPackageWarning(data.missing_packages || [], data.latest_error || '');
          populateClassSelect(data.classes || [], data.class_id);
          const classNameInput = document.getElementById('classNameInput');
          if (classNameInput && document.activeElement !== classNameInput) {{
            classNameInput.value = data.class_name || '';
          }}
          document.getElementById('nInput').value = data.n;
          window.__currentN = Number(data.n || window.__currentN || 8);
          const progressSection = document.getElementById('progressSection');
          const progress = document.getElementById('progress');
          if (data.running) {{
            progressSection.style.display = 'block';
            progress.innerText = data.progress_logs.length ? data.progress_logs.join('\\n') : 'Starting run...';
          }} else {{
            progressSection.style.display = 'none';
            progress.innerText = '';
          }}
          const hasResults = renderResults(data);
          document.getElementById('resultsSection').style.display = hasResults ? 'block' : 'none';
        }} catch (err) {{
          document.getElementById('status').innerText = 'Server unavailable';
        }}
      }}
      let pollHandle = setInterval(poll, 1200);
      window.addEventListener('pageshow', () => {{ poll(); }});
      document.addEventListener('visibilitychange', () => {{
        if (!document.hidden) poll();
      }});
      initThemeMode();
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
    path = urlparse(environ.get("PATH_INFO", "/")).path
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
            """
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
                <p class='muted' id='hint' style='display:none;'>If a pop-up was blocked, click this link: <a href='/' target='_blank'>Open app</a></p>
              </div>
              <script>
                (function () {
                  var opened = null;
                  try {
                    opened = window.open('/?script_opened=1', '_blank');
                  } catch (e) {
                    opened = null;
                  }
                  if (opened) {
                    try { opened.focus(); } catch (e) {}
                    document.getElementById('msg').textContent = 'App opened. This launcher will close.';
                    setTimeout(function () { try { window.close(); } catch (e) {} }, 600);
                  } else {
                    document.getElementById('msg').textContent = 'Your browser blocked automatic tab opening.';
                    document.getElementById('hint').style.display = 'block';
                  }
                })();
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

    if method == "POST" and path == "/api/admin/reset":
        conn = db_conn()
        class_id = get_current_class_id(conn)
        run_ids = [int(r[0]) for r in conn.execute("SELECT id FROM match_runs WHERE class_id=?", (class_id,)).fetchall()]
        with conn:
            conn.execute("DELETE FROM students WHERE class_id=?", (class_id,))
            conn.execute("DELETE FROM preferences WHERE class_id=?", (class_id,))
            for rid in run_ids:
                conn.execute("DELETE FROM progress_logs WHERE run_id=?", (rid,))
                conn.execute("DELETE FROM selected_topics WHERE run_id=?", (rid,))
                conn.execute("DELETE FROM assignments WHERE run_id=?", (rid,))
                conn.execute("DELETE FROM overlaps WHERE run_id=?", (rid,))
            conn.execute("DELETE FROM match_runs WHERE class_id=?", (class_id,))
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
        run_ids = [int(r[0]) for r in conn.execute("SELECT id FROM match_runs WHERE class_id=?", (class_id,)).fetchall()]
        with conn:
            for rid in run_ids:
                conn.execute("DELETE FROM progress_logs WHERE run_id=?", (rid,))
                conn.execute("DELETE FROM selected_topics WHERE run_id=?", (rid,))
                conn.execute("DELETE FROM assignments WHERE run_id=?", (rid,))
                conn.execute("DELETE FROM overlaps WHERE run_id=?", (rid,))
            conn.execute("DELETE FROM match_runs WHERE class_id=?", (class_id,))
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

    if method == "POST" and path == "/api/admin/stop":
        ok = request_server_shutdown()
        if not ok:
            return json_response(start_response, {"error": "Server is not running."}, "503 Service Unavailable")
        return json_response(start_response, {"message": "Server shutdown requested.", "redirect": "/stopped"})

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
