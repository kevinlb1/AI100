from __future__ import annotations

import json
import sqlite3
import threading
import time
from datetime import datetime
from html import escape
from urllib.parse import parse_qs, urlparse
from wsgiref.simple_server import make_server


DB_PATH = "match_app.db"
APP_VERSION = "v2-multiuser"

run_state_lock = threading.Lock()
run_state = {
    "thread": None,
    "stop_event": None,
    "current_run_id": None,
}
server_state_lock = threading.Lock()
server_instance = None


def db_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def get_meta(conn: sqlite3.Connection, key: str, default: str) -> str:
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute("INSERT INTO meta(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))


def initialize_db() -> None:
    conn = db_conn()
    with conn:
        conn.execute("CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT NOT NULL)")
        conn.execute("CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY, name TEXT NOT NULL, topic_title TEXT NOT NULL)")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS preferences(
                student_id INTEGER NOT NULL,
                topic_id INTEGER NOT NULL,
                score INTEGER NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY(student_id, topic_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS match_runs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
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

        if conn.execute("SELECT COUNT(*) FROM meta").fetchone()[0] == 0:
            set_meta(conn, "n", "8")
            set_meta(conn, "finalized", "0")
    ensure_students_and_preferences()
    conn.close()


def ensure_students_and_preferences() -> None:
    conn = db_conn()
    n = int(get_meta(conn, "n", "8"))
    now = datetime.utcnow().isoformat()
    with conn:
        for i in range(n):
            conn.execute(
                "INSERT INTO students(id, name, topic_title) VALUES (?, ?, ?) ON CONFLICT(id) DO NOTHING",
                (i, f"Student {i + 1}", f"Topic {i + 1}"),
            )
        conn.execute("DELETE FROM students WHERE id >= ?", (n,))

        for i in range(n):
            for j in range(n):
                default_score = 4 if i == j else 3
                conn.execute(
                    "INSERT INTO preferences(student_id, topic_id, score, updated_at) VALUES (?, ?, ?, ?) ON CONFLICT(student_id, topic_id) DO NOTHING",
                    (i, j, default_score, now),
                )
        conn.execute("DELETE FROM preferences WHERE student_id >= ? OR topic_id >= ?", (n, n))

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


def get_students(conn: sqlite3.Connection) -> list[dict]:
    return [dict(r) for r in conn.execute("SELECT id, name, topic_title FROM students ORDER BY id").fetchall()]


def get_pref_row(conn: sqlite3.Connection, sid: int, n: int) -> list[int]:
    rows = conn.execute("SELECT topic_id, score FROM preferences WHERE student_id=? ORDER BY topic_id", (sid,)).fetchall()
    if len(rows) != n:
        ensure_students_and_preferences()
        rows = conn.execute("SELECT topic_id, score FROM preferences WHERE student_id=? ORDER BY topic_id", (sid,)).fetchall()
    return [int(r[1]) for r in rows]


def student_has_non_default_preferences(pref_row: list[int], sid: int) -> bool:
    for tid, score in enumerate(pref_row):
        default_score = 4 if tid == sid else 3
        if score != default_score:
            return True
    return False


def render_home() -> str:
    conn = db_conn()
    students = get_students(conn)
    n = int(get_meta(conn, "n", "8"))
    finalized = get_meta(conn, "finalized", "0") == "1"
    latest = conn.execute("SELECT id, status FROM match_runs ORDER BY id DESC LIMIT 1").fetchone()
    pref_rows = [(s["id"], get_pref_row(conn, s["id"], n)) for s in students]
    conn.close()

    options = "".join(f"<option value='{s['id']}'>S{s['id'] + 1}: {escape(s['name'])}</option>" for s in students)
    latest_html = f"<p>Latest run: #{latest['id']} ({latest['status']})</p>" if latest else "<p>No runs yet.</p>"
    topic_headers = "".join(f"<th>T{s['id'] + 1}<br><span class='muted'>{escape(s['topic_title'])}</span></th>" for s in students)
    matrix_rows = []
    updated_students = []
    for sid, pref in pref_rows:
        changed = student_has_non_default_preferences(pref, sid)
        if changed:
            updated_students.append(f"S{sid + 1}")
        status_badge = "<span class='status-updated'>updated</span>" if changed else "<span class='status-default'>default</span>"
        cells = "".join(f"<td class='score-cell score-{score}'>{score}</td>" for score in pref)
        matrix_rows.append(f"<tr><td><strong>S{sid + 1}</strong> {status_badge}</td>{cells}</tr>")
    updated_html = ", ".join(updated_students) if updated_students else "None"
    return f"""
    <html><head><title>Topic Match ({APP_VERSION})</title><style>{base_css()}</style></head><body>
    <main class='container'>
      <h1>Topic Matching Portal ({APP_VERSION})</h1>
      <p>Finalized: <strong>{'Yes' if finalized else 'No'}</strong></p>
      {latest_html}
      <section class='card'>
        <h2>Overview: Topics & preferences</h2>
        <p>Students with non-default preferences: <strong>{updated_html}</strong></p>
        <p class='muted'>Default baseline: own topic = 4, all others = 3.</p>
        <table>
          <thead>
            <tr><th>Student</th>{topic_headers}</tr>
          </thead>
          <tbody>
            {"".join(matrix_rows)}
          </tbody>
        </table>
      </section>
      <section class='card'>
        <h2>Student</h2>
        <form action='/student' method='get'>
          <select name='sid'>{options}</select>
          <button type='submit'>Open Student View</button>
        </form>
      </section>
      <section class='card'>
        <h2>Administrator</h2>
        <a class='button-link' href='/admin'>Open Admin Dashboard</a>
      </section>
    </main>
    </body></html>
    """


def base_css() -> str:
    return """
    body { font-family: Arial, sans-serif; background:#f3f6fb; margin:0; }
    .container { max-width:1200px; margin:20px auto; background:white; border-radius:12px; padding:18px; box-shadow:0 2px 9px rgba(0,0,0,.08); }
    .card { border:1px solid #dbe4ef; border-radius:8px; padding:12px; margin:12px 0; }
    button, .button-link { background:#2563eb; color:white; border:none; border-radius:6px; padding:8px 12px; text-decoration:none; cursor:pointer; transition:transform .08s ease, filter .08s ease; }
    button.pressed { transform:translateY(1px) scale(.98); filter:brightness(.9); }
    .muted { color:#667; }
    .buckets { display:flex; flex-direction:column; gap:10px; }
    .bucket { background:#f8fbff; border:2px dashed #b7c8dd; min-height:90px; border-radius:8px; padding:8px; }
    .bucket h3 { margin:0 0 8px 0; font-size:14px; }
    .bucket-topics { display:flex; flex-wrap:wrap; gap:6px; }
    .topic { background:#dbeafe; border:1px solid #8ab5ff; border-radius:16px; padding:6px 10px; margin-bottom:6px; font-size:12px; cursor:grab; display:inline-flex; align-items:center; gap:5px; }
    .row { display:flex; gap:14px; flex-wrap:wrap; }
    .col { flex:1; min-width:320px; }
    table { border-collapse:collapse; width:100%; }
    th,td { border:1px solid #ccd7e4; padding:6px; text-align:left; }
    .score-cell { text-align:center; font-weight:600; }
    .score-0 { background:#fee2e2; }
    .score-1 { background:#ffedd5; }
    .score-2 { background:#fef9c3; }
    .score-3 { background:#e0f2fe; }
    .score-4 { background:#dcfce7; }
    .score-5 { background:#bbf7d0; }
    .status-updated { color:#1d4ed8; font-weight:700; }
    .status-default { color:#64748b; font-weight:700; }
    """


def render_student(sid: int) -> str:
    conn = db_conn()
    n = int(get_meta(conn, "n", "8"))
    students = get_students(conn)
    sid = max(0, min(sid, len(students) - 1))
    pref = get_pref_row(conn, sid, n)
    finalized = get_meta(conn, "finalized", "0") == "1"
    latest_id_row = conn.execute("SELECT id FROM match_runs WHERE status IN ('optimal','feasible','interrupted') ORDER BY id DESC LIMIT 1").fetchone()
    assignment = None
    if latest_id_row:
        assignment = conn.execute("SELECT * FROM assignments WHERE run_id=? AND student_id=?", (latest_id_row[0], sid)).fetchone()
    conn.close()

    topics_json = json.dumps([{"id": s["id"], "short": f"S{s['id'] + 1}", "title": s["topic_title"]} for s in students])
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
      <p>Editing enabled: <strong>{'Yes' if not finalized else 'No (finalized by admin)'}</strong></p>
      {assignment_html}
      <div class='card'>
        <h2>Your Topic Title</h2>
        <label>
          Topic title:
          <input id='topicTitleInput' type='text' value='{escape(students[sid]["topic_title"])}' maxlength='200' style='width:min(520px, 100%); margin-right:8px;'>
        </label>
        <button id='saveTitleBtn'>Save title</button>
        <span id='titleMsg' class='muted'></span>
      </div>
      <div id='app'></div>
      <button id='saveBtn'>Save preferences</button> <span id='msg' class='muted'></span>
    </main>
    <script>
      const sid = {sid};
      const editable = {editable};
      const topics = {topics_json};
      let scores = {pref_json};
      const vetoMax = Math.floor(topics.length / 4);

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
          div.title = t.title;
          div.innerText = `${{t.short}} (${{t.title}})`;
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
      }}

      document.getElementById('saveBtn').onclick = async () => {{
        const msg = document.getElementById('msg');
        const res = await fetch(`/api/student/${{sid}}/preferences`, {{
          method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{scores}})
        }});
        const data = await res.json();
        msg.innerText = data.message || data.error || 'Saved';
      }};
      document.getElementById('saveTitleBtn').onclick = async () => {{
        if (!editable) return;
        const title = document.getElementById('topicTitleInput').value.trim();
        const msg = document.getElementById('titleMsg');
        const res = await fetch(`/api/student/${{sid}}/topic_title`, {{
          method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{title}})
        }});
        const data = await res.json();
        msg.innerText = data.message || data.error || 'Saved';
      }};
      if (!editable) {{
        document.getElementById('saveTitleBtn').disabled = true;
        document.getElementById('topicTitleInput').disabled = true;
      }}

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
      <div class='row'>
        <div class='col card'>
          <h2>Configuration</h2>
          <label>Number of students/topics: <input id='nInput' type='number' min='4' value='8'></label>
          <button onclick='setN(this)'>Apply & reset matrix</button>
          <button onclick='resetDb(this)'>Reset database</button>
          <label><input id='finalToggle' type='checkbox' onchange='setFinalized()'> Finalize matching (lock student edits)</label>
        </div>
        <div class='col card'>
          <h2>Run matching</h2>
          <button onclick='runMatch(this)'>Run matching</button>
          <button onclick='interruptMatch(this)'>Interrupt run</button>
          <button onclick='stopServer(this)' style='background:#b91c1c;'>Stop server</button>
          <p id='status' class='muted'></p>
          <p id='adminMsg' class='muted'></p>
        </div>
      </div>
      <section class='card'>
        <h2>Solver progress</h2>
        <pre id='progress' style='background:#0b1220;color:#dbeafe;padding:8px;max-height:240px;overflow:auto;'></pre>
      </section>
      <section class='card'>
        <h2>Results</h2>
        <div id='results'></div>
      </section>
    </main>
    <script>
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
      async function setN(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/set_n', {{n:Number(document.getElementById('nInput').value)}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
      }}
      async function resetDb(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/reset');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
      }}
      async function setFinalized() {{
        const final = document.getElementById('finalToggle').checked;
        const data = await post('/api/admin/finalize', {{finalized: final}});
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
      }}
      async function runMatch(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/run');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
      }}
      async function interruptMatch(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/interrupt');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
      }}
      async function stopServer(btn) {{
        pulseButton(btn);
        const data = await post('/api/admin/stop');
        setAdminMsg(data.message || data.error || 'Completed', !data.__ok || !!data.error);
        if (data.__ok && !data.error) {{
          if (pollHandle) {{
            clearInterval(pollHandle);
            pollHandle = null;
          }}
          document.getElementById('status').innerText = 'Stopping server...';
        }}
      }}

      function renderResults(data) {{
        const results = document.getElementById('results');
        if (!data.latest_run) {{ results.innerHTML = '<p>No completed run yet.</p>'; return; }}
        const run = data.latest_run;
        let html = `<p>Run #${{run.id}} status: <strong>${{run.status}}</strong>; utility=${{run.utility ?? 'n/a'}}, penalty=${{run.penalty ?? 'n/a'}}, overlap violations=${{run.overlap_count ?? 'n/a'}}</p>`;

        if (data.selected_topics.length) {{
          html += '<h3>Selected topics</h3><ul>' + data.selected_topics.map(t => `<li>#${{t.topic_id + 1}} (${{t.partition}}): ${{t.title}}</li>`).join('') + '</ul>';
        }}

        if (data.assignments.length) {{
          html += '<h3>Assignments</h3><table><thead><tr><th>Student</th><th>Main</th><th>M score</th><th>Shadow</th><th>S score</th></tr></thead><tbody>';
          html += data.assignments.map(a => `<tr><td>S${{a.student_id + 1}}</td><td>${{a.main_title}}</td><td>${{a.main_score}}</td><td>${{a.shadow_title}}</td><td>${{a.shadow_score}}</td></tr>`).join('');
          html += '</tbody></table>';
          const avgMain = (data.assignments.reduce((acc,x)=>acc+x.main_score,0)/data.assignments.length).toFixed(2);
          const avgShadow = (data.assignments.reduce((acc,x)=>acc+x.shadow_score,0)/data.assignments.length).toFixed(2);
          html += `<p><strong>Average scores:</strong> main=${{avgMain}}, shadow=${{avgShadow}}</p>`;
        }}

        if (data.overlaps.length) {{
          html += '<h3>Overlap violations</h3><ul>' + data.overlaps.map(o => `<li>S${{o.s1 + 1}} and S${{o.s2 + 1}}</li>`).join('') + '</ul>';
        }}
        results.innerHTML = html;
      }}

      async function poll() {{
        try {{
          const r = await fetch('/api/admin/status');
          const data = await r.json();
          document.getElementById('status').innerText = data.running ? 'Matching in progress...' : 'Idle';
          document.getElementById('finalToggle').checked = data.finalized;
          document.getElementById('nInput').value = data.n;
          document.getElementById('progress').innerText = data.progress_logs.join('\n');
          renderResults(data);
        }} catch (err) {{
          document.getElementById('status').innerText = 'Server unavailable';
        }}
      }}
      let pollHandle = setInterval(poll, 1200);
      poll();
    </script>
    </body></html>
    """


def save_preferences(sid: int, scores: list[int]) -> tuple[bool, str]:
    conn = db_conn()
    n = int(get_meta(conn, "n", "8"))
    if sid < 0 or sid >= n:
        conn.close()
        return False, "Invalid student id."
    finalized = get_meta(conn, "finalized", "0") == "1"
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

    now = datetime.utcnow().isoformat()
    with conn:
        for j, score in enumerate(scores):
            conn.execute(
                "UPDATE preferences SET score=?, updated_at=? WHERE student_id=? AND topic_id=?",
                (score, now, sid, j),
            )
    conn.close()
    return True, "Preferences saved."


def save_topic_title(sid: int, title: str) -> tuple[bool, str]:
    conn = db_conn()
    n = int(get_meta(conn, "n", "8"))
    if sid < 0 or sid >= n:
        conn.close()
        return False, "Invalid student id."
    finalized = get_meta(conn, "finalized", "0") == "1"
    if finalized:
        conn.close()
        return False, "Matching finalized; edits are locked."
    clean = title.strip()
    if not clean:
        conn.close()
        return False, "Topic title cannot be empty."
    with conn:
        conn.execute("UPDATE students SET topic_title=? WHERE id=?", (clean[:200], sid))
    conn.close()
    return True, "Topic title saved."


def collect_problem(conn: sqlite3.Connection) -> tuple[list[str], list[list[int]]]:
    students = get_students(conn)
    n = len(students)
    topics = [s["topic_title"] for s in students]
    matrix = []
    for i in range(n):
        matrix.append(get_pref_row(conn, i, n))
    return topics, matrix


def run_matching_background(run_id: int, stop_event: threading.Event) -> None:
    conn = db_conn()

    def log(msg: str) -> None:
        with conn:
            idx = conn.execute("SELECT COALESCE(MAX(idx), -1)+1 FROM progress_logs WHERE run_id=?", (run_id,)).fetchone()[0]
            conn.execute("INSERT INTO progress_logs(run_id, idx, message) VALUES (?, ?, ?)", (run_id, idx, msg))

    try:
        topics, matrix = collect_problem(conn)
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
                    datetime.utcnow().isoformat(),
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
                (datetime.utcnow().isoformat(), f"error: {exc}", run_id),
            )
            idx = conn.execute("SELECT COALESCE(MAX(idx), -1)+1 FROM progress_logs WHERE run_id=?", (run_id,)).fetchone()[0]
            conn.execute("INSERT INTO progress_logs(run_id, idx, message) VALUES (?, ?, ?)", (run_id, idx, f"ERROR: {exc}"))
    finally:
        conn.close()
        with run_state_lock:
            run_state["thread"] = None
            run_state["stop_event"] = None


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

    def shutdown_later() -> None:
        if run_thread is not None and run_thread.is_alive():
            run_thread.join(timeout=5)
        time.sleep(0.2)
        try:
            server.shutdown()
        except Exception:
            pass

    threading.Thread(target=shutdown_later, daemon=True).start()
    return True


def application(environ, start_response):
    initialize_db()
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

    if method == "POST" and path == "/api/admin/set_n":
        data = read_json(environ)
        n = int(data.get("n", 8))
        n = max(4, n)
        conn = db_conn()
        with conn:
            set_meta(conn, "n", str(n))
            conn.execute("DELETE FROM students")
            conn.execute("DELETE FROM preferences")
        conn.close()
        ensure_students_and_preferences()
        return json_response(start_response, {"message": f"Reset to n={n}."})

    if method == "POST" and path == "/api/admin/reset":
        conn = db_conn()
        with conn:
            conn.execute("DELETE FROM students")
            conn.execute("DELETE FROM preferences")
            conn.execute("DELETE FROM match_runs")
            conn.execute("DELETE FROM progress_logs")
            conn.execute("DELETE FROM selected_topics")
            conn.execute("DELETE FROM assignments")
            conn.execute("DELETE FROM overlaps")
            set_meta(conn, "n", "8")
            set_meta(conn, "finalized", "0")
        conn.close()
        ensure_students_and_preferences()
        return json_response(start_response, {"message": "Database reset."})

    if method == "POST" and path == "/api/admin/finalize":
        data = read_json(environ)
        conn = db_conn()
        with conn:
            set_meta(conn, "finalized", "1" if data.get("finalized") else "0")
        conn.close()
        return json_response(start_response, {"message": "Updated."})

    if method == "POST" and path == "/api/admin/run":
        with run_state_lock:
            if run_state["thread"] is not None:
                return json_response(start_response, {"error": "Run already in progress."}, "409 Conflict")
            conn = db_conn()
            finalized = 1 if get_meta(conn, "finalized", "0") == "1" else 0
            with conn:
                cur = conn.execute(
                    "INSERT INTO match_runs(started_at, status, finalized_snapshot) VALUES (?, ?, ?)",
                    (datetime.utcnow().isoformat(), "running", finalized),
                )
                run_id = cur.lastrowid
            conn.close()

            stop_event = threading.Event()
            t = threading.Thread(target=run_matching_background, args=(run_id, stop_event), daemon=True)
            run_state["thread"] = t
            run_state["stop_event"] = stop_event
            run_state["current_run_id"] = run_id
            t.start()
        return json_response(start_response, {"message": "Run started."})

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
        return json_response(start_response, {"message": "Server shutdown requested."})

    if method == "GET" and path == "/api/admin/status":
        conn = db_conn()
        n = int(get_meta(conn, "n", "8"))
        finalized = get_meta(conn, "finalized", "0") == "1"
        latest = conn.execute("SELECT * FROM match_runs ORDER BY id DESC LIMIT 1").fetchone()
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
        conn.close()

        with run_state_lock:
            running = run_state["thread"] is not None
        return json_response(
            start_response,
            {
                "n": n,
                "finalized": finalized,
                "running": running,
                "progress_logs": logs,
                "latest_run": dict(latest) if latest else None,
                "selected_topics": selected_topics,
                "assignments": assignments,
                "overlaps": overlaps,
            },
        )

    return html_response(start_response, "Not found", "404 Not Found")


if __name__ == "__main__":
    initialize_db()
    with make_server("0.0.0.0", 8000, application) as server:
        with server_state_lock:
            server_instance = server
        try:
            print("Serving on http://0.0.0.0:8000")
            server.serve_forever()
        finally:
            with server_state_lock:
                server_instance = None
