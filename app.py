from __future__ import annotations

from html import escape
from urllib.parse import parse_qs
from wsgiref.simple_server import make_server



CSS = """
body { font-family: Arial, sans-serif; background: #f5f7fb; margin: 0; }
.container { max-width: 1100px; margin: 2rem auto; background: #fff; padding: 1.25rem; border-radius: 12px; box-shadow: 0 1px 8px rgba(0,0,0,.08); }
h1, h2, h3 { margin-top: 1rem; }
label { display: block; margin: .35rem 0; }
input[type='text'] { width: 100%; padding: .4rem; }
input[type='number'] { padding: .25rem; }
button { margin-top: 1rem; background: #2563eb; color: #fff; border: 0; border-radius: 6px; padding: .5rem .9rem; cursor: pointer; }
button.secondary { background: #64748b; margin-left: .5rem; }
.topics-grid { display: grid; grid-template-columns: repeat(auto-fit,minmax(250px,1fr)); gap: .7rem; }
.table-wrap { overflow-x: auto; }
table { border-collapse: collapse; width: 100%; margin-top: .75rem; }
th, td { border: 1px solid #dbe2ea; padding: .35rem; text-align: center; }
.score { width: 3.6rem; }
.error { margin-top: 1rem; color: #b91c1c; font-weight: 600; }
"""


def get_int(data: dict[str, list[str]], key: str, default: int) -> int:
    try:
        return int(data.get(key, [str(default)])[0])
    except Exception:
        return default


def render_page(n: int, topics: list[str], preferences: list[list[int]], result: dict | None = None, error: str | None = None) -> str:
    topic_inputs = "".join(
        f"<label>Student {i + 1} topic<input type='text' name='topic_{i}' value='{escape(topics[i])}' /></label>"
        for i in range(n)
    )

    header_cols = "".join(f"<th>{j + 1}</th>" for j in range(n))
    rows = []
    for i in range(n):
        cells = "".join(
            f"<td><input class='score' type='number' min='0' max='5' name='pref_{i}_{j}' value='{preferences[i][j]}' /></td>"
            for j in range(n)
        )
        rows.append(f"<tr><th>S{i + 1}</th>{cells}</tr>")

    result_html = ""
    if error:
        result_html += f"<div class='error'>{escape(error)}</div>"
    if result:
        selected = "".join(
            f"<li>#{t['id'] + 1} ({t['partition']}): {escape(t['title'])}</li>"
            for t in result["selected_topics"]
        )
        assignments = "".join(
            "<tr>"
            f"<td>S{a['student'] + 1}</td>"
            f"<td>{escape(a['main_title'])}</td><td>{a['main_score']}</td>"
            f"<td>{escape(a['shadow_title'])}</td><td>{a['shadow_score']}</td>"
            "</tr>"
            for a in result["assignments"]
        )
        result_html += f"""
        <section>
          <h2>Results</h2>
          <p>Status: <strong>{result['status']}</strong> · Selected topics: {result['K']} · Utility: {result['utility']} · Penalty: {result['penalty']}</p>
          <h3>Selected topics</h3><ul>{selected}</ul>
          <h3>Student assignments</h3>
          <div class='table-wrap'><table><thead><tr><th>Student</th><th>Main</th><th>Score</th><th>Shadow</th><th>Score</th></tr></thead><tbody>{assignments}</tbody></table></div>
        </section>
        """

    return f"""
<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Topic Matching Tool</title><style>{CSS}</style></head><body>
<main class='container'>
<h1>Class Topic Matching</h1>
<p>Step 1: Enter one topic per student. Step 2: Enter voting scores (0..5). Step 3: Run matching.</p>
<form method='post'>
<label for='n'>Number of students/topics:</label>
<input id='n' name='n' type='number' min='4' value='{n}' />
<button type='submit' class='secondary'>Update Size</button>
<h2>Topic proposals</h2><div class='topics-grid'>{topic_inputs}</div>
<h2>Voting matrix (0 veto, 5 best)</h2>
<div class='table-wrap'><table><thead><tr><th>Voter \\ Topic</th>{header_cols}</tr></thead><tbody>{''.join(rows)}</tbody></table></div>
<button type='submit'>Run Matching</button>
</form>{result_html}</main></body></html>
"""


def application(environ, start_response):
    n = 8
    topics = [f"Student {i + 1} topic" for i in range(n)]
    prefs = [[5 if i == j else 3 for j in range(n)] for i in range(n)]
    result = None
    error = None

    if environ.get("REQUEST_METHOD") == "POST":
        try:
            size = int(environ.get("CONTENT_LENGTH") or 0)
        except ValueError:
            size = 0
        body = environ["wsgi.input"].read(size).decode("utf-8")
        data = parse_qs(body)
        n = max(4, get_int(data, "n", 8))
        topics = [data.get(f"topic_{i}", [f"Topic {i + 1}"])[0].strip() or f"Topic {i + 1}" for i in range(n)]
        prefs = []
        for i in range(n):
            row = []
            for j in range(n):
                row.append(max(0, min(5, get_int(data, f"pref_{i}_{j}", 0))))
            row[i] = 5
            prefs.append(row)
        try:
            from match import solve_with_preferences
            result = solve_with_preferences(topic_titles=topics, preferences=prefs, M=4, time_limit_s=15)
        except Exception as exc:  # noqa: BLE001
            error = str(exc)

    html = render_page(n=n, topics=topics, preferences=prefs, result=result, error=error)
    start_response("200 OK", [("Content-Type", "text/html; charset=utf-8")])
    return [html.encode("utf-8")]


if __name__ == "__main__":
    with make_server("0.0.0.0", 8000, application) as server:
        print("Serving on http://0.0.0.0:8000")
        server.serve_forever()
