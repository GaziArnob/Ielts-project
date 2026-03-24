import html
import json
import mimetypes
import os
import secrets
import sqlite3
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from spreadsheet_export import export_database_to_workbook


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DATA_DIR = Path(os.environ.get("APP_DATA_DIR", str(BASE_DIR))).resolve()
UPLOADS_DIR = DATA_DIR / "uploads"
EXPORTS_DIR = DATA_DIR / "output" / "spreadsheet"
EXPORT_WORKBOOK = EXPORTS_DIR / "bandforge_data.xlsx"
DB_PATH = DATA_DIR / "ielts_app.db"
HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8000"))
ADMIN_PASSWORD = os.environ.get("IELTS_ADMIN_PASSWORD", "admin123")
ADMIN_COOKIE = "ielts_admin_session"
ADMIN_SESSIONS = set()

LISTENING_SCRIPT = [
    "Good morning. This is a message for students joining the weekend study camp at Greenford College.",
    "Registration starts at 8 30 in the main hall, and the first workshop begins at 9 15.",
    "Students should bring a notebook, a blue pen, and their student card.",
    "Lunch is served in the riverside cafe at 12 45, and vegetarian meals are available.",
    "In the afternoon, there is a pronunciation class in room B12, followed by a speaking circle in room C4.",
    "If you need help, call the campus office on 01632 480 221.",
]

LISTENING_ITEMS = [
    {"id": "listen_1", "prompt": "What time does registration start?", "answer": "8 30"},
    {"id": "listen_2", "prompt": "Where does registration happen?", "answer": "main hall"},
    {"id": "listen_3", "prompt": "What color pen should students bring?", "answer": "blue"},
    {"id": "listen_4", "prompt": "What time is lunch served?", "answer": "12 45"},
    {"id": "listen_5", "prompt": "Which room is used for the pronunciation class?", "answer": "b12"},
    {"id": "listen_6", "prompt": "What is the campus office phone number?", "answer": "01632 480 221"},
]

READING_PASSAGE = (
    "Cities around the world are redesigning public libraries to serve as learning hubs rather than silent storage rooms. "
    "Modern libraries still protect books, but they now also provide digital labs, language corners, and community workshops. "
    "Researchers note that this shift is especially valuable for students and job seekers because public libraries often offer "
    "free internet, quiet study areas, and staff support. One survey found that visitors stayed longer when libraries included "
    "flexible seating and technology access. Another report argued that a successful library should remain calm and welcoming "
    "while adapting to new patterns of study, communication, and work."
)

READING_ITEMS = [
    {"id": "read_1", "prompt": "Libraries are being redesigned to become what?", "answer": "learning hubs"},
    {"id": "read_2", "prompt": "Who benefits especially from this change?", "answer": "students and job seekers"},
    {"id": "read_3", "prompt": "What do public libraries often offer for free?", "answer": "internet"},
    {"id": "read_4", "prompt": "What made visitors stay longer in one survey?", "answer": "flexible seating and technology access"},
    {"id": "read_5", "prompt": "A successful library should remain calm and what else?", "answer": "welcoming"},
]

WRITING_TASK_1 = (
    "You work for a community center. Write an email to a new student explaining the study facilities, opening times, and how to join a speaking club. Write at least 150 words."
)

WRITING_TASK_2 = (
    "Some people believe students should study only subjects that lead directly to jobs. Others think education should include a wide range of subjects. Discuss both views and give your own opinion. Write at least 250 words."
)

SPEAKING_PART_1 = [
    "Introduce yourself and talk about your hometown.",
    "Describe your daily study routine.",
]

SPEAKING_PART_2 = (
    "Describe a skill you want to improve in English. You should say what it is, why it matters to you, how you are trying to improve it, and explain what makes it difficult."
)

SPEAKING_PART_3 = [
    "Why do some learners improve speaking faster than writing?",
    "How can technology help language learners prepare for international exams?",
]


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def esc(value):
    return html.escape(str(value))


def init_db():
    connection = sqlite3.connect(DB_PATH)
    cursor = connection.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS candidates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_code TEXT NOT NULL UNIQUE,
            full_name TEXT NOT NULL,
            email TEXT NOT NULL,
            age INTEGER NOT NULL,
            country TEXT NOT NULL,
            target_band REAL NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            candidate_id INTEGER NOT NULL,
            status TEXT NOT NULL,
            current_step TEXT NOT NULL,
            listening_answers TEXT NOT NULL,
            listening_correct INTEGER,
            listening_band REAL,
            reading_answers TEXT NOT NULL,
            reading_correct INTEGER,
            reading_band REAL,
            writing_task_1 TEXT NOT NULL,
            writing_task_2 TEXT NOT NULL,
            speaking_part_1_text TEXT NOT NULL,
            speaking_part_2_text TEXT NOT NULL,
            speaking_part_3_text TEXT NOT NULL,
            speaking_part_1_audio TEXT NOT NULL,
            speaking_part_2_audio TEXT NOT NULL,
            speaking_part_3_audio TEXT NOT NULL,
            candidate_notes TEXT NOT NULL,
            writing_band REAL,
            speaking_band REAL,
            overall_band REAL,
            position_label TEXT NOT NULL,
            examiner_feedback TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            submitted_at TEXT NOT NULL,
            scored_at TEXT NOT NULL,
            FOREIGN KEY (candidate_id) REFERENCES candidates(id)
        )
        """
    )
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    cursor.execute("INSERT OR IGNORE INTO app_settings (key, value) VALUES ('public_access', 'on')")
    connection.commit()
    connection.close()


def db():
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def get_setting(key, default_value=""):
    connection = db()
    row = connection.execute("SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()
    connection.close()
    return row["value"] if row else default_value


def set_setting(key, value):
    connection = db()
    connection.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)", (key, value))
    connection.commit()
    connection.close()


def public_access_on():
    return get_setting("public_access", "on") == "on"


def normalize_answer(value):
    return " ".join(str(value).strip().lower().replace("-", " ").split())


def score_short_answers(items, answers):
    total = 0
    for item in items:
        if normalize_answer(answers.get(item["id"], "")) == normalize_answer(item["answer"]):
            total += 1
    return total


def convert_correct_to_band(correct, total):
    ratio = correct / total if total else 0
    if ratio >= 0.95:
        return 9.0
    if ratio >= 0.85:
        return 8.0
    if ratio >= 0.75:
        return 7.0
    if ratio >= 0.65:
        return 6.5
    if ratio >= 0.55:
        return 6.0
    if ratio >= 0.45:
        return 5.5
    if ratio >= 0.35:
        return 5.0
    if ratio >= 0.25:
        return 4.5
    return 4.0


def round_overall(score):
    decimal = score % 1
    if decimal < 0.25:
        return float(int(score))
    if decimal < 0.75:
        return float(int(score) + 0.5)
    return float(int(score) + 1)


def position_label(score):
    if score >= 8:
        return "Exam Ready"
    if score >= 7:
        return "Strong Progress"
    if score >= 6:
        return "Developing Competence"
    if score >= 5:
        return "Emerging Control"
    return "Foundation Stage"


def generate_candidate_code():
    return f"IELTS-{secrets.token_hex(3).upper()}"


def auth_token(handler):
    cookie = SimpleCookie()
    cookie.load(handler.headers.get("Cookie", ""))
    morsel = cookie.get(ADMIN_COOKIE)
    return morsel.value if morsel else ""


def is_admin(handler):
    return auth_token(handler) in ADMIN_SESSIONS


def make_admin_cookie():
    token = secrets.token_urlsafe(24)
    ADMIN_SESSIONS.add(token)
    return token


def drop_admin_cookie(handler):
    token = auth_token(handler)
    if token in ADMIN_SESSIONS:
        ADMIN_SESSIONS.remove(token)


def parse_form(handler):
    content_type = handler.headers.get("Content-Type", "")
    length = int(handler.headers.get("Content-Length", "0"))
    raw = handler.rfile.read(length)
    if "multipart/form-data" in content_type:
        envelope = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + raw
        message = BytesParser(policy=default).parsebytes(envelope)
        fields = {}
        files = {}
        for part in message.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            filename = part.get_filename()
            payload = part.get_payload(decode=True) or b""
            if filename:
                files[name] = {"filename": filename, "content_type": part.get_content_type(), "data": payload}
            else:
                fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="ignore")
        return fields, files
    parsed = parse_qs(raw.decode("utf-8"), keep_blank_values=True)
    return {key: values[0] for key, values in parsed.items()}, {}


def audio_extension(file_info):
    suffix = Path(file_info["filename"]).suffix.lower()
    if suffix in {".webm", ".wav", ".mp3", ".m4a", ".ogg"}:
        return suffix
    return mimetypes.guess_extension(file_info["content_type"] or "") or ".webm"


def save_audio_file(candidate_code, field_name, file_info):
    if not file_info or not file_info.get("data"):
        return ""
    name = f"{candidate_code.lower()}_{field_name}{audio_extension(file_info)}"
    target = UPLOADS_DIR / name
    target.write_bytes(file_info["data"])
    return name


def get_submission_by_code(code):
    connection = db()
    row = connection.execute(
        """
        SELECT submissions.*, candidates.candidate_code, candidates.full_name, candidates.email,
               candidates.age, candidates.country, candidates.target_band
        FROM submissions
        JOIN candidates ON candidates.id = submissions.candidate_id
        WHERE candidates.candidate_code = ?
        ORDER BY submissions.id DESC
        LIMIT 1
        """,
        (code,),
    ).fetchone()
    connection.close()
    return row


def get_submission_by_id(submission_id):
    connection = db()
    row = connection.execute(
        """
        SELECT submissions.*, candidates.candidate_code, candidates.full_name, candidates.email,
               candidates.age, candidates.country, candidates.target_band
        FROM submissions
        JOIN candidates ON candidates.id = submissions.candidate_id
        WHERE submissions.id = ?
        """,
        (submission_id,),
    ).fetchone()
    connection.close()
    return row


def export_reports():
    try:
        export_database_to_workbook(DB_PATH, EXPORT_WORKBOOK)
    except Exception as exc:
        print(f"Spreadsheet export skipped: {exc}")


def top_nav(admin=False):
    admin_link = '<a href="/admin">Admin</a>' if not admin else '<a href="/admin/logout">Logout</a>'
    return f"""
    <header class="site-header">
      <a class="logo" href="/">BandForge</a>
      <nav class="site-nav">
        <a href="/">Home</a>
        <a href="/start">Start Exam</a>
        <a href="/results">Results</a>
        {admin_link}
        <button id="themeToggle" class="theme-toggle" type="button">Dark Mode</button>
      </nav>
    </header>
    """


def page_shell(title, body, admin=False):
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{esc(title)}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="/static/styles.css">
</head>
<body>
  <div class="page-shell">
    {top_nav(admin)}
    {body}
  </div>
  <script>
    window.__LISTENING_SCRIPT__ = {json.dumps(LISTENING_SCRIPT)};
    window.__SPEAKING_PART_1__ = {json.dumps(SPEAKING_PART_1)};
    window.__SPEAKING_PART_2__ = {json.dumps(SPEAKING_PART_2)};
    window.__SPEAKING_PART_3__ = {json.dumps(SPEAKING_PART_3)};
  </script>
  <script src="/static/app.js"></script>
</body>
</html>"""


def status_banner():
    if public_access_on():
        return '<p class="status-banner online">Public exam is live now.</p>'
    return '<p class="status-banner offline">Public exam is currently turned off by the admin.</p>'


def render_home():
    start_link = "/start" if public_access_on() else "#"
    start_class = "button button-primary" if public_access_on() else "button button-disabled"
    body = f"""
    <main class="home-grid">
      <section class="hero-card">
        {status_banner()}
        <p class="eyebrow">Public IELTS Platform</p>
        <h1>Real multi-page exam flow with audio listening, spoken prompts, essays, and examiner review.</h1>
        <p class="lead">Everyone can use the public exam when it is turned on. Only the admin area can control availability, review candidates, publish scores, and download the Excel files.</p>
        <div class="hero-actions">
          <a class="{start_class}" href="{start_link}">Start Public Exam</a>
          <a class="button button-secondary" href="/results">Check Candidate Result</a>
        </div>
      </section>
      <section class="feature-grid">
        <article class="feature-card"><span>Listening</span><strong>Audio prompt plus typed answers</strong></article>
        <article class="feature-card"><span>Reading</span><strong>Passage with short answer questions</strong></article>
        <article class="feature-card"><span>Writing</span><strong>Task 1 and Task 2 essays on separate page</strong></article>
        <article class="feature-card"><span>Speaking</span><strong>Examiner voice prompts with transcript and audio upload</strong></article>
      </section>
    </main>
    """
    return page_shell("BandForge IELTS Exam", body)


def render_public_closed():
    body = """
    <main class="single-panel">
      <section class="panel-card">
        <p class="eyebrow">Public Access Off</p>
        <h1>The exam is not accepting new candidates right now.</h1>
        <p>You can still check already published results, and the admin dashboard stays available to you.</p>
        <div class="hero-actions">
          <a class="button button-secondary" href="/results">Check Results</a>
          <a class="button button-primary" href="/admin">Open Admin</a>
        </div>
      </section>
    </main>
    """
    return page_shell("Exam Offline", body)


def render_start(error_message=""):
    error = f'<p class="form-alert">{esc(error_message)}</p>' if error_message else ""
    body = f"""
    <main class="single-panel">
      <section class="panel-card">
        <p class="eyebrow">Candidate Registration</p>
        <h1>Create a candidate and begin the IELTS journey</h1>
        <p>This is the first step. After this page the candidate moves through Listening, Reading, Writing, and Speaking on separate pages.</p>
        {error}
        <form method="post" action="/start" class="form-grid">
          <label><span>Full name</span><input name="full_name" type="text" required></label>
          <label><span>Email</span><input name="email" type="email" required></label>
          <label><span>Age</span><input name="age" type="number" min="10" max="80" required></label>
          <label><span>Country</span><input name="country" type="text" required></label>
          <label><span>Target band</span><select name="target_band"><option>5.5</option><option selected>6.0</option><option>6.5</option><option>7.0</option><option>7.5</option><option>8.0</option></select></label>
          <button class="button button-primary full-width" type="submit">Continue To Listening Page</button>
        </form>
      </section>
    </main>
    """
    return page_shell("Start Public Exam", body)


def progress_header(code, step, title, description):
    return f"""
    <section class="progress-head">
      <div>
        <p class="eyebrow">Candidate {esc(code)}</p>
        <h1>{esc(title)}</h1>
        <p>{esc(description)}</p>
      </div>
      <div class="progress-pills">
        <span class="progress-pill {'active' if step == 'listening' else ''}">Listening</span>
        <span class="progress-pill {'active' if step == 'reading' else ''}">Reading</span>
        <span class="progress-pill {'active' if step == 'writing' else ''}">Writing</span>
        <span class="progress-pill {'active' if step == 'speaking' else ''}">Speaking</span>
      </div>
    </section>
    """


def render_listening(row, error_message=""):
    saved = json.loads(row["listening_answers"] or "{}")
    answers = []
    for idx, item in enumerate(LISTENING_ITEMS, start=1):
        answers.append(
            f"""
            <label class="answer-card">
              <span class="small-label">Question {idx}</span>
              <strong>{esc(item['prompt'])}</strong>
              <input name="{esc(item['id'])}" value="{esc(saved.get(item['id'], ''))}" type="text" required>
            </label>
            """
        )
    error = f'<p class="form-alert">{esc(error_message)}</p>' if error_message else ""
    script_lines = "".join(f"<li>{esc(line)}</li>" for line in LISTENING_SCRIPT)
    body = f"""
    <main class="exam-page">
      {progress_header(row['candidate_code'], 'listening', 'Listening Page', 'Play the audio and type short answers. There are no multiple choice questions here.')}
      {error}
      <form method="post" action="/exam/listening?code={esc(row['candidate_code'])}" class="exam-section">
        <div class="split-layout">
          <article class="audio-panel">
            <p class="eyebrow">Audio Console</p>
            <h2>Listening script player</h2>
            <p>The browser will read the script aloud. Candidates can replay it while practicing.</p>
            <div class="audio-controls">
              <button class="button button-primary" type="button" id="playListening">Play Audio</button>
              <button class="button button-secondary" type="button" id="stopAudio">Stop Audio</button>
            </div>
            <ol class="script-list">{script_lines}</ol>
          </article>
          <section class="answer-grid">{''.join(answers)}</section>
        </div>
        <button class="button button-primary submit-row" type="submit">Save Listening And Continue</button>
      </form>
    </main>
    """
    return page_shell("Listening Page", body)


def render_reading(row, error_message=""):
    saved = json.loads(row["reading_answers"] or "{}")
    answers = []
    for idx, item in enumerate(READING_ITEMS, start=1):
        answers.append(
            f"""
            <label class="answer-card">
              <span class="small-label">Question {idx}</span>
              <strong>{esc(item['prompt'])}</strong>
              <input name="{esc(item['id'])}" value="{esc(saved.get(item['id'], ''))}" type="text" required>
            </label>
            """
        )
    error = f'<p class="form-alert">{esc(error_message)}</p>' if error_message else ""
    body = f"""
    <main class="exam-page">
      {progress_header(row['candidate_code'], 'reading', 'Reading Page', 'Read the passage and answer with short text responses.')}
      {error}
      <form method="post" action="/exam/reading?code={esc(row['candidate_code'])}" class="exam-section">
        <div class="split-layout">
          <article class="reading-panel">
            <p class="eyebrow">Reading Passage</p>
            <h2>Academic style passage</h2>
            <p>{esc(READING_PASSAGE)}</p>
          </article>
          <section class="answer-grid">{''.join(answers)}</section>
        </div>
        <button class="button button-primary submit-row" type="submit">Save Reading And Continue</button>
      </form>
    </main>
    """
    return page_shell("Reading Page", body)


def render_writing(row, error_message=""):
    error = f'<p class="form-alert">{esc(error_message)}</p>' if error_message else ""
    body = f"""
    <main class="exam-page">
      {progress_header(row['candidate_code'], 'writing', 'Writing Page', 'Write a shorter practical response and then a longer essay response.')}
      {error}
      <form method="post" action="/exam/writing?code={esc(row['candidate_code'])}" class="exam-section">
        <label class="essay-card">
          <span class="small-label">Task 1</span>
          <strong>{esc(WRITING_TASK_1)}</strong>
          <textarea name="writing_task_1" rows="10" required>{esc(row['writing_task_1'])}</textarea>
        </label>
        <label class="essay-card">
          <span class="small-label">Task 2</span>
          <strong>{esc(WRITING_TASK_2)}</strong>
          <textarea name="writing_task_2" rows="14" required>{esc(row['writing_task_2'])}</textarea>
        </label>
        <button class="button button-primary submit-row" type="submit">Save Writing And Continue</button>
      </form>
    </main>
    """
    return page_shell("Writing Page", body)


def speaking_prompt_block(title, prompts, button_id):
    lines = prompts if isinstance(prompts, list) else [prompts]
    items = "".join(f"<li>{esc(line)}</li>" for line in lines)
    joined = " ".join(lines)
    return f"""
    <div class="prompt-box">
      <div class="prompt-head">
        <div>
          <span class="small-label">{esc(title)}</span>
          <ul class="prompt-list">{items}</ul>
        </div>
        <button class="button button-secondary prompt-play" id="{button_id}" type="button" data-text="{esc(joined)}">Play Examiner Voice</button>
      </div>
    </div>
    """


def render_speaking(row, error_message=""):
    error = f'<p class="form-alert">{esc(error_message)}</p>' if error_message else ""
    body = f"""
    <main class="exam-page">
      {progress_header(row['candidate_code'], 'speaking', 'Speaking Page', 'Listen to the examiner voice prompts, answer aloud, and save transcript notes or audio recordings.')}
      {error}
      <form method="post" enctype="multipart/form-data" action="/exam/speaking?code={esc(row['candidate_code'])}" class="exam-section" id="speakingForm">
        {speaking_prompt_block('Part 1', SPEAKING_PART_1, 'playSpeaking1')}
        <label class="essay-card">
          <span class="small-label">Part 1 Transcript Or Notes</span>
          <textarea name="speaking_part_1_text" rows="6" required>{esc(row['speaking_part_1_text'])}</textarea>
          <input name="speaking_part_1_audio" type="file" accept="audio/*" capture="user">
          <div class="record-strip" data-field="speaking_part_1_audio" data-preview="preview_part_1">
            <button class="button button-secondary record-start" type="button">Start Recording</button>
            <button class="button button-secondary record-stop" type="button">Stop Recording</button>
            <audio controls id="preview_part_1"></audio>
          </div>
        </label>
        {speaking_prompt_block('Part 2', SPEAKING_PART_2, 'playSpeaking2')}
        <label class="essay-card">
          <span class="small-label">Part 2 Transcript Or Notes</span>
          <textarea name="speaking_part_2_text" rows="7" required>{esc(row['speaking_part_2_text'])}</textarea>
          <input name="speaking_part_2_audio" type="file" accept="audio/*" capture="user">
          <div class="record-strip" data-field="speaking_part_2_audio" data-preview="preview_part_2">
            <button class="button button-secondary record-start" type="button">Start Recording</button>
            <button class="button button-secondary record-stop" type="button">Stop Recording</button>
            <audio controls id="preview_part_2"></audio>
          </div>
        </label>
        {speaking_prompt_block('Part 3', SPEAKING_PART_3, 'playSpeaking3')}
        <label class="essay-card">
          <span class="small-label">Part 3 Transcript Or Notes</span>
          <textarea name="speaking_part_3_text" rows="7" required>{esc(row['speaking_part_3_text'])}</textarea>
          <input name="speaking_part_3_audio" type="file" accept="audio/*" capture="user">
          <div class="record-strip" data-field="speaking_part_3_audio" data-preview="preview_part_3">
            <button class="button button-secondary record-start" type="button">Start Recording</button>
            <button class="button button-secondary record-stop" type="button">Stop Recording</button>
            <audio controls id="preview_part_3"></audio>
          </div>
        </label>
        <label class="essay-card">
          <span class="small-label">Candidate Notes</span>
          <textarea name="candidate_notes" rows="5">{esc(row['candidate_notes'])}</textarea>
        </label>
        <button class="button button-primary submit-row" type="submit">Submit Full Exam</button>
      </form>
    </main>
    """
    return page_shell("Speaking Page", body)


def render_submitted(row):
    body = f"""
    <main class="single-panel">
      <section class="panel-card">
        <p class="eyebrow">Submission Complete</p>
        <h1>Your public exam has been submitted</h1>
        <p>Candidate code: <strong>{esc(row['candidate_code'])}</strong></p>
        <p>The listening and reading pages are scored automatically, and the writing and speaking pages wait for examiner review before the final band is published.</p>
        <div class="hero-actions">
          <a class="button button-primary" href="/results">Check Result Later</a>
          <a class="button button-secondary" href="/">Back Home</a>
        </div>
      </section>
    </main>
    """
    return page_shell("Exam Submitted", body)


def render_results_lookup(message="", result_html=""):
    alert = f'<p class="form-alert">{esc(message)}</p>' if message else ""
    body = f"""
    <main class="single-panel">
      <section class="panel-card">
        <p class="eyebrow">Result Lookup</p>
        <h1>Check a candidate result</h1>
        <p>Enter the candidate code received at the end of the public exam.</p>
        {alert}
        <form method="get" action="/results" class="lookup-form">
          <input name="code" type="text" placeholder="Example: IELTS-ABC123" required>
          <button class="button button-primary" type="submit">Find Result</button>
        </form>
      </section>
      {result_html}
    </main>
    """
    return page_shell("Check Result", body)


def published_result_card(row):
    return f"""
    <section class="published-card">
      <p class="eyebrow">Published Result</p>
      <h2>{esc(row['full_name'])}</h2>
      <div class="metric-grid">
        <article><span>Listening</span><strong>{row['listening_band']:.1f}</strong></article>
        <article><span>Reading</span><strong>{row['reading_band']:.1f}</strong></article>
        <article><span>Writing</span><strong>{row['writing_band']:.1f}</strong></article>
        <article><span>Speaking</span><strong>{row['speaking_band']:.1f}</strong></article>
        <article class="featured"><span>Overall</span><strong>{row['overall_band']:.1f}</strong></article>
      </div>
      <p class="result-chip">{esc(row['position_label'])}</p>
      <p class="feedback-box">{esc(row['examiner_feedback'])}</p>
    </section>
    """


def render_admin_login(error_message=""):
    alert = f'<p class="form-alert">{esc(error_message)}</p>' if error_message else ""
    body = f"""
    <main class="single-panel">
      <section class="panel-card admin-card">
        <p class="eyebrow">Admin Login</p>
        <h1>Private examiner control</h1>
        <p>Only the admin can switch the public app on or off, review writing and speaking, and publish the final result.</p>
        {alert}
        <form method="post" action="/admin/login" class="form-grid">
          <label><span>Password</span><input name="password" type="password" required></label>
          <button class="button button-primary full-width" type="submit">Enter Dashboard</button>
        </form>
      </section>
    </main>
    """
    return page_shell("Admin Login", body, admin=True)


def render_admin_dashboard(rows):
    cards = []
    for row in rows:
        cards.append(
            f"""
            <article class="submission-card">
              <div class="submission-meta"><span>{esc(row['candidate_code'])}</span><strong>{esc(row['status'].upper())}</strong></div>
              <h3>{esc(row['full_name'])}</h3>
              <p>{esc(row['email'])}</p>
              <p>{esc(row['country'])} / target {row['target_band']:.1f}</p>
              <p>Created {esc(row['created_at'])}</p>
              <a class="button button-primary" href="/admin/review?id={row['id']}">Review Candidate</a>
            </article>
            """
        )
    next_state = "off" if public_access_on() else "on"
    body = f"""
    <main class="admin-layout">
      <section class="panel-card">
        <p class="eyebrow">Admin Control</p>
        <h1>Public access and exam records</h1>
        <div class="admin-toolbar">
          <form method="post" action="/admin/toggle-public">
            <input type="hidden" name="next_state" value="{next_state}">
            <button class="button button-primary" type="submit">Turn Public App {next_state.upper()}</button>
          </form>
          <a class="button button-secondary" href="/exports/bandforge_data.xlsx">Download Data Workbook</a>
        </div>
        {status_banner()}
        <p class="admin-note">Default password is <code>admin123</code> unless you set <code>IELTS_ADMIN_PASSWORD</code> before running the server.</p>
      </section>
      <section class="submission-list">{''.join(cards) or '<section class="panel-card"><p>No candidates have submitted yet.</p></section>'}</section>
    </main>
    """
    return page_shell("Admin Dashboard", body, admin=True)


def audio_player(filename):
    if not filename:
        return '<p class="audio-missing">No audio uploaded for this part.</p>'
    return f'<audio controls src="/uploads/{esc(filename)}"></audio>'


def render_admin_review(row):
    listening_answers = json.loads(row["listening_answers"] or "{}")
    reading_answers = json.loads(row["reading_answers"] or "{}")
    listening_list = "".join(
        f"<li><strong>{esc(item['prompt'])}</strong><span>{esc(listening_answers.get(item['id'], ''))}</span><em>Expected: {esc(item['answer'])}</em></li>"
        for item in LISTENING_ITEMS
    )
    reading_list = "".join(
        f"<li><strong>{esc(item['prompt'])}</strong><span>{esc(reading_answers.get(item['id'], ''))}</span><em>Expected: {esc(item['answer'])}</em></li>"
        for item in READING_ITEMS
    )
    writing_selected = f"{row['writing_band']:.1f}" if row["writing_band"] is not None else "6.0"
    speaking_selected = f"{row['speaking_band']:.1f}" if row["speaking_band"] is not None else "6.0"
    feedback = row["examiner_feedback"] or ""
    options = lambda selected: "".join(f'<option value="{value}" {"selected" if value == selected else ""}>{value}</option>' for value in ["5.0", "5.5", "6.0", "6.5", "7.0", "7.5", "8.0"])
    body = f"""
    <main class="review-layout">
      <section class="panel-card review-head">
        <p class="eyebrow">Review Candidate</p>
        <h1>{esc(row['full_name'])}</h1>
        <p>{esc(row['candidate_code'])} / {esc(row['email'])} / {esc(row['country'])} / target {row['target_band']:.1f}</p>
      </section>
      <section class="review-grid">
        <article class="panel-card"><p class="eyebrow">Listening</p><p>Correct answers: {row['listening_correct']} / {len(LISTENING_ITEMS)} / band {row['listening_band']:.1f}</p><ul class="review-list">{listening_list}</ul></article>
        <article class="panel-card"><p class="eyebrow">Reading</p><p>Correct answers: {row['reading_correct']} / {len(READING_ITEMS)} / band {row['reading_band']:.1f}</p><ul class="review-list">{reading_list}</ul></article>
        <article class="panel-card"><p class="eyebrow">Writing Task 1</p><p class="long-copy">{esc(row['writing_task_1'])}</p></article>
        <article class="panel-card"><p class="eyebrow">Writing Task 2</p><p class="long-copy">{esc(row['writing_task_2'])}</p></article>
        <article class="panel-card"><p class="eyebrow">Speaking Part 1</p><p class="long-copy">{esc(row['speaking_part_1_text'])}</p>{audio_player(row['speaking_part_1_audio'])}</article>
        <article class="panel-card"><p class="eyebrow">Speaking Part 2</p><p class="long-copy">{esc(row['speaking_part_2_text'])}</p>{audio_player(row['speaking_part_2_audio'])}</article>
        <article class="panel-card"><p class="eyebrow">Speaking Part 3</p><p class="long-copy">{esc(row['speaking_part_3_text'])}</p>{audio_player(row['speaking_part_3_audio'])}</article>
        <article class="panel-card">
          <p class="eyebrow">Admin Scoring</p>
          <form method="post" action="/admin/score" class="form-grid">
            <input type="hidden" name="submission_id" value="{row['id']}">
            <label><span>Writing band</span><select name="writing_band">{options(writing_selected)}</select></label>
            <label><span>Speaking band</span><select name="speaking_band">{options(speaking_selected)}</select></label>
            <label><span>Examiner feedback</span><textarea name="examiner_feedback" rows="8" required>{esc(feedback)}</textarea></label>
            <button class="button button-primary full-width" type="submit">Publish Result</button>
          </form>
        </article>
      </section>
    </main>
    """
    return page_shell("Review Candidate", body, admin=True)


class Handler(BaseHTTPRequestHandler):
    def html_response(self, body, status=HTTPStatus.OK):
        data = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def file_response(self, path, content_type):
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, location, cookie=None, clear_cookie=False):
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        if cookie:
            self.send_header("Set-Cookie", cookie)
        if clear_cookie:
            self.send_header("Set-Cookie", f"{ADMIN_COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax")
        self.end_headers()

    def guard_public(self):
        if public_access_on():
            return False
        self.html_response(render_public_closed(), status=HTTPStatus.SERVICE_UNAVAILABLE)
        return True

    def guard_admin(self):
        if is_admin(self):
            return False
        self.redirect("/admin")
        return True

    def do_GET(self):
        parsed = urlparse(self.path)
        route = parsed.path
        params = parse_qs(parsed.query)

        if route.startswith("/static/"):
            path = STATIC_DIR / route.replace("/static/", "", 1)
            content_type = "text/plain; charset=utf-8"
            if path.suffix == ".css":
                content_type = "text/css; charset=utf-8"
            elif path.suffix == ".js":
                content_type = "application/javascript; charset=utf-8"
            self.file_response(path, content_type)
            return
        if route.startswith("/uploads/"):
            path = UPLOADS_DIR / route.replace("/uploads/", "", 1)
            self.file_response(path, mimetypes.guess_type(str(path))[0] or "application/octet-stream")
            return
        if route.startswith("/exports/"):
            path = EXPORTS_DIR / route.replace("/exports/", "", 1)
            self.file_response(path, mimetypes.guess_type(str(path))[0] or "application/octet-stream")
            return
        if route == "/":
            self.html_response(render_home())
            return
        if route == "/start":
            if self.guard_public():
                return
            self.html_response(render_start())
            return
        if route == "/exam/listening":
            if self.guard_public():
                return
            row = get_submission_by_code(params.get("code", [""])[0])
            if not row:
                self.redirect("/start")
                return
            self.html_response(render_listening(row))
            return
        if route == "/exam/reading":
            if self.guard_public():
                return
            row = get_submission_by_code(params.get("code", [""])[0])
            if not row:
                self.redirect("/start")
                return
            self.html_response(render_reading(row))
            return
        if route == "/exam/writing":
            if self.guard_public():
                return
            row = get_submission_by_code(params.get("code", [""])[0])
            if not row:
                self.redirect("/start")
                return
            self.html_response(render_writing(row))
            return
        if route == "/exam/speaking":
            if self.guard_public():
                return
            row = get_submission_by_code(params.get("code", [""])[0])
            if not row:
                self.redirect("/start")
                return
            self.html_response(render_speaking(row))
            return
        if route == "/submitted":
            row = get_submission_by_code(params.get("code", [""])[0])
            if not row:
                self.redirect("/")
                return
            self.html_response(render_submitted(row))
            return
        if route == "/results":
            code = params.get("code", [""])[0].strip()
            if not code:
                self.html_response(render_results_lookup())
                return
            row = get_submission_by_code(code)
            if not row:
                self.html_response(render_results_lookup("No candidate found for that code."), status=HTTPStatus.NOT_FOUND)
                return
            if row["status"] != "published":
                self.html_response(render_results_lookup("This candidate exists, but the final result has not been published yet."))
                return
            self.html_response(render_results_lookup(result_html=published_result_card(row)))
            return
        if route == "/admin":
            if is_admin(self):
                connection = db()
                rows = connection.execute(
                    """
                    SELECT submissions.id, submissions.status, submissions.created_at,
                           candidates.candidate_code, candidates.full_name, candidates.email,
                           candidates.country, candidates.target_band
                    FROM submissions
                    JOIN candidates ON candidates.id = submissions.candidate_id
                    ORDER BY submissions.id DESC
                    """
                ).fetchall()
                connection.close()
                self.html_response(render_admin_dashboard(rows))
                return
            self.html_response(render_admin_login())
            return
        if route == "/admin/review":
            if self.guard_admin():
                return
            row = get_submission_by_id(params.get("id", [""])[0])
            if not row:
                self.redirect("/admin")
                return
            self.html_response(render_admin_review(row))
            return
        if route == "/admin/logout":
            drop_admin_cookie(self)
            self.redirect("/admin", clear_cookie=True)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        parsed = urlparse(self.path)
        route = parsed.path
        fields, files = parse_form(self)
        code = parse_qs(parsed.query).get("code", [""])[0]

        if route == "/start":
            if not public_access_on():
                self.html_response(render_public_closed(), status=HTTPStatus.SERVICE_UNAVAILABLE)
                return
            full_name = fields.get("full_name", "").strip()
            email = fields.get("email", "").strip()
            age = fields.get("age", "").strip()
            country = fields.get("country", "").strip()
            target_band = fields.get("target_band", "6.0").strip()
            if not all([full_name, email, age, country, target_band]):
                self.html_response(render_start("Please complete all candidate fields."), status=HTTPStatus.BAD_REQUEST)
                return
            candidate_code = generate_candidate_code()
            timestamp = now_text()
            connection = db()
            cursor = connection.cursor()
            cursor.execute(
                "INSERT INTO candidates (candidate_code, full_name, email, age, country, target_band, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (candidate_code, full_name, email, int(age), country, float(target_band), timestamp),
            )
            candidate_id = cursor.lastrowid
            cursor.execute(
                """
                INSERT INTO submissions (
                    candidate_id, status, current_step, listening_answers, listening_correct, listening_band,
                    reading_answers, reading_correct, reading_band, writing_task_1, writing_task_2,
                    speaking_part_1_text, speaking_part_2_text, speaking_part_3_text,
                    speaking_part_1_audio, speaking_part_2_audio, speaking_part_3_audio,
                    candidate_notes, writing_band, speaking_band, overall_band, position_label,
                    examiner_feedback, created_at, updated_at, submitted_at, scored_at
                ) VALUES (?, 'in_progress', 'listening', '{}', NULL, NULL, '{}', NULL, NULL, '', '', '', '', '', '', '', '', '', NULL, NULL, NULL, '', '', ?, ?, '', '')
                """,
                (candidate_id, timestamp, timestamp),
            )
            connection.commit()
            connection.close()
            export_reports()
            self.redirect(f"/exam/listening?code={candidate_code}")
            return

        if route == "/exam/listening":
            if self.guard_public():
                return
            row = get_submission_by_code(code)
            if not row:
                self.redirect("/start")
                return
            answers = {item["id"]: fields.get(item["id"], "").strip() for item in LISTENING_ITEMS}
            if any(not value for value in answers.values()):
                self.html_response(render_listening(row, "Please answer every listening question."), status=HTTPStatus.BAD_REQUEST)
                return
            correct = score_short_answers(LISTENING_ITEMS, answers)
            band = convert_correct_to_band(correct, len(LISTENING_ITEMS))
            connection = db()
            connection.execute(
                "UPDATE submissions SET listening_answers = ?, listening_correct = ?, listening_band = ?, current_step = 'reading', updated_at = ? WHERE id = ?",
                (json.dumps(answers), correct, band, now_text(), row["id"]),
            )
            connection.commit()
            connection.close()
            export_reports()
            self.redirect(f"/exam/reading?code={code}")
            return

        if route == "/exam/reading":
            if self.guard_public():
                return
            row = get_submission_by_code(code)
            if not row:
                self.redirect("/start")
                return
            answers = {item["id"]: fields.get(item["id"], "").strip() for item in READING_ITEMS}
            if any(not value for value in answers.values()):
                self.html_response(render_reading(row, "Please answer every reading question."), status=HTTPStatus.BAD_REQUEST)
                return
            correct = score_short_answers(READING_ITEMS, answers)
            band = convert_correct_to_band(correct, len(READING_ITEMS))
            connection = db()
            connection.execute(
                "UPDATE submissions SET reading_answers = ?, reading_correct = ?, reading_band = ?, current_step = 'writing', updated_at = ? WHERE id = ?",
                (json.dumps(answers), correct, band, now_text(), row["id"]),
            )
            connection.commit()
            connection.close()
            export_reports()
            self.redirect(f"/exam/writing?code={code}")
            return

        if route == "/exam/writing":
            if self.guard_public():
                return
            row = get_submission_by_code(code)
            if not row:
                self.redirect("/start")
                return
            task_1 = fields.get("writing_task_1", "").strip()
            task_2 = fields.get("writing_task_2", "").strip()
            if not task_1 or not task_2:
                self.html_response(render_writing(row, "Please complete both writing tasks."), status=HTTPStatus.BAD_REQUEST)
                return
            connection = db()
            connection.execute(
                "UPDATE submissions SET writing_task_1 = ?, writing_task_2 = ?, current_step = 'speaking', updated_at = ? WHERE id = ?",
                (task_1, task_2, now_text(), row["id"]),
            )
            connection.commit()
            connection.close()
            export_reports()
            self.redirect(f"/exam/speaking?code={code}")
            return

        if route == "/exam/speaking":
            if self.guard_public():
                return
            row = get_submission_by_code(code)
            if not row:
                self.redirect("/start")
                return
            part_1 = fields.get("speaking_part_1_text", "").strip()
            part_2 = fields.get("speaking_part_2_text", "").strip()
            part_3 = fields.get("speaking_part_3_text", "").strip()
            notes = fields.get("candidate_notes", "").strip()
            if not part_1 or not part_2 or not part_3:
                self.html_response(render_speaking(row, "Please complete all speaking parts."), status=HTTPStatus.BAD_REQUEST)
                return
            audio_1 = save_audio_file(code, "part1", files.get("speaking_part_1_audio"))
            audio_2 = save_audio_file(code, "part2", files.get("speaking_part_2_audio"))
            audio_3 = save_audio_file(code, "part3", files.get("speaking_part_3_audio"))
            connection = db()
            connection.execute(
                """
                UPDATE submissions
                SET speaking_part_1_text = ?, speaking_part_2_text = ?, speaking_part_3_text = ?,
                    speaking_part_1_audio = CASE WHEN ? <> '' THEN ? ELSE speaking_part_1_audio END,
                    speaking_part_2_audio = CASE WHEN ? <> '' THEN ? ELSE speaking_part_2_audio END,
                    speaking_part_3_audio = CASE WHEN ? <> '' THEN ? ELSE speaking_part_3_audio END,
                    candidate_notes = ?, status = 'submitted', current_step = 'done', updated_at = ?, submitted_at = ?
                WHERE id = ?
                """,
                (part_1, part_2, part_3, audio_1, audio_1, audio_2, audio_2, audio_3, audio_3, notes, now_text(), now_text(), row["id"]),
            )
            connection.commit()
            connection.close()
            export_reports()
            self.redirect(f"/submitted?code={code}")
            return

        if route == "/admin/login":
            if fields.get("password", "") != ADMIN_PASSWORD:
                self.html_response(render_admin_login("Incorrect password."), status=HTTPStatus.FORBIDDEN)
                return
            token = make_admin_cookie()
            self.redirect("/admin", cookie=f"{ADMIN_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax")
            return

        if route == "/admin/toggle-public":
            if self.guard_admin():
                return
            set_setting("public_access", "off" if fields.get("next_state", "on").strip().lower() == "off" else "on")
            self.redirect("/admin")
            return

        if route == "/admin/score":
            if self.guard_admin():
                return
            row = get_submission_by_id(fields.get("submission_id", ""))
            if not row:
                self.redirect("/admin")
                return
            writing_band = float(fields.get("writing_band", "6.0"))
            speaking_band = float(fields.get("speaking_band", "6.0"))
            feedback = fields.get("examiner_feedback", "").strip()
            overall = round_overall((row["listening_band"] + row["reading_band"] + writing_band + speaking_band) / 4)
            label = position_label(overall)
            connection = db()
            connection.execute(
                "UPDATE submissions SET writing_band = ?, speaking_band = ?, overall_band = ?, position_label = ?, examiner_feedback = ?, status = 'published', scored_at = ?, updated_at = ? WHERE id = ?",
                (writing_band, speaking_band, overall, label, feedback, now_text(), now_text(), row["id"]),
            )
            connection.commit()
            connection.close()
            export_reports()
            self.redirect(f"/admin/review?id={row['id']}")
            return

        self.send_error(HTTPStatus.NOT_FOUND)


def main():
    os.makedirs(STATIC_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    os.makedirs(EXPORTS_DIR, exist_ok=True)
    init_db()
    export_reports()
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"BandForge running at http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
