import os
import sys
from functools import wraps
from pathlib import Path

workspace_packages = (Path(__file__).resolve().parent / ".python_packages").resolve()
if workspace_packages.exists() and str(workspace_packages) not in sys.path:
    sys.path.insert(0, str(workspace_packages))

from flask import Flask, redirect, request, send_from_directory, session, url_for

import app as core


web_app = Flask(__name__, static_folder="static", static_url_path="/static")
web_app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "change-me-on-pythonanywhere")


def html_response(body, status=200):
    return body, status, {"Content-Type": "text/html; charset=utf-8"}


def admin_required(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if session.get("admin_logged_in"):
            return view_func(*args, **kwargs)
        return redirect(url_for("admin"))

    return wrapper


def public_guard():
    if core.public_access_on():
        return None
    return html_response(core.render_public_closed(), 503)


def form_file_info(field_name):
    uploaded = request.files.get(field_name)
    if not uploaded or not uploaded.filename:
        return None
    return {
        "filename": uploaded.filename,
        "content_type": uploaded.mimetype,
        "data": uploaded.read(),
    }


def latest_admin_rows():
    connection = core.db()
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
    return rows


@web_app.get("/")
def home():
    return html_response(core.render_home())


@web_app.route("/start", methods=["GET", "POST"])
def start():
    blocked = public_guard()
    if blocked:
        return blocked
    if request.method == "GET":
        return html_response(core.render_start())

    full_name = request.form.get("full_name", "").strip()
    email = request.form.get("email", "").strip()
    age = request.form.get("age", "").strip()
    country = request.form.get("country", "").strip()
    target_band = request.form.get("target_band", "6.0").strip()
    if not all([full_name, email, age, country, target_band]):
        return html_response(core.render_start("Please complete all candidate fields."), 400)

    candidate_code = core.generate_candidate_code()
    timestamp = core.now_text()
    connection = core.db()
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
    core.export_reports()
    return redirect(url_for("exam_listening", code=candidate_code))


@web_app.route("/exam/listening", methods=["GET", "POST"])
def exam_listening():
    blocked = public_guard()
    if blocked:
        return blocked
    code = request.args.get("code", "")
    row = core.get_submission_by_code(code)
    if not row:
        return redirect(url_for("start"))
    if request.method == "GET":
        return html_response(core.render_listening(row))

    answers = {item["id"]: request.form.get(item["id"], "").strip() for item in core.LISTENING_ITEMS}
    if any(not value for value in answers.values()):
        return html_response(core.render_listening(row, "Please answer every listening question."), 400)
    correct = core.score_short_answers(core.LISTENING_ITEMS, answers)
    band = core.convert_correct_to_band(correct, len(core.LISTENING_ITEMS))
    connection = core.db()
    connection.execute(
        "UPDATE submissions SET listening_answers = ?, listening_correct = ?, listening_band = ?, current_step = 'reading', updated_at = ? WHERE id = ?",
        (core.json.dumps(answers), correct, band, core.now_text(), row["id"]),
    )
    connection.commit()
    connection.close()
    core.export_reports()
    return redirect(url_for("exam_reading", code=code))


@web_app.route("/exam/reading", methods=["GET", "POST"])
def exam_reading():
    blocked = public_guard()
    if blocked:
        return blocked
    code = request.args.get("code", "")
    row = core.get_submission_by_code(code)
    if not row:
        return redirect(url_for("start"))
    if request.method == "GET":
        return html_response(core.render_reading(row))

    answers = {item["id"]: request.form.get(item["id"], "").strip() for item in core.READING_ITEMS}
    if any(not value for value in answers.values()):
        return html_response(core.render_reading(row, "Please answer every reading question."), 400)
    correct = core.score_short_answers(core.READING_ITEMS, answers)
    band = core.convert_correct_to_band(correct, len(core.READING_ITEMS))
    connection = core.db()
    connection.execute(
        "UPDATE submissions SET reading_answers = ?, reading_correct = ?, reading_band = ?, current_step = 'writing', updated_at = ? WHERE id = ?",
        (core.json.dumps(answers), correct, band, core.now_text(), row["id"]),
    )
    connection.commit()
    connection.close()
    core.export_reports()
    return redirect(url_for("exam_writing", code=code))


@web_app.route("/exam/writing", methods=["GET", "POST"])
def exam_writing():
    blocked = public_guard()
    if blocked:
        return blocked
    code = request.args.get("code", "")
    row = core.get_submission_by_code(code)
    if not row:
        return redirect(url_for("start"))
    if request.method == "GET":
        return html_response(core.render_writing(row))

    task_1 = request.form.get("writing_task_1", "").strip()
    task_2 = request.form.get("writing_task_2", "").strip()
    if not task_1 or not task_2:
        return html_response(core.render_writing(row, "Please complete both writing tasks."), 400)
    connection = core.db()
    connection.execute(
        "UPDATE submissions SET writing_task_1 = ?, writing_task_2 = ?, current_step = 'speaking', updated_at = ? WHERE id = ?",
        (task_1, task_2, core.now_text(), row["id"]),
    )
    connection.commit()
    connection.close()
    core.export_reports()
    return redirect(url_for("exam_speaking", code=code))


@web_app.route("/exam/speaking", methods=["GET", "POST"])
def exam_speaking():
    blocked = public_guard()
    if blocked:
        return blocked
    code = request.args.get("code", "")
    row = core.get_submission_by_code(code)
    if not row:
        return redirect(url_for("start"))
    if request.method == "GET":
        return html_response(core.render_speaking(row))

    part_1 = request.form.get("speaking_part_1_text", "").strip()
    part_2 = request.form.get("speaking_part_2_text", "").strip()
    part_3 = request.form.get("speaking_part_3_text", "").strip()
    notes = request.form.get("candidate_notes", "").strip()
    if not part_1 or not part_2 or not part_3:
        return html_response(core.render_speaking(row, "Please complete all speaking parts."), 400)

    audio_1 = core.save_audio_file(code, "part1", form_file_info("speaking_part_1_audio"))
    audio_2 = core.save_audio_file(code, "part2", form_file_info("speaking_part_2_audio"))
    audio_3 = core.save_audio_file(code, "part3", form_file_info("speaking_part_3_audio"))
    connection = core.db()
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
        (part_1, part_2, part_3, audio_1, audio_1, audio_2, audio_2, audio_3, audio_3, notes, core.now_text(), core.now_text(), row["id"]),
    )
    connection.commit()
    connection.close()
    core.export_reports()
    return redirect(url_for("submitted", code=code))


@web_app.get("/submitted")
def submitted():
    row = core.get_submission_by_code(request.args.get("code", ""))
    if not row:
        return redirect(url_for("home"))
    return html_response(core.render_submitted(row))


@web_app.get("/results")
def results():
    code = request.args.get("code", "").strip()
    if not code:
        return html_response(core.render_results_lookup())
    row = core.get_submission_by_code(code)
    if not row:
        return html_response(core.render_results_lookup("No candidate found for that code."), 404)
    if row["status"] != "published":
        return html_response(core.render_results_lookup("This candidate exists, but the final result has not been published yet."))
    return html_response(core.render_results_lookup(result_html=core.published_result_card(row)))


@web_app.route("/admin", methods=["GET"])
def admin():
    if session.get("admin_logged_in"):
        return html_response(core.render_admin_dashboard(latest_admin_rows()))
    return html_response(core.render_admin_login())


@web_app.post("/admin/login")
def admin_login():
    if request.form.get("password", "") != core.ADMIN_PASSWORD:
        return html_response(core.render_admin_login("Incorrect password."), 403)
    session["admin_logged_in"] = True
    return redirect(url_for("admin"))


@web_app.get("/admin/logout")
def admin_logout():
    session.pop("admin_logged_in", None)
    return redirect(url_for("admin"))


@web_app.post("/admin/toggle-public")
@admin_required
def admin_toggle_public():
    next_state = request.form.get("next_state", "on").strip().lower()
    core.set_setting("public_access", "off" if next_state == "off" else "on")
    return redirect(url_for("admin"))


@web_app.get("/admin/review")
@admin_required
def admin_review():
    row = core.get_submission_by_id(request.args.get("id", ""))
    if not row:
        return redirect(url_for("admin"))
    return html_response(core.render_admin_review(row))


@web_app.post("/admin/score")
@admin_required
def admin_score():
    row = core.get_submission_by_id(request.form.get("submission_id", ""))
    if not row:
        return redirect(url_for("admin"))
    writing_band = float(request.form.get("writing_band", "6.0"))
    speaking_band = float(request.form.get("speaking_band", "6.0"))
    feedback = request.form.get("examiner_feedback", "").strip()
    overall = core.round_overall((row["listening_band"] + row["reading_band"] + writing_band + speaking_band) / 4)
    label = core.position_label(overall)
    connection = core.db()
    connection.execute(
        "UPDATE submissions SET writing_band = ?, speaking_band = ?, overall_band = ?, position_label = ?, examiner_feedback = ?, status = 'published', scored_at = ?, updated_at = ? WHERE id = ?",
        (writing_band, speaking_band, overall, label, feedback, core.now_text(), core.now_text(), row["id"]),
    )
    connection.commit()
    connection.close()
    core.export_reports()
    return redirect(url_for("admin_review", id=row["id"]))


@web_app.get("/uploads/<path:filename>")
def uploaded_file(filename):
    return send_from_directory(core.UPLOADS_DIR, filename)


@web_app.get("/exports/<path:filename>")
@admin_required
def exported_file(filename):
    return send_from_directory(core.EXPORTS_DIR, filename, mimetype="application/vnd.ms-excel", as_attachment=True)


def prepare_runtime():
    os.makedirs(core.STATIC_DIR, exist_ok=True)
    os.makedirs(core.DATA_DIR, exist_ok=True)
    os.makedirs(core.UPLOADS_DIR, exist_ok=True)
    os.makedirs(core.EXPORTS_DIR, exist_ok=True)
    core.init_db()
    core.export_reports()


prepare_runtime()
app = web_app


if __name__ == "__main__":
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", "8000")), debug=True)
