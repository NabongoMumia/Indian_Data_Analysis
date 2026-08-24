import os
import io
import json
import psycopg2
from psycopg2.extras import RealDictCursor
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv

import openpyxl
from pypdf import PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "fallback-secret-key-change-me")

SUBJECTS = ['English', 'Maths', 'Chemistry', 'Biology', 'ICT', 'Physics', 'Business', 'Geography', 'Islamiyat']


def get_db_connection():
    db_url = os.getenv("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL environment variable is missing.")
    
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
        
    return psycopg2.connect(db_url, cursor_factory=RealDictCursor, sslmode='require')


def init_db():
    conn = None
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        
        cur.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username VARCHAR(100) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL
            );
        ''')
        
        cur.execute('''
            CREATE TABLE IF NOT EXISTS students (
                adm_no VARCHAR(50) PRIMARY KEY,
                student_name VARCHAR(150) NOT NULL,
                class_name VARCHAR(50) NOT NULL,
                assessment_term VARCHAR(100),
                opening_date VARCHAR(20),
                closing_date VARCHAR(20),
                class_teacher VARCHAR(100),
                principal_name VARCHAR(100),
                english FLOAT DEFAULT 0,
                maths FLOAT DEFAULT 0,
                chemistry FLOAT DEFAULT 0,
                biology FLOAT DEFAULT 0,
                ict FLOAT DEFAULT 0,
                physics FLOAT DEFAULT 0,
                business FLOAT DEFAULT 0,
                geography FLOAT DEFAULT 0,
                islamiyat FLOAT DEFAULT 0,
                teachers JSONB DEFAULT '{}'::jsonb
            );
        ''')
        conn.commit()
        cur.close()
    except Exception as e:
        print(f"Database initialization error: {e}")
    finally:
        if conn:
            conn.close()


with app.app_context():
    init_db()


@app.route("/")
def index():
    if "username" in session:
        return redirect(url_for("dashboard"))
    return render_template("login.html")

@app.route("/login", methods=["POST"])
def login():
    action = request.form.get("action")
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "").strip()

    if not username or not password:
        flash("Username and password are required.", "danger")
        return redirect(url_for("index"))

    conn = get_db_connection()
    cur = conn.cursor()

    try:
        if action == "register":
            cur.execute("SELECT id FROM users WHERE username = %s;", (username,))
            if cur.fetchone():
                flash("Username already exists.", "warning")
            else:
                hashed_pw = generate_password_hash(password)
                cur.execute("INSERT INTO users (username, password) VALUES (%s, %s);", (username, hashed_pw))
                conn.commit()
                session["username"] = username
                flash("Account registered successfully!", "success")
                return redirect(url_for("dashboard"))

        elif action == "login":
            cur.execute("SELECT * FROM users WHERE username = %s;", (username,))
            user = cur.fetchone()
            
            if user and check_password_hash(user["password"], password):
                session["username"] = username
                flash("Logged in successfully!", "success")
                return redirect(url_for("dashboard"))
            else:
                flash("Invalid credentials.", "danger")

    except Exception as e:
        conn.rollback()
        print(f"Login/Register Error: {e}")
        flash(f"An unexpected error occurred: {str(e)}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(url_for("index"))


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("index"))


@app.route("/dashboard")
def dashboard():
    if "username" not in session:
        return redirect(url_for("index"))

    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT DISTINCT class_name FROM students WHERE class_name IS NOT NULL AND class_name != '';")
        classes = [row["class_name"] for row in cur.fetchall()]
    finally:
        cur.close()
        conn.close()

    return render_template("dashboard.html", username=session["username"], subjects=SUBJECTS, classes=classes)


@app.route("/api/student/<adm_no>")
def get_student(adm_no):
    if "username" not in session:
        return jsonify({"found": False, "error": "Unauthorized"}), 401

    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT * FROM students WHERE adm_no = %s;", (adm_no,))
        student = cur.fetchone()
    finally:
        cur.close()
        conn.close()

    if student:
        student["found"] = True
        return jsonify(dict(student))
    return jsonify({"found": False})


@app.route("/save_student", methods=["POST"])
def save_student():
    if "username" not in session:
        return redirect(url_for("index"))

    adm_no = request.form.get("adm_no")
    student_name = request.form.get("student_name")
    class_name = request.form.get("class_name")
    assessment_term = request.form.get("assessment_term")
    opening_date = request.form.get("opening_date")
    closing_date = request.form.get("closing_date")
    class_teacher = request.form.get("class_teacher")
    principal_name = request.form.get("principal_name")

    teachers = {}
    scores = {}
    for sub in SUBJECTS:
        score_val = request.form.get(f"score_{sub}", 0)
        try:
            scores[sub.lower()] = float(score_val)
        except ValueError:
            scores[sub.lower()] = 0.0

        teachers[sub] = request.form.get(f"teacher_{sub}", "")

    conn = get_db_connection()
    cur = conn.cursor()

    try:
        cur.execute("""
            INSERT INTO students (
                adm_no, student_name, class_name, assessment_term, opening_date, closing_date,
                class_teacher, principal_name, english, maths, chemistry, biology, ict,
                physics, business, geography, islamiyat, teachers
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (adm_no) DO UPDATE SET
                student_name = EXCLUDED.student_name,
                class_name = EXCLUDED.class_name,
                assessment_term = EXCLUDED.assessment_term,
                opening_date = EXCLUDED.opening_date,
                closing_date = EXCLUDED.closing_date,
                class_teacher = EXCLUDED.class_teacher,
                principal_name = EXCLUDED.principal_name,
                english = EXCLUDED.english,
                maths = EXCLUDED.maths,
                chemistry = EXCLUDED.chemistry,
                biology = EXCLUDED.biology,
                ict = EXCLUDED.ict,
                physics = EXCLUDED.physics,
                business = EXCLUDED.business,
                geography = EXCLUDED.geography,
                islamiyat = EXCLUDED.islamiyat,
                teachers = EXCLUDED.teachers;
        """, (
            adm_no, student_name, class_name, assessment_term, opening_date, closing_date,
            class_teacher, principal_name, scores.get("english", 0), scores.get("maths", 0),
            scores.get("chemistry", 0), scores.get("biology", 0), scores.get("ict", 0),
            scores.get("physics", 0), scores.get("business", 0), scores.get("geography", 0),
            scores.get("islamiyat", 0), json.dumps(teachers)
        ))
        conn.commit()
    finally:
        cur.close()
        conn.close()

    flash(f"Record for {student_name} saved successfully!", "success")
    return redirect(url_for("dashboard"))


def create_student_pdf_buffer(student):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("<b>SKY SCHOOLS ACADEMIC REPORT CARD</b>", styles['Title']))
    elements.append(Spacer(1, 10))

    meta_data = [
        [f"Student Name: {student['student_name']}", f"Adm No: {student['adm_no']}"],
        [f"Class: {student['class_name']}", f"Term: {student['assessment_term']}"],
        [f"Opening Date: {student['opening_date']}", f"Closing Date: {student['closing_date']}"],
        [f"Class Teacher: {student.get('class_teacher', 'N/A')}", f"Principal: {student.get('principal_name', 'N/A')}"]
    ]
    t_meta = Table(meta_data, colWidths=[270, 270])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('PADDING', (0,0), (-1,-1), 6)
    ]))
    elements.append(t_meta)
    elements.append(Spacer(1, 15))

    score_data = [["Subject", "Score", "Subject Teacher"]]
    teachers = student.get("teachers") or {}
    if isinstance(teachers, str):
        try:
            teachers = json.loads(teachers)
        except Exception:
            teachers = {}

    total_score = 0
    count = 0

    for sub in SUBJECTS:
        score = student.get(sub.lower(), 0.0)
        total_score += score
        count += 1
        t_name = teachers.get(sub, "")
        score_data.append([sub, f"{score:.1f}", t_name])

    avg_score = total_score / count if count > 0 else 0
    score_data.append(["TOTAL / AVERAGE", f"Total: {total_score:.1f}", f"Average: {avg_score:.1f}%"])

    t_scores = Table(score_data, colWidths=[180, 160, 200])
    t_scores.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.navy),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 6)
    ]))
    elements.append(t_scores)

    doc.build(elements)
    buffer.seek(0)
    return buffer


@app.route("/generate_pdf/<adm_no>")
def generate_pdf(adm_no):
    if "username" not in session:
        return redirect(url_for("index"))

    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT * FROM students WHERE adm_no = %s;", (adm_no,))
        student = cur.fetchone()
    finally:
        cur.close()
        conn.close()

    if not student:
        flash("Student record not found.", "danger")
        return redirect(url_for("dashboard"))

    pdf_buffer = create_student_pdf_buffer(student)
    return send_file(pdf_buffer, as_attachment=True, download_name=f"Report_{adm_no}.pdf", mimetype="application/pdf")


@app.route("/download_class_pdf", methods=["POST"])
def download_class_pdf():
    if "username" not in session:
        return redirect(url_for("index"))

    class_name = request.form.get("class_name")
    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT * FROM students WHERE class_name = %s ORDER BY student_name ASC;", (class_name,))
        students = cur.fetchall()
    finally:
        cur.close()
        conn.close()

    if not students:
        flash("No students found for this class.", "warning")
        return redirect(url_for("dashboard"))

    merger = PdfWriter()
    for st in students:
        st_buffer = create_student_pdf_buffer(st)
        merger.append(st_buffer)

    output_buffer = io.BytesIO()
    merger.write(output_buffer)
    merger.close()
    output_buffer.seek(0)

    return send_file(output_buffer, as_attachment=True, download_name=f"Class_Reports_{class_name}.pdf", mimetype="application/pdf")


@app.route("/export_excel", methods=["POST"])
def export_excel():
    if "username" not in session:
        return redirect(url_for("index"))

    class_name = request.form.get("class_name")
    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute("SELECT * FROM students WHERE class_name = %s ORDER BY student_name ASC;", (class_name,))
        students = cur.fetchall()
    finally:
        cur.close()
        conn.close()

    if not students:
        flash("No students found to export.", "warning")
        return redirect(url_for("dashboard"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Class {class_name}"

    headers = ["Adm No", "Student Name", "Class", "Term"] + SUBJECTS + ["Total Score", "Average (%)"]
    ws.append(headers)

    for st in students:
        row = [st["adm_no"], st["student_name"], st["class_name"], st.get("assessment_term", "")]
        total = 0
        for sub in SUBJECTS:
            val = st.get(sub.lower(), 0.0)
            row.append(val)
            total += val
        avg = total / len(SUBJECTS)
        row.extend([total, round(avg, 2)])
        ws.append(row)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name=f"Marklist_{class_name}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(debug=True)