import os
import sqlite3
from flask import Flask, request, redirect, session, render_template_string
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook
from pathlib import Path

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")

DB_PATH = os.environ.get("DB_PATH", "users.db")
EXCEL_PATH = Path("words.xlsx")  # ← あなたの単語Excelに合わせてOK

# -------------------------
# DB
# -------------------------

def get_db():
    return sqlite3.connect(DB_PATH)

def init_db():
    with get_db() as db:
        db.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'student',
            is_active INTEGER NOT NULL DEFAULT 0
        )
        """)

def ensure_admin():
    with get_db() as db:
        cur = db.execute("SELECT * FROM users WHERE role='admin'")
        if cur.fetchone() is None:
            db.execute("""
                INSERT INTO users (username, password_hash, role, is_active)
                VALUES (?, ?, 'admin', 1)
            """, ("teacher", generate_password_hash("changeme")))

# -------------------------
# HTML（templates無し）
# -------------------------

LOGIN_HTML = """
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>ログイン</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
body{font-family:sans-serif;background:#f5f5f5;
display:flex;justify-content:center;align-items:center;height:100vh}
.box{background:#fff;padding:24px;width:320px;border-radius:8px}
input,button{width:100%;padding:10px;margin-top:10px}
button{background:#007bff;color:#fff;border:none}
a{display:block;text-align:center;margin-top:10px}
</style>
</head>
<body>
<div class="box">
<h2>ログイン</h2>
<form method="post">
<input name="username" placeholder="ID" required>
<input type="password" name="password" placeholder="パスワード" required>
<button type="submit">ログイン</button>
</form>
<a href="/register">新規登録</a>
</div>
</body>
</html>
"""

REGISTER_HTML = """
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>新規登録</title>
<style>
body{font-family:sans-serif;background:#f5f5f5;
display:flex;justify-content:center;align-items:center;height:100vh}
.box{background:#fff;padding:24px;width:320px;border-radius:8px}
input,button{width:100%;padding:10px;margin-top:10px}
</style>
</head>
<body>
<div class="box">
<h2>新規登録</h2>
<form method="post">
<input name="username" placeholder="名前(ID)" required>
<input type="password" name="password" placeholder="パスワード" required>
<button type="submit">登録</button>
</form>
<a href="/login">戻る</a>
</div>
</body>
</html>
"""

# ※ ここにあなたの HTML単語テストをそのまま貼ってOK
INDEX_HTML = """
<!doctype html>
<html>
<head><meta charset="utf-8"><title>単語テスト</title></head>
<body>
<h2>単語テスト</h2>

<form action="/logout">
<button>ログアウト</button>
</form>

<ul>
{% for s in sheets %}
<li>{{ s }}</li>
{% endfor %}
</ul>

</body>
</html>
"""

ADMIN_HTML = """
<!doctype html>
<html>
<head><meta charset="utf-8"><title>管理者</title></head>
<body>
<h2>管理者画面</h2>
<table border="1">
<tr><th>ID</th><th>名前</th><th>状態</th><th>操作</th></tr>
{% for u in users %}
<tr>
<td>{{u[0]}}</td>
<td>{{u[1]}}</td>
<td>{{"OK" if u[2] else "承認待ち"}}</td>
<td>
{% if not u[2] %}
<a href="/approve/{{u[0]}}">承認</a>
{% endif %}
<a href="/reset/{{u[0]}}">PWリセット</a>
<a href="/delete/{{u[0]}}">削除</a>
</td>
</tr>
{% endfor %}
</table>
<br>
<a href="/logout">ログアウト</a>
</body>
</html>
"""

# -------------------------
# ログイン制御
# -------------------------

@app.before_request
def require_login():
    open_paths = ["/login", "/register", "/static"]
    if any(request.path.startswith(p) for p in open_paths):
        return
    if not session.get("user_id"):
        return redirect("/login")

# -------------------------
# Routes
# -------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form["username"]
        p = request.form["password"]

        with get_db() as db:
            cur = db.execute("SELECT * FROM users WHERE username=?", (u,))
            user = cur.fetchone()

        if not user or not check_password_hash(user[2], p):
            return "ログイン失敗"

        if not user[4]:
            return "承認待ちです"

        session["user_id"] = user[0]
        session["role"] = user[3]

        return redirect("/admin" if user[3]=="admin" else "/")

    return render_template_string(LOGIN_HTML)

@app.route("/register", methods=["GET","POST"])
def register():
    if request.method=="POST":
        with get_db() as db:
            db.execute(
                "INSERT INTO users (username,password_hash) VALUES (?,?)",
                (request.form["username"],
                 generate_password_hash(request.form["password"]))
            )
        return "登録しました。承認待ちです。<br><a href='/login'>戻る</a>"
    return render_template_string(REGISTER_HTML)

@app.route("/")
def index():
    wb = load_workbook(str(EXCEL_PATH), read_only=True)
    return render_template_string(INDEX_HTML, sheets=wb.sheetnames)

@app.route("/admin")
def admin():
    if session.get("role")!="admin":
        return redirect("/")
    with get_db() as db:
        users = db.execute(
            "SELECT id,username,is_active FROM users WHERE role='student'"
        ).fetchall()
    return render_template_string(ADMIN_HTML, users=users)

@app.route("/approve/<int:uid>")
def approve(uid):
    with get_db() as db:
        db.execute("UPDATE users SET is_active=1 WHERE id=?", (uid,))
    return redirect("/admin")

@app.route("/reset/<int:uid>")
def reset(uid):
    with get_db() as db:
        db.execute(
            "UPDATE users SET password_hash=? WHERE id=?",
            (generate_password_hash("1234"), uid)
        )
    return redirect("/admin")

@app.route("/delete/<int:uid>")
def delete(uid):
    with get_db() as db:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
    return redirect("/admin")

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# -------------------------
# 起動時
# -------------------------

with app.app_context():
    init_db()
    ensure_admin()
