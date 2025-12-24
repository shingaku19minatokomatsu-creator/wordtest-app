import os
import sqlite3
from flask import Flask, request, redirect, session, render_template_string
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook
from pathlib import Path

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")

DB_PATH = os.environ.get("DB_PATH", "users.db")
EXCEL_PATH = Path("英単語テスト.xlsx")  # ← あなたの単語Excelに合わせてOK

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
            """, ("minato", generate_password_hash("3710")))

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

# ===== HTML ======
INDEX_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>単語テスト</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<style>
body {
  font-family: Arial, sans-serif;
  margin: 0 auto;
  padding: 6mm;
  font-size: 14px;
  max-width: none;
  touch-action: pan-y pinch-zoom;
}

html, body {
  overscroll-behavior: none;
}

@media print {
  @page {
    size: A4 landscape;
    margin: 15mm;
  }

  body {
    width: 297mm;
    height: 210mm;
    padding: 0;
  }
}


h2 {
    font-size: 26px;
    margin-bottom: 10px;
}

label {
    display: block;
    font-size: 18px;
    margin-bottom: 4px;
}

input, select, button {
    padding: 12px;
    font-size: 18px;
    width: 100%;
    box-sizing: border-box;
}

.row {
    margin: 15px 0;
}

button {
    background-color: #007bff;
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 20px;
    padding: 14px;
    cursor: pointer;
}

button:hover {
    background-color: #0056c7;
}

.note {
    color: #666;
    font-size: 15px;
    margin-bottom: 10px;
}

/* スマホ用 */
@media (max-width: 600px) {
    body {
        padding: 14px;
        font-size: 17px;
    }
    input, select, button {
        font-size: 18px;
        padding: 14px;
    }
    h2 {
        font-size: 24px;
    }
}
</style>
</head>

<body>

<div style="text-align:right;margin-bottom:10px;">
  <a href="/logout">ログアウト</a>
</div>

<form>
  <div class="row">
    <label>単語帳（シート）</label>
    <select id="sheet">
      {% for s in sheets %}
      <option value="{{s}}">{{s}}</option>
      {% endfor %}
    </select>
  </div>

  <div class="row">
    <label>開始番号</label>
    <input id="start" required>
  </div>

  <div class="row">
    <label>終了番号</label>
    <input id="end" required>
  </div>

  <div class="row">
    <button type="button" onclick="doPdf()">PDF出力</button>
  </div>

  <div class="row">
    <button type="button" onclick="doHtml()">HTMLテスト</button>
  </div>
</form>

<script>
function getParams(){
  const sheet = document.getElementById('sheet').value;
  const start = document.getElementById('start').value;
  const end   = document.getElementById('end').value;
  if(!sheet || !start || !end){
    alert("シート・開始・終了番号が必要です");
    return null;
  }
  return {sheet, start, end};
}

async function doPdf(){
  const p = getParams();
  if(!p) return;

  const win = window.open("about:blank");
  const res = await fetch("/generate", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify(p)
  });

  if(!res.ok){
    win.close();
    alert(await res.text());
    return;
  }
  const data = await res.json();
  win.location.href = data.pdf_url;
}

async function doHtml(){
  const p = getParams();
  if(!p) return;

  const win = window.open("about:blank");
  const res = await fetch("/generate_html_test", {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify(p)
  });

  if(!res.ok){
    win.close();
    alert(await res.text());
    return;
  }
  const html = await res.text();
  win.document.open();
  win.document.write(html);
  win.document.close();
}
</script>



</body>
</html>
"""


# ①〜④ HTML版テスト機能 追加コード（PDF完全一致レイアウト版）
# 既存 app.py に追記する想定

# ①〜④ HTML版テスト機能 追加コード（PDF完全一致レイアウト版）
# 既存 app.py に追記する想定

HTML_TEST_TEMPLATE = """
<!doctype html>
<html>
<head>
<meta charset=\"utf-8\">
<title>単語テスト（HTML）</title>
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">

<style>
/* ===== 画面表示 ===== */
body {
  font-family: Arial, sans-serif;
  margin: 0;
  max-width: 100%;

  touch-action: pan-y pinch-zoom;   /* ★ここを追加 */
}
html, body {
  overscroll-behavior: none;
}


/* ===== 印刷時のみ A4 ===== */
@media print {

  @page {
    size: A4 landscape;
    margin: 15mm;
  }

  /* ===== 65%をCSS側で固定 ===== */
  #print-root {
    transform: scale(0.65);
    transform-origin: top left;

    /* 縮小分の横幅補正（重要） */
    width: calc(100% / 0.65);
  }

  body {
    margin: 0;
    overflow: hidden;
  }

  /* 操作UIは消す */
  button {
    display: none !important;
  }
}




/* ===== ヘッダ ===== */
.header {
  display: flex;
  justify-content: space-between;
  margin-bottom: 10mm;
}

/* ===== 2列（PDFと同じ） ===== */

.item {
  display: grid;
  grid-template-columns:
    44px                 /* 番号 */
    minmax(220px, 1fr)   /* 問題 */
    minmax(120px, 160px) /* 解答 */
    190px                /* canvas */
    44px
    minmax(220px, 1fr)
    minmax(120px, 160px)
    190px;

  height: 40px;
  align-items: center;
  font-size: 13px;
  box-sizing: border-box;
}


.answer {
  min-width: 0;
  font-weight: bold;
  color: red;
  opacity: 0.85;

  visibility: hidden;

  font-size: 11px;
  line-height: 1.2;

  white-space: normal;
  word-break: break-word;

  /* ★ここから追加 */
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

.answer.show {
  visibility: visible;
}


/* ===== canvas ===== */

canvas {
  display: block;
  background: #f2f2f2;
  border: 1px solid #ccc;

  touch-action: none;     /* ★ canvasだけロック */
  user-select: none;
  pointer-events: auto;
}

.item {
  user-select: none;      /* 文字選択防止だけ */
}

/* ★ item * は消す */

.small-text {
  font-size: 9px;
  line-height: 1.1;
}

/* ===== 印刷時 ===== */
@media print {
  button { display: none; }
}

</style>
</head>

<body>
<div id="print-root">


<div class=\"header\">
  <div>
    <h2>shingaku19minato test</h2>
    <div>words {{sheet}}（{{start}}～{{end}}）</div>
  </div>
    <div>
        name：
        <canvas width="160" height="28"></canvas><br>
        score：
        <canvas width="160" height="28"></canvas>
    </div>
</div>

<div style="margin-bottom:5mm">
<button onclick="toggleAll()">解答 表示／非表示</button>
<button onclick="setColor('black')">⚫ 黒</button>
<button onclick="setColor('red')">🔴 赤</button>
<button onclick="setMode('eraser')">🧽 消しゴム</button>
<button onclick="clearAll()">🗑 全消し</button>
<button onclick="window.print()">🖨 印刷</button>



</div>

    {% for i in range(20) %}
    {% set item  = items[i] %}
    {% set item2 = items[i+20] %}

    <div class="item">
        <!-- 左（1〜20） -->
        <div>{{item.no}}.</div>
        <div>{{item.q}}</div>
        <div class="answer" id="ans-{{item.no}}">{{item.a}}</div>
        <canvas width="180" height="36"></canvas>

        <!-- 右（21〜40） -->
        <div>{{item2.no}}.</div>
        <div>{{item2.q}}</div>
        <div class="answer" id="ans-{{item2.no}}">{{item2.a}}</div>
        <canvas width="180" height="36"></canvas>
    </div>

    {% endfor %}

<script>
let mode = "pen";
let color = "#000";

function setColor(c){
  color = (c === "red") ? "#d00" : "#000";
  mode = "pen";
}

function setMode(m){
  mode = m;
}

function clearAll(){
  document.querySelectorAll("canvas").forEach(c=>{
    c.getContext("2d").clearRect(0,0,c.width,c.height);
  });
}

function toggleAll(){
  document.querySelectorAll('.answer')
    .forEach(a => a.classList.toggle('show'));
}


document.querySelectorAll("canvas").forEach(c=>{
  const ratio = window.devicePixelRatio || 1;

  // ===== ① CSSサイズを保存（テンプレそのまま）=====
  const cssW = c.width;
  const cssH = c.height;

  // ===== ② 内部解像度だけ拡大 =====
  c.width  = cssW * ratio;
  c.height = cssH * ratio;

  // ===== ③ 見た目サイズは固定 =====
  c.style.width  = cssW + "px";
  c.style.height = cssH + "px";

  const ctx = c.getContext("2d");

  // ★ ここが最重要（座標系を元に戻す）
  ctx.scale(ratio, ratio);

  let drawing = false;

  ctx.lineWidth = 0.6;        // ← 今まで通りでOK
  ctx.lineCap = "round";
  ctx.lineJoin = "round";
  ctx.strokeStyle = color;

  function getPos(e){
    const rect = c.getBoundingClientRect();
    return {
      x: e.clientX - rect.left,
      y: e.clientY - rect.top
    };
  }

  c.addEventListener("touchstart", e=>{
    e.preventDefault();
  }, { passive: false });

  c.addEventListener("pointerdown", e=>{
    e.preventDefault();
    e.stopPropagation();

    drawing = true;
    c.setPointerCapture(e.pointerId);

    const p = getPos(e);
    ctx.beginPath();
    ctx.moveTo(p.x, p.y);
  });

  c.addEventListener("pointermove", e=>{
    if(!drawing) return;
    e.preventDefault();

    const p = getPos(e);

    if(mode === "eraser"){
      ctx.clearRect(p.x - 6, p.y - 6, 12, 12);
    }else{
      ctx.strokeStyle = color;
      ctx.lineTo(p.x, p.y);
      ctx.stroke();
    }
  });

  c.addEventListener("pointerup", e=>{
    drawing = false;
    c.releasePointerCapture(e.pointerId);
  });

  c.addEventListener("pointercancel", ()=>{
    drawing = false;
  });
});


document.querySelectorAll('.answer, .item > div:nth-child(2), .item > div:nth-child(6)')
  .forEach(el=>{
    if(el.textContent.length > 30){
      el.classList.add('small-text');
    }
  });

</script>





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
    open_paths = ["/login", "/register", "/static", "/favicon.ico"]
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

  
@app.route("/generate_html_test", methods=["POST"])
def generate_html_test():
    data = request.json
    sheet = data["sheet"]
    start = int(data["start"])
    end   = int(data["end"])

    items = []
    for i in range(start, end + 1):
        items.append({
            "no": i,
            "q": f"Question {i}",
            "a": f"Answer {i}"
        })

    # ★ 40問未満でも落ちないようにする（重要）
    while len(items) < 40:
        items.append({"no":"", "q":"", "a":""})

    return render_template_string(
        HTML_TEST_TEMPLATE,
        sheet=sheet,
        start=start,
        end=end,
        items=items
    )

  
@app.route("/generate", methods=["POST"])
def generate():
    data = request.json

    # 仮：とりあえず動作確認用
    return {
        "pdf_url": "/test.pdf"
    }


# -------------------------
# 起動時
# -------------------------

with app.app_context():
    init_db()
    ensure_admin()
