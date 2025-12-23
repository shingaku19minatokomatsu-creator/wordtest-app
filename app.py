# app.py
# A方式：1つの PDF に「問題ページ → 解答ページ」の2ページ構成で test.pdf を生成

import random
import uuid
from pathlib import Path
from tempfile import gettempdir
from flask import Flask, request, render_template_string, jsonify, send_file
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import os
from reportlab.pdfbase.pdfmetrics import stringWidth
from flask import request, abort
from reportlab.pdfbase.ttfonts import TTFont
from tempfile import gettempdir
from flask import session, redirect


app = Flask(__name__)

app.secret_key = "change-this-to-random-string"
      
# ====== 設定 ======
EXCEL_PATH = Path("英単語テスト.xlsx")

# 安定した PDF 保存フォルダ（Render対応）
try:
    TMPDIR = Path(gettempdir()) / "word_a_mode"
    TMPDIR.mkdir(parents=True, exist_ok=True)
except Exception:
    TMPDIR = Path("/tmp/word_a_mode")

# Render で mkdir が効かないケースに追加安全策
if not TMPDIR.exists():
    try:
        TMPDIR.mkdir(parents=True, exist_ok=True)
    except:
        pass

print("📁 PDF 保存先:", TMPDIR.absolute())

# ====== 日本語フォント（同梱） ======
FONT_PATH = Path("fonts/ipaexm.ttf")
DEFAULT_FONT = "IPAEX_M"

try:
    pdfmetrics.registerFont(TTFont(DEFAULT_FONT, str(FONT_PATH)))
except Exception as e:
    print("⚠ 日本語フォントの読み込みに失敗 → Helveticaに変更", e)
    DEFAULT_FONT = "Helvetica"


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

<h2>単語テスト</h2>
<div class="note">※「印刷用」を押すと test.pdf（問題→解答）が開きます。</div>

<form id="form" onsubmit="return doGenerate(event)">
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
    <button type="submit">印刷用</button>
  </div>
  
  <div class="row">
    <button type="button" onclick="doHtmlTest()">テスト</button>
  </div>


</form>

<script>
async function doGenerate(e){
  e.preventDefault();

  const sheet = document.getElementById('sheet').value;
  const start = document.getElementById('start').value;
  const end   = document.getElementById('end').value;

  if(!sheet || !start || !end){
    alert("シート・開始・終了番号が必要です。");
    return false;
  }

  const win = window.open("about:blank", "_blank");

  try {
    const res = await fetch("/generate", {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify({sheet, start, end})
    });

    if(!res.ok){
      const tx = await res.text();
      win.close();
      alert("エラー: " + tx);
      return false;
    }

    const data = await res.json();
    win.location.href = data.pdf_url;

  } catch(err){
    win.close();
    alert("通信エラー: " + err);
  }

  return false;
}


async function doHtmlTest(){
  const sheet = document.getElementById('sheet').value;
  const start = document.getElementById('start').value;
  const end   = document.getElementById('end').value;

  if(!sheet || !start || !end){
    alert("シート・開始・終了番号が必要です。");
    return;
  }

  const win = window.open("about:blank", "_blank");

  try {
    const res = await fetch("/generate_html_test", {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify({sheet, start, end})
    });

    if(!res.ok){
      const tx = await res.text();
      win.close();
      alert("エラー: " + tx);
      return;
    }

    const html = await res.text();
    win.document.open();
    win.document.write(html);
    win.document.close();

  } catch(err){
    win.close();
    alert("通信エラー: " + err);
  }
}
</script>


</body>
</html>
"""

LOGIN_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>ログイン</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body { font-family: sans-serif; max-width: 420px; margin: 60px auto; padding: 20px; }
input, button { width: 100%; padding: 14px; font-size: 18px; margin: 10px 0; }
button { background:#007bff; color:#fff; border:none; border-radius:6px; }
</style>
</head>
<body>
<h2>ログイン</h2>
<form method="post">
  <input name="username" placeholder="ID" required>
  <input name="password" type="password" placeholder="パスワード" required>
  <button>ログイン</button>
</form>
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
  @page { size: A4 landscape; margin: 15mm; }

  body {
    width: 297mm;
    height: 210mm;
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




@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form["username"]
        pw   = request.form["password"]

        # ★ 好きなIDとパスワードに変更
        if user == "minato" and pw == "3710":
            session["login"] = True
            return redirect("/")
        else:
            return "ログイン失敗", 401

    return render_template_string(LOGIN_HTML)


@app.before_request
def require_login():
    path = request.path

    if path.startswith("/login") or path.startswith("/static"):
        return

    if not session.get("login"):
        return redirect("/login")
    





def draw_text_fitted(c, text, font, base_x, base_y, max_width, max_height):
    if not text:
        return

    max_font  = 10
    min_font  = 3
    max_lines = 2

    if len(text) > 80:
        max_font = 7

    for size in range(max_font, min_font - 1, -1):
        line_gap = max(2, int(size * 0.3))

        lines = wrap_text(text, font, size, max_width)

        # ★ ここが最重要：先に行数で弾く
        if len(lines) > max_lines:
            continue

        total_h = len(lines) * size + (len(lines) - 1) * line_gap

        if total_h <= max_height:
            y = base_y
            c.setFont(font, size)
            for ln in lines:
                c.drawString(base_x, y, ln)
                y -= (size + line_gap)
            return

    # 最後の保険（強制切り）
    size = min_font
    c.setFont(font, size)
    lines = wrap_text(text, font, size, max_width)[:max_lines]
    y = base_y
    for ln in lines:
        c.drawString(base_x, y, ln)
        y -= (size + line_gap)


def draw_answer_fitted(c, text, font, base_x, base_y, max_width, max_height):
    if not text:
        return

    max_font  = 10
    min_font  = 3
    max_lines = 2

    if len(text) > 80:
        max_font = 7

    for size in range(max_font, min_font - 1, -1):
        line_gap = max(2, int(size * 0.3))

        lines = wrap_text(text, font, size, max_width)

        # ★ 先に行数オーバーならフォントを下げる
        if len(lines) > max_lines:
            continue

        total_h = len(lines) * size + (len(lines) - 1) * line_gap

        if total_h <= max_height:
            y = base_y
            c.setFont(font, size)
            for ln in lines:
                c.drawString(base_x, y, ln)
                y -= (size + line_gap)
            return

    # 最後の保険：最小フォントで強制表示（途中切れOK）
    size = min_font
    line_gap = max(2, int(size * 0.3))
    c.setFont(font, size)
    lines = wrap_text(text, font, size, max_width)[:max_lines]
    y = base_y
    for ln in lines:
        c.drawString(base_x, y, ln)
        y -= (size + line_gap)



def fit_font_size(text, font, max_width, max_size=10, min_size=4):
    """
    文字が max_width に収まるフォントサイズを返す
    """
    for size in range(max_size, min_size - 1, -1):
        w = stringWidth(text, font, size)
        if w <= max_width:
            return size
    return min_size
    
def wrap_text(text, font, size, max_width):
    if " " in text:
        units = text.split(" ")
    else:
        units = list(text)

    lines = []
    current = ""

    for u in units:
        test = (current + " " + u).strip() if " " in text else (current + u)
        if stringWidth(test, font, size) <= max_width:
            current = test
        else:
            lines.append(current)
            current = u

    if current:
        lines.append(current)

    return lines




# ===== Excel 読込 ======
def load_sheet_rows(path, sheet):
    wb = load_workbook(str(path), data_only=True)
    ws = wb[sheet]
    rows = []
    for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
        a, b, c = row
        if a is None and (not b) and (not c):
            continue
        try:
            num = int(float(a))
        except:
            num = None
        rows.append({
            "num": num,
            "q": "" if b is None else str(b),
            "a": "" if c is None else str(c)
        })
    return rows

# ===== 40問抽出 ======
def pick40(rows, start, end):
    r = [x for x in rows if x["num"] is not None and start <= x["num"] <= end]
    random.shuffle(r)
    r = r[:40]
    while len(r) < 40:
        r.append({"num": None, "q": "", "a": ""})
    for i, rr in enumerate(r):
        rr["no"] = i + 1
    return r

# ===== 1つのPDFに「問題→解答」2ページ作成 ======
def make_two_page_pdf(items, sheet, start, end):
    filename = TMPDIR / f"{uuid.uuid4().hex}_final.pdf"
    c = canvas.Canvas(str(filename), pagesize=landscape(A4))
    PW, PH = landscape(A4)

    margin = 15*mm
    col_gap = 15*mm
    usable_w = PW - margin*2
    col_w = (usable_w - col_gap)/2

    left_x = margin
    right_x = left_x + col_w + col_gap
    # ====== ページ描画 ======
    def draw_page(mode_label):
        title_y  = PH - 10*mm
        words_y  = title_y - 10*mm
        start_y  = words_y - 14*mm

        c.setFont(DEFAULT_FONT, 16)
        c.drawString(left_x, title_y, "shingaku19minato test")
        
        c.setFont(DEFAULT_FONT, 12)
        c.drawString(left_x, words_y, f"words  {sheet}（{start}～{end}）")
        
        # ←★ これを忘れずに入れる
        c.setFont(DEFAULT_FONT, 12)
        c.drawString(PW - margin - 170, title_y, "name：________________")
        c.drawString(PW - margin - 170, title_y - 8*mm, "score：________________")

        rows_per_col = 20
        bottom = 12*mm
        avail_h = start_y - bottom
        line_h = avail_h / rows_per_col
        if line_h > 12*mm: line_h = 12*mm
        if line_h < 9*mm:  line_h = 9*mm


        # ===== 20行の表を2列に描く =====
        def draw_col(base_x, idx0):
            for i in range(rows_per_col):
                if idx0+i >= len(items): break
                r = items[idx0+i]
        
                y = start_y - i * line_h
        
                # 番号
                c.setFont(DEFAULT_FONT, 10)
                c.drawString(base_x, y, f"{r['no']}.")
        
                # ▼ 幅設定（安全マージン）
                question_width = col_w * 0.50    # 問題の横幅
                answer_width   = col_w * 0.40    # 解答の横幅
                margin_between = col_w * 0.10    # 問題〜解答の間隔
        
                qx = base_x + 10*mm
        
                # ▼ 高さを3行分確保
                max_h = line_h * 3.2
        
                # ▼ 問題
                draw_text_fitted(
                    c, r['q'], DEFAULT_FONT,
                    qx, y,
                    question_width,
                    max_h
                )
        
                if mode_label == "q":
                    lx1 = qx + question_width + 2*mm
                    lx2 = base_x + col_w - 5*mm
                    c.setLineWidth(0.5)
                    c.line(lx1, y - 3, lx2, y - 3)
                else:
                    # ▼ 解答（右に寄せる）
                    ax = base_x + question_width + margin_between
                    # 修正: draw_text_fitted から draw_answer_fitted に変更
                    draw_answer_fitted( 
                        c, r['a'], DEFAULT_FONT,
                        ax, y,
                        answer_width,
                        max_h
                    )



        

        draw_col(left_x, 0)
        draw_col(right_x, 20)

        c.showPage()


    # ===== 1ページ目：問題 =====
    draw_page("q")

    # ===== 2ページ目：解答 =====
    draw_page("a")

    c.save()
    return filename

  
@app.route("/")
def index():
    wb = load_workbook(str(EXCEL_PATH), read_only=True)
    return render_template_string(INDEX_HTML, sheets=wb.sheetnames)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    sheet = data["sheet"]
    start = int(data["start"])
    end   = int(data["end"])

    print(f"📌 読み込むシート: {sheet}, 範囲: {start}-{end}")

    rows = load_sheet_rows(EXCEL_PATH, sheet)
    print(f"📄 取得した行数: {len(rows)}")

    items = pick40(rows, start, end)
    print(f"🧮 抽出した問題数: {len(items)}")

    try:
        final_pdf = make_two_page_pdf(items, sheet, start, end)
        print(f"📦 PDF 出力パス: {final_pdf}")
    except Exception as e:
        print("🚨 PDF 生成中にエラー:", e)
        return jsonify({"error": "PDF作成に失敗しました"}), 500

    if final_pdf is None or not final_pdf.exists():
        print("🚨 PDF が None または存在しません")
        return jsonify({"error": "PDF作成に失敗しました"}), 500

    return jsonify({
        "pdf_url": f"/pdf/{final_pdf.name}"
    })


@app.route("/pdf/<filename>")
def serve_pdf(filename):
    p = TMPDIR / filename
    if not p.exists():
        return "PDF not found", 404
    resp = send_file(str(p), mimetype="application/pdf", as_attachment=False)
    resp.headers["Content-Disposition"] = 'inline; filename="test.pdf"'
    return resp

@app.route("/generate_html_test", methods=["POST"])
def generate_html_test():
    data = request.get_json()
    sheet = data["sheet"]
    start = int(data["start"])
    end = int(data["end"])


    rows = load_sheet_rows(EXCEL_PATH, sheet)
    items = pick40(rows, start, end)


    return render_template_string(
    HTML_TEST_TEMPLATE,
    items=items,
    sheet=sheet,
    start=start,
    end=end
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3710))
    app.run(host="0.0.0.0", port=port)

