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

ALLOWED_IP_PREFIX = "192.168.0."


app = Flask(__name__)

@app.before_request
def limit_local_only():
    ip = request.remote_addr
    # LAN（ 192.168.x.x と 10.x.x.x ）のみ許可
    if not (ip.startswith("192.168.") or ip.startswith("10.")):
        return "このネットワークからのみ利用可能です", 403

        
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

app = Flask(__name__)


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
    max-width: 900px;
    margin: 0 auto;
    padding: 18px;
    font-size: 18px;
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
<div class="note">※「表示」を押すと test.pdf（問題→解答）が開きます。</div>

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
    <button type="submit">表示</button>
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

  try {
    const win = window.open("about:blank", "_blank");

    const res = await fetch("/generate", {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify({sheet, start, end})
    });

    if(!res.ok){
      const tx = await res.text();
      alert("エラー: " + tx);
      return false;
    }

    const data = await res.json();
    win.location.href = data.pdf_url;

  }catch(err){
    alert("通信エラー: " + err);
  }

  return false;
}
</script>

</body>
</html>
"""

def draw_text_fitted(c, text, font, base_x, base_y, max_width, max_height):
    """
    3行まで折り返し＋フォント縮小して強制フィット
    """

    if not text:
        return

    # ▼ 最大 → 最小フォントサイズ
    max_font = 11
    min_font = 6

    # 最大3行
    max_lines = 3

    # フォント調整ループ
    for size in range(max_font, min_font - 1, -1):
        line_gap = max(1, int(size * 0.22))

        # 折り返し
        words = text.split(" ")
        lines = []
        current = ""

        for w in words:
            test = (current + " " + w).strip()
            if stringWidth(test, font, size) <= max_width:
                current = test
            else:
                if current:
                    lines.append(current)
                current = w
                # 3行超え→省略
                if len(lines) >= max_lines:
                    current = ""
                    break

        if current and len(lines) < max_lines:
            lines.append(current)

        # 行数が0は表示しない
        if not lines:
            continue

        total_h = len(lines) * size + (len(lines) - 1) * line_gap

        if total_h <= max_height:
            # ✔ 収まる → 描画
            c.setFont(font, size)
            y = base_y
            for ln in lines:
                # 文字描画
                c.drawString(base_x, y, ln)
                y -= (size + line_gap)
            return

    # ❗ 超非常手段 → 最小フォントで省略せず表示
    c.setFont(font, min_font)
    c.drawString(base_x, base_y, text[:50] + "...")



def draw_answer_fitted(c, text, font, base_x, base_y, max_width, max_height):
    """
    1行ぶんの縦スペース(max_height)の中に
    長いテキストを折り返し + 縮小して強制フィットさせる
    """

    # ▼ フィットさせたい最大フォントサイズ
    font_size = 11

    while font_size >= 5:  # 最低 5pt まで縮小
        # 行間はフォントサイズに応じて可変
        line_gap = max(1, int(font_size * 0.2))

        # テキストを折り返し
        words = text.split(" ")
        lines = []
        current = ""
        for w in words:
            tmp = (current + " " + w).strip()
            if stringWidth(tmp, font, font_size) <= max_width:
                current = tmp
            else:
                if current:
                    lines.append(current)
                current = w
        if current:
            lines.append(current)

        # 総高さ計算
        total_h = len(lines) * font_size + (len(lines) - 1) * line_gap

        if total_h <= max_height:
            # ✔ 収まった → 描画して終了
            c.setFont(font, font_size)
            y = base_y
            for ln in lines:
                c.drawString(base_x, y, ln)
                y -= (font_size + line_gap)
            return

        # 収まらない → フォント縮小
        font_size -= 1

    # 非常事態：5ptでも入らない → 省略表示
    c.setFont(font, 5)
    c.drawString(base_x, base_y, text[:20] + "...")


def fit_font_size(text, font, max_width, max_size=11, min_size=8):
    """
    文字が max_width に収まるフォントサイズを返す
    """
    for size in range(max_size, min_size - 1, -1):
        w = stringWidth(text, font, size)
        if w <= max_width:
            return size
    return min_size
    
def wrap_text(text, font, size, max_width):
    """
    指定フォント・サイズで max_width に収まるように折り返す
    """
    words = text.split(" ")
    lines = []
    current = ""

    for w in words:
        test = (current + " " + w).strip()
        if stringWidth(test, font, size) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = w
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
        line_h = avail_h / (rows_per_col - 1)
        if line_h > 12*mm: line_h = 12*mm
        if line_h < 9*mm:  line_h = 9*mm


        # ===== 20行の表を2列に描く =====
        def draw_col(base_x, idx0):
            for i in range(rows_per_col):
                if idx0+i >= len(items): break
                r = items[idx0+i]
        
                # ← これだけで縦位置バッチリ揃う
                y = start_y - i * line_h
        
                # 番号
                c.setFont(DEFAULT_FONT, 11)
                c.drawString(base_x, y, f"{r['no']}.")
        
                # ▼ 問題（折り返し縮小）
                qx = base_x + 10*mm
                max_q_width = col_w - 45*mm
                max_h = line_h - 3
                draw_text_fitted(c, r['q'], DEFAULT_FONT, qx, y, max_q_width, max_h)
        
                if mode_label == "q":
                    lx1 = qx + max_q_width + 2*mm
                    lx2 = base_x + col_w - 5*mm
                    c.setLineWidth(0.5)
                    c.line(lx1, y - 3, lx2, y - 3)
                else:
                    ax = qx + 45*mm
                    max_answer_width = col_w - (ax - base_x) - 5*mm
                    draw_text_fitted(c, r['a'], DEFAULT_FONT, ax, y, max_answer_width, max_h)
        

        draw_col(left_x, 0)
        draw_col(right_x, 20)

        c.showPage()


    # ===== 1ページ目：問題 =====
    draw_page("q")

    # ===== 2ページ目：解答 =====
    draw_page("a")

    c.save()
    return filename


# ===== Routes ======
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

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3710))
    app.run(host="0.0.0.0", port=port)
















