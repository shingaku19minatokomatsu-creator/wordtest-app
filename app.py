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

# ====== 設定 ======
EXCEL_PATH = Path("英単語テスト.xlsx")
TMPDIR = Path(gettempdir()) / "word_a_mode"
TMPDIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)

# 日本語フォント
DEFAULT_FONT = "HeiseiMin-W3"
try:
    pdfmetrics.registerFont(UnicodeCIDFont(DEFAULT_FONT))
except Exception:
    DEFAULT_FONT = "Helvetica"

# ===== HTML ======
INDEX_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>単語テスト生成（A方式）</title>
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

<h2>単語テスト PDF 生成（A方式）</h2>
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

    def draw_page(mode_label):
        # header
        title_y = PH - 15*mm
        words_y = title_y - 8*mm
        start_y = words_y - 10*mm

        c.setFont(DEFAULT_FONT, 16)
        c.drawString(left_x, title_y, "shingaku19minato test")

        c.setFont(DEFAULT_FONT, 12)
        c.drawString(left_x, words_y, f"words  {sheet}（{start}～{end}）")

        c.setFont(DEFAULT_FONT, 11)
        c.drawString(PW - margin - 170, title_y, "name：________________")
        c.drawString(PW - margin - 170, title_y - 8*mm, "score：________________")

        # body
        rows_per_col = 20
        bottom = 15*mm
        avail_h = start_y - bottom
        line_h = avail_h / 22
        if line_h > 13*mm: line_h = 13*mm
        if line_h < 8*mm:  line_h = 9*mm

        def draw_col(base_x, idx0):
            for i in range(rows_per_col):
                if idx0+i >= len(items): break
                r = items[idx0+i]
                y = start_y - i*line_h

                c.setFont(DEFAULT_FONT, 11)
                c.drawString(base_x, y, f"{r['no']}.")

                qx = base_x + 10*mm
                c.drawString(qx, y, r['q'])

                if mode_label == "q":
                    # underline
                    lx1 = qx + 45*mm
                    lx2 = base_x + col_w - 5*mm
                    c.setLineWidth(0.5)
                    c.line(lx1, y - 3, lx2, y - 3)
                else:
                    ax = qx + 60*mm
                    c.drawString(ax, y, r['a'])

        draw_col(left_x, 0)
        draw_col(right_x, 20)
        c.showPage()

    # 1ページ目：問題
    draw_page("q")

    # 2ページ目：解答
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

    rows = load_sheet_rows(EXCEL_PATH, sheet)
    items = pick40(rows, start, end)

    # 2ページPDFを生成
    final_pdf = make_two_page_pdf(items, sheet, start, end)

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




