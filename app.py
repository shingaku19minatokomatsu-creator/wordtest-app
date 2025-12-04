# app.py
# A方式：問題PDFと解答PDFを1つのPDFに結合して test.pdf を表示する

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
from reportlab.pdfbase.pdfmetrics import stringWidth
from PyPDF2 import PdfMerger
import os


# ====== 設定 ======
EXCEL_PATH = Path("英単語テスト.xlsx")
TMPDIR = Path(gettempdir()) / "word_a_mode"
TMPDIR.mkdir(parents=True, exist_ok=True)


app = Flask(__name__)

# ===== 日本語フォント =====
DEFAULT_FONT = "HeiseiMin-W3"
try:
    pdfmetrics.registerFont(UnicodeCIDFont(DEFAULT_FONT))
except Exception:
    DEFAULT_FONT = "Helvetica"


# ===== UI（スマホ対応で大きめ） =====
INDEX_HTML = """
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>単語テスト</title>
<style>
body{
    font-family: Arial, sans-serif;
    max-width:980px;
    margin:18px auto;
    padding:14px;
    font-size:18px;
}
label{display:inline-block; width:180px; font-size:20px;}
input,select,button{
    padding:10px;
    font-size:20px;
}
.row{margin:14px 0;}
button{
    padding:12px 28px;
    font-size:22px;
}
.note{color:#666; font-size:16px;}
</style>
</head>
<body>

<h2 style="font-size:26px;">単語テスト</h2>
<div class="note">※「表示」を押すと test.pdf（2ページ構成）が開きます。</div>

<form id="form" onsubmit="return doGenerate(event)">
  <div class="row">
    <label>単語帳（シート）</label>
    <select id="sheet">
      {% for s in sheets %}
      <option value="{{s}}">{{s}}</option>
      {% endfor %}
    </select>
  </div>

  <div class="row"><label>開始番号</label><input id="start" required></div>
  <div class="row"><label>終了番号</label><input id="end" required></div>

  <div class="row"><button type="submit">表示</button></div>
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
        except Exception:
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


# ===== 解答テキストを強制フィット（絶対重ならない） =====
def draw_answer_fitted(c, text, font, base_x, base_y, max_width, max_height):
    font_size = 11  # 最大
    while font_size >= 5:  # 最小5pt
        line_gap = max(1, int(font_size * 0.2))

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

        total_h = len(lines) * font_size + (len(lines) - 1) * line_gap

        if total_h <= max_height:
            c.setFont(font, font_size)
            y = base_y
            for ln in lines:
                c.drawString(base_x, y, ln)
                y -= (font_size + line_gap)
            return

        font_size -= 1

    c.setFont(font, 5)
    c.drawString(base_x, base_y, text[:20] + "...")


# ===== 単独PDF生成 =====
def make_single_pdf(items, sheet, start, end, mode_label):
    filename = TMPDIR / f"{uuid.uuid4().hex}_{mode_label}.pdf"
    c = canvas.Canvas(str(filename), pagesize=landscape(A4))
    PW, PH = landscape(A4)

    margin = 15*mm
    col_gap = 15*mm
    usable_w = PW - margin*2
    col_w = (usable_w - col_gap)/2

    left_x = margin
    right_x = left_x + col_w + col_gap

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
                lx1 = qx + 45*mm
                lx2 = base_x + col_w - 5*mm
                c.setLineWidth(0.5)
                c.line(lx1, y - 3, lx2, y - 3)
            else:
                ax = qx + 60*mm
                max_w = col_w - (ax - base_x) - 5*mm
                max_h = line_h - 3

                draw_answer_fitted(
                    c,
                    r['a'],
                    DEFAULT_FONT,
                    ax,
                    y,
                    max_w,
                    max_h
                )

    draw_col(left_x, 0)
    draw_col(right_x, 20)

    c.showPage()
    c.save()
    return filename


# ===== PDF結合 =====
def merge_two_pdfs(qpdf, apdf):
    final = TMPDIR / f"{uuid.uuid4().hex}_final.pdf"
    merger = PdfMerger()
    merger.append(str(qpdf))
    merger.append(str(apdf))
    merger.write(str(final))
    merger.close()
    return final


# ===== Routes =====
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

    qpdf = make_single_pdf(items, sheet, start, end, "q")
    apdf = make_single_pdf(items, sheet, start, end, "a")

    final_pdf = merge_two_pdfs(qpdf, apdf)

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

