from flask import Flask, request, render_template, send_file
import pytesseract
from PIL import Image
import tempfile, re, io
from openpyxl import Workbook

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/ocr", methods=["POST"])
def ocr():
    file = request.files["file"]
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
        file.save(tmp.name)
        img = Image.open(tmp.name)

    text = pytesseract.image_to_string(img)

    amounts = re.findall(r"\d+[.,]\d{2}", text)
    amount = amounts[-1] if amounts else ""

    dates = re.findall(r"\d{2}/\d{2}/\d{4}", text)
    date = dates[0] if dates else ""

    return {"importo": amount, "data": date, "raw": text}

@app.route("/generate", methods=["POST"])
def generate():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Nota Spese Generata"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="nota_spese.xlsx")

@app.route("/health")
def health():
    return "OK"
