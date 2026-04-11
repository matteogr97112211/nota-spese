from __future__ import annotations

import io
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from flask import Flask, Response, render_template, request, send_file
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_FILE = BASE_DIR / "Nota Spese Modello.xlsx"

APP_TITLE = "Nota Spese Mobile"
SPESA_TYPES = [
    "Vitto/Alloggio",
    "Varie",
    "Spostamenti",
    "Carburante IT",
    "Cambio € -> Valuta",
    "Cambio Valuta -> €",
    "Prelievo Contante",
]
PAYMENT_TYPES = ["CC-Personale", "Contante", "CC-Visa", "CC-Master", "DKV"]

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024


def sanitize_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return value[:80] or "nota_spese"


def parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        return None


def parse_float(value: str | None) -> float | None:
    if value is None or value == "":
        return None
    value = value.replace("€", "").replace(" ", "").replace(",", ".")
    try:
        return float(value)
    except ValueError:
        return None


def normalize_flag(value: str | None) -> str | None:
    return "X" if value and value.upper() == "X" else None


def clear_cell(cell) -> None:
    cell.value = None


def set_riepilogo_fields(wb, data: dict[str, Any]) -> None:
    ws = wb["RIEPILOGO_RIMBORSO_SPESE"]
    mapping = {
        "H6": data.get("rif_ebs"),
        "H7": data.get("autore"),
        "C10": parse_date(data.get("data_inizio")),
        "C11": parse_date(data.get("data_fine")),
        "C12": data.get("cliente"),
        "C13": data.get("luogo"),
        "C14": data.get("tipo"),
        "C15": data.get("causale_trasferta"),
        "C17": data.get("centro_costo"),
        "C18": data.get("dip_esecutore"),
        "C19": data.get("altro_personale_1"),
        "C20": data.get("altro_personale_2"),
        "C21": data.get("altro_personale_3"),
        "C22": data.get("targa_auto_ebs"),
        "C24": parse_float(data.get("notti_italia")),
        "C25": parse_float(data.get("notti_estero")),
        "E10": parse_float(data.get("euro_anticipati")) or 0,
        "E13": parse_float(data.get("euro_resi")) or 0,
        "E17": parse_float(data.get("valuta_anticipata")) or 0,
        "E23": parse_float(data.get("valuta_resa")) or 0,
        "E25": parse_float(data.get("tasso_cambio")) or 1,
        "E16": data.get("tipo_valuta") or "---",
        "J11": parse_float(data.get("bs_causale_1_importo")) or 0,
        "J12": parse_float(data.get("bs_causale_2_importo")) or 0,
        "J13": parse_float(data.get("bs_causale_3_importo")) or 0,
        "G11": data.get("bs_causale_1") or "Confort Viaggi  rif: ",
        "G12": data.get("bs_causale_2") or "TAXI EMMEPI, ",
        "G13": data.get("bs_causale_3") or "AVIS",
    }
    for cell_ref, value in mapping.items():
        ws[cell_ref] = value


def fill_spese_sheet(wb, spese: list[dict[str, Any]]) -> None:
    ws = wb["RIMBORSO_SPESE"]
    for row in range(11, 141):
        for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
            clear_cell(ws[f"{col}{row}"])
    limit = 130
    for idx, item in enumerate(spese[:limit], start=11):
        ws[f"B{idx}"] = parse_date(item.get("data"))
        ws[f"C{idx}"] = item.get("tipo") or None
        ws[f"D{idx}"] = item.get("causale") or None
        ws[f"E{idx}"] = parse_float(item.get("importo_euro"))
        ws[f"F{idx}"] = parse_float(item.get("importo_valuta"))
        ws[f"G{idx}"] = item.get("pagamento") or None
        ws[f"H{idx}"] = normalize_flag(item.get("con_fattura"))
        ws[f"I{idx}"] = normalize_flag(item.get("non_giustificata"))
        ws[f"J{idx}"] = normalize_flag(item.get("rappresentanza"))
        if idx <= 34:
            ws[f"K{idx}"] = normalize_flag(item.get("fiera"))


def fill_km_sheet(wb, tragitti: list[dict[str, Any]]) -> None:
    ws = wb["RIMBORSO KM"]
    for row in range(10, 57):
        ws[f"B{row}"] = None
        ws[f"C{row}"] = None
        ws[f"G{row}"] = None
    limit = 47
    for idx, item in enumerate(tragitti[:limit], start=10):
        ws[f"B{idx}"] = parse_date(item.get("data"))
        ws[f"C{idx}"] = item.get("causale") or None
        ws[f"G{idx}"] = parse_float(item.get("km"))


def workbook_from_form(form: dict[str, Any]) -> io.BytesIO:
    wb = load_workbook(TEMPLATE_FILE)
    set_riepilogo_fields(wb, form)
    fill_spese_sheet(wb, form.get("spese", []))
    fill_km_sheet(wb, form.get("tragitti", []))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


@app.get("/")
def index() -> str:
    defaults = {
        "tasso_cambio": "1",
        "tipo_valuta": "---",
        "bs_causale_1": "Confort Viaggi  rif: ",
        "bs_causale_2": "TAXI EMMEPI, ",
        "bs_causale_3": "AVIS",
        "spese": [
            {
                "data": "",
                "tipo": "",
                "causale": "",
                "importo_euro": "",
                "importo_valuta": "",
                "pagamento": "",
                "con_fattura": "",
                "non_giustificata": "",
                "rappresentanza": "",
                "fiera": "",
            }
        ],
        "tragitti": [{"data": "", "causale": "", "km": ""}],
    }
    return render_template(
        "index.html",
        app_title=APP_TITLE,
        spesa_types=SPESA_TYPES,
        payment_types=PAYMENT_TYPES,
        data=defaults,
        generated=False,
    )


@app.post("/generate")
def generate() -> Response:
    payload_raw = request.form.get("payload")
    if not payload_raw:
        return Response("Dati mancanti.", status=400)
    try:
        payload = json.loads(payload_raw)
    except json.JSONDecodeError:
        return Response("Formato dati non valido.", status=400)

    workbook = workbook_from_form(payload)
    filename_base = sanitize_filename(
        f"nota_spese_{payload.get('cliente') or 'cliente'}_{payload.get('data_inizio') or datetime.now().date()}"
    )
    return send_file(
        workbook,
        as_attachment=True,
        download_name=f"{filename_base}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/health")
def health():
    return "OK", 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug_mode = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug_mode)
