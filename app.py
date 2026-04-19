from __future__ import annotations

import base64
import io
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from flask import Flask, Response, jsonify, render_template, request, send_file
from openai import OpenAI
from openpyxl import load_workbook
from pydantic import BaseModel, Field

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
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4.1-mini")
MAX_RECEIPT_SIZE = 8 * 1024 * 1024

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024


class ReceiptExtraction(BaseModel):
    merchant: Optional[str] = Field(default=None, description="Nome esercente o ragione sociale leggibile sullo scontrino.")
    date_iso: Optional[str] = Field(default=None, description="Data dello scontrino in formato YYYY-MM-DD. Null se non chiara.")
    amount_eur: Optional[float] = Field(default=None, description="Totale finale in euro, se lo scontrino è in EUR.")
    amount_original: Optional[float] = Field(default=None, description="Totale finale nella valuta originale se non EUR oppure se non convertibile con certezza.")
    currency: Optional[str] = Field(default=None, description="Codice valuta ISO, ad esempio EUR, USD, GBP. Null se non chiaro.")
    suggested_type: Optional[str] = Field(default=None, description="Categoria più probabile tra: Vitto/Alloggio, Varie, Spostamenti, Carburante IT, Cambio € -> Valuta, Cambio Valuta -> €, Prelievo Contante.")
    expense_description: Optional[str] = Field(default=None, description="Causale breve e pulita da mettere in nota spese.")
    confidence: Optional[str] = Field(default=None, description="alta, media o bassa.")
    raw_text_summary: Optional[str] = Field(default=None, description="Riassunto molto breve del contenuto letto.")


def get_openai_client() -> OpenAI:
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY non configurata")
    return OpenAI(api_key=api_key)


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


def detect_mimetype(filename: str | None) -> str:
    lower = (filename or "").lower()
    if lower.endswith(".png"):
        return "image/png"
    if lower.endswith(".webp"):
        return "image/webp"
    if lower.endswith(".heic") or lower.endswith(".heif"):
        return "image/heic"
    return "image/jpeg"


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
        "ai_ready": bool(os.environ.get("OPENAI_API_KEY", "").strip()),
    }
    return render_template(
        "index.html",
        app_title=APP_TITLE,
        spesa_types=SPESA_TYPES,
        payment_types=PAYMENT_TYPES,
        data=defaults,
        generated=False,
    )


@app.post("/scan-receipt")
def scan_receipt():
    uploaded = request.files.get("receipt")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Nessuna immagine ricevuta."}), 400

    raw = uploaded.read()
    if not raw:
        return jsonify({"error": "Immagine vuota."}), 400
    if len(raw) > MAX_RECEIPT_SIZE:
        return jsonify({"error": "Immagine troppo pesante. Tieni la foto sotto 8 MB."}), 400

    try:
        client = get_openai_client()
    except RuntimeError as exc:
        return jsonify({"error": str(exc)}), 503

    data_url = f"data:{detect_mimetype(uploaded.filename)};base64,{base64.b64encode(raw).decode('utf-8')}"

    try:
        response = client.responses.parse(
            model=OPENAI_MODEL,
            input=[
                {
                    "role": "system",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "Analizza uno scontrino o giustificativo spese e restituisci solo dati realmente leggibili. "
                                "Se un valore non è chiaro usa null. Il totale deve essere il totale finale pagato, "
                                "non il subtotale. Se trovi EUR compila amount_eur, altrimenti amount_original e currency. "
                                "La data va in formato YYYY-MM-DD. suggested_type deve essere una sola tra: "
                                + ", ".join(SPESA_TYPES)
                                + ". expense_description deve essere breve e pronta da inserire nella nota spese. "
                                "Se sembra ristorante, hotel o bar usa Vitto/Alloggio. Taxi, treno, volo, parcheggio, pedaggio -> Spostamenti. "
                                "Benzina, diesel, stazione servizio italiana -> Carburante IT. Bancomat/ATM -> Prelievo Contante. "
                                "Cambio valuta -> usa la categoria relativa al cambio."
                            ),
                        }
                    ],
                },
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": "Estrai i dati utili per una nota spese da questa immagine."},
                        {"type": "input_image", "image_url": data_url},
                    ],
                },
            ],
            text_format=ReceiptExtraction,
        )
        parsed = response.output_parsed
    except Exception as exc:
        return jsonify({"error": f"Analisi AI fallita: {exc}"}), 502

    result = parsed.model_dump()
    if result.get("currency") == "EUR" and result.get("amount_eur") is None and result.get("amount_original") is not None:
        result["amount_eur"] = result["amount_original"]

    return jsonify(result)


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
def health() -> dict[str, str]:
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug_mode = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug_mode)
