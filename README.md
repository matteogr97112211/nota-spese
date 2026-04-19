# Nota Spese Mobile + scansione scontrino AI

Questa versione mantiene la web app originale e aggiunge la scansione scontrino su ogni riga spesa.

## Cosa serve
- account OpenAI API
- chiave API impostata come variabile ambiente `OPENAI_API_KEY`
- facoltativo: `OPENAI_MODEL` (default `gpt-4.1-mini`)

## Uso locale
```bash
pip install -r requirements.txt
export OPENAI_API_KEY=la_tua_chiave
python app.py
```

## Render
Nelle Environment Variables del servizio aggiungi:
- `OPENAI_API_KEY` = la tua chiave API
- `OPENAI_MODEL` = `gpt-4.1-mini` (oppure un altro modello compatibile con immagini)

## Come funziona
- Apri una riga spesa
- Premi `📸 Scansiona scontrino`
- Scatta o carica la foto
- L'app prova a compilare data, importo, causale e tipo spesa
- Controlla i campi e correggi se serve

## Nota importante
La scansione è molto più robusta del vecchio OCR locale, ma non è perfetta: se la foto è storta, mossa o il totale è poco leggibile, conviene controllare sempre il risultato.
