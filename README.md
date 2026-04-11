# Nota Spese Mobile - versione pronta per pubblicazione online

Web app mobile-first per compilare la nota spese dal telefono e generare un file Excel basato sul modello originale.

## Cosa fa
- compilazione dati trasferta
- inserimento spese una per una
- inserimento tragitti per rimborso km
- salvataggio bozza nel browser del telefono
- generazione dell'Excel finale dal modello `.xlsx`

## Uso locale rapido

### 1) Installa le dipendenze
```bash
pip install -r requirements.txt
```

### 2) Avvia la web app
```bash
python app.py
```

### 3) Apri dal browser
Sul PC:
```text
http://127.0.0.1:5000
```

Dal cellulare sulla stessa rete Wi‑Fi:
```text
http://IP-DEL-PC:5000
```

## Pubblicazione online consigliata: Render
Questa cartella è già pronta per Render.

### Passi
1. crea un account su Render
2. carica questo progetto su GitHub
3. su Render scegli **New + > Web Service**
4. collega il repository GitHub
5. Render rileverà già:
   - `requirements.txt`
   - `Procfile`
   - `render.yaml`
6. avvia il deploy
7. a fine deploy avrai un link pubblico tipo:
   - `https://nome-app.onrender.com`

## Pubblicazione online alternativa: Railway
Funziona anche su Railway senza modifiche.

### Passi
1. crea un account su Railway
2. importa il progetto da GitHub
3. Railway installerà le dipendenze da `requirements.txt`
4. imposta come comando di start:
```bash
gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
```
5. completa il deploy e usa il link pubblico generato

## Struttura file deploy
- `app.py` → backend Flask
- `templates/index.html` → interfaccia mobile
- `static/style.css` → grafica responsive
- `Nota Spese Modello.xlsx` → modello Excel compilato automaticamente
- `requirements.txt` → dipendenze
- `Procfile` → avvio production
- `render.yaml` → configurazione rapida Render
- `runtime.txt` → versione Python consigliata

## Limiti attuali
- fino a 130 righe spese
- fino a 47 tragitti km
- il campo `Fiera` viene scritto solo nelle righe supportate dal modello originale

## Nota importante
Io posso prepararti il progetto già pronto per la pubblicazione, ma la messa online effettiva richiede un tuo account Render, Railway o simile.
